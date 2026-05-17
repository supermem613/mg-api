# Authentication Deep Dive

A technical deep-dive into how `mg-api auth` uses Playwright persistent context authentication for Microsoft Graph and Outlook REST access.

## How Playwright Persistent Context Auth Works

1. `mg-api auth login` invokes the internal Playwright auth helper.
2. Playwright launches a Chromium-based browser (Microsoft Edge) using `chromium.launchPersistentContext`.
3. The persistent context stores its profile at `~/.mg-api/browser-profile/`.
4. On first run, the browser opens visibly and the user signs into Microsoft 365.
5. The browser profile saves cookies, localStorage, and session state to disk.
6. The auth helper navigates through Outlook Web, Teams, a Teams chat URL, and an Office page so the browser issues bearer tokens against each audience.
7. Tokens are captured from outbound `Authorization` headers, classified by JWT `aud` claim, and written to `~/.mg-api/auth.json`.
8. Semantic `mg-api` commands read that file and pick the right token per verb.

### Why Persistent Context Matters

Unlike regular Playwright contexts that start fresh each time, a persistent context stores its state in a user data directory just like a normal browser profile. This means:

- **Login persists.** SSO tokens, session cookies, and auth state survive across runs.
- **Windows SSO/WAM integration.** If you are signed into Windows with your corp account, Graph auth may complete automatically with no manual login.
- **No app registration.** The browser session has the same permissions as your normal browser.
- **No secrets to manage.** No client IDs, client secrets, or certificates.

### How Windows SSO/WAM Provides Frictionless Auth

On Windows machines joined to Entra ID, the Web Account Manager (WAM) provides Single Sign-On:

1. When you sign into Windows with your corporate account, WAM caches your auth tokens.
2. Edge (built on Chromium) has native WAM integration.
3. When Playwright launches Edge with `channel: 'msedge'`, it uses the real Edge binary.
4. Edge's WAM integration picks up your Windows login automatically.
5. Microsoft 365 recognizes the WAM-issued tokens and grants session bearer tokens without a manual login prompt.

This means on a corp-joined Windows machine, the very first run may complete without any manual login.

## What Gets Captured

Bearer tokens for three audiences are stored in `~/.mg-api/auth.json`:

| Field | Source audience | Used by |
|-------|-----------------|---------|
| `GRAPH_TOKEN` | `https://graph.microsoft.com` | Most read verbs |
| `OUTLOOK_TOKEN` | `https://outlook.office.com` / `outlook.office365.com` | `email send|reply`, `chats list|messages` |
| `GRAPH_CHAT_TOKEN` | Graph audience with `Chat.*` scopes | `teams list-channels|send-channel-message`, `chats send` |
| `GRAPH_SCOPES`, `OUTLOOK_SCOPES`, `GRAPH_CHAT_SCOPES` | JWT `scp` claim | Diagnostics |

Tokens are short-lived (typically 60–90 minutes). Re-run `mg-api auth login` when they expire — the persistent profile usually refreshes silently.

## Profile Storage

The browser profile is stored at:

```
~/.mg-api/browser-profile/
```

This directory contains:

- Cookie database
- localStorage / sessionStorage data
- Cached auth tokens
- Browser state (history, preferences)

### Profile Lifecycle

| Event | What Happens |
|-------|-------------|
| First run | Profile created, Edge opens for login |
| Subsequent runs | Profile reused, headless, tokens captured silently |
| `mg-api auth login --force` | Profile reused but Edge opens visibly for re-login |
| `mg-api auth logout` | Profile directory and `auth.json` deleted entirely |
| Tokens expire | Re-run `mg-api auth login`, the profile usually refreshes silently |

## Security Considerations

### Profile on Disk

- The browser profile is stored in your home directory with standard file permissions.
- It contains the same data as your normal Edge profile: cookies and cached auth tokens.
- Anyone with access to your home directory can read these cookies.
- This is the same security model as using Edge normally.
- The profile does NOT contain your password, only session tokens.

### Best Practices

- Do not log or commit token values.
- Re-authenticate periodically (tokens are short-lived).
- Use `mg-api auth logout` to clear the profile when switching accounts or machines.
- The profile directory is under your home directory, never inside the repo.

### Why Three Tokens

| Audience | Why it is needed |
|----------|------------------|
| Graph | Most Microsoft Graph reads and writes |
| Outlook REST v2.0 | `sendmail` and `reply` keep Outlook-specific behavior, and `/me/chats` reads are gated by Outlook-audience permissions in many tenants |
| Graph (chat scopes) | Some tenants split `Chat.Read` and `Chat.ReadWrite` into a separate consent — capturing them on a chat-bearing page avoids 403s on `/teams/*/channels/*/messages` and `/chats/*/messages` |

If a chat endpoint returns 403 "Insufficient privileges", the most likely cause is that the chat-scoped token did not refresh. Run `mg-api auth login --force` and try again.
