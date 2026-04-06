# Authentication Deep Dive

A technical deep-dive into how Playwright persistent context authentication works for Microsoft Graph API access.

## How Playwright Persistent Context Auth Works

1. Playwright launches Microsoft Edge using `chromium.launchPersistentContext`
2. The persistent context stores its profile at `~/.microsoft-graph-skill/browser-profile/`
3. On first run, the browser opens visibly — user signs into their Microsoft 365 account
4. The browser profile saves all cookies, localStorage, and session state to disk
5. On subsequent runs, the browser launches headlessly and loads the saved profile
6. The auth script navigates to Outlook and Teams, intercepting network requests
7. Bearer tokens are extracted from `Authorization` headers and classified by audience
8. Tokens and scopes are saved to `~/.microsoft-graph-skill/auth.json`

### Why Persistent Context Matters

Unlike regular Playwright contexts that start fresh each time, a persistent context stores its state in a user data directory — just like a normal browser profile. This means:

- **Login persists** — SSO tokens, session cookies, and auth state survive across runs
- **Windows SSO/WAM integration** — if you're logged into Windows with your corp account, Graph auth may work automatically (no manual login at all)
- **No app registration** — the browser session has the same permissions as your normal browser
- **No secrets to manage** — no client IDs, client secrets, or certificates

### How Windows SSO/WAM Provides Frictionless Auth

On Windows machines joined to Azure AD (Entra ID), the Web Account Manager (WAM) provides Single Sign-On:

1. When you sign into Windows with your corporate account, WAM caches your auth tokens
2. Edge (built on Chromium) has native WAM integration
3. When Playwright launches Edge with `channel: 'msedge'`, it uses the real Edge binary (not Chromium)
4. Edge's WAM integration picks up your Windows login automatically
5. Microsoft 365 recognizes the WAM-issued tokens and sets session cookies without a manual login prompt

This means on a corp-joined Windows machine, the very first run may complete without any manual login.

## Dual-Token Strategy

The auth flow captures **two tokens** with different audiences:

### Graph Token (`GRAPH_TOKEN`)

- **Audience:** `https://graph.microsoft.com`
- **Source:** Captured from Teams page load (Teams makes Graph API calls on load)
- **Used for:** Most Graph API operations — email read, calendar CRUD, Teams channels, user profiles
- **Endpoint:** `https://graph.microsoft.com/v1.0/...`

### Outlook Token (`OUTLOOK_TOKEN`)

- **Audience:** `https://outlook.office.com` or `https://outlook.office365.com`
- **Source:** Captured from Outlook page load
- **Used for:** Send email, reply to email, list chats, get chat messages
- **Endpoint:** `https://outlook.office.com/api/v2.0/...`

### Why Two Tokens?

Some Microsoft 365 operations require tokens issued to the specific service audience. The Graph API accepts Graph-audience tokens for most operations, but certain Outlook and Teams operations (particularly send/reply email and chat access) work more reliably with Outlook-audience tokens. The skill captures both and uses the appropriate one for each operation.

### Token Classification

The auth module (`mg-auth.js`) classifies intercepted tokens by decoding the JWT payload and checking the `aud` (audience) claim:

| Audience Contains | Token Type | Variable |
|-------------------|------------|----------|
| `graph.microsoft.com` | Graph | `GRAPH_TOKEN` |
| `outlook.office.com` or `outlook.office365.com` | Outlook | `OUTLOOK_TOKEN` |

When multiple tokens of the same type are captured (Outlook and Teams each make several API calls), the one with the most scopes wins.

## Token Scopes

### Outlook Token Scopes (up to ~74)

The Outlook token captured from `outlook.office.com` typically includes scopes like:
- `Mail.ReadWrite`, `Mail.Send` — email operations
- `Calendars.ReadWrite` — calendar operations
- `Contacts.ReadWrite` — contact access
- `Tasks.ReadWrite` — task management
- `People.Read` — people search
- And many more (the Outlook web app requests a broad scope set)

### Graph Token Scopes (up to ~30)

The Graph token captured from Teams typically includes:
- `Chat.ReadWrite`, `ChatMessage.Send` — chat operations
- `Team.ReadBasic.All`, `Channel.ReadBasic.All` — Teams read access
- `ChannelMessage.Send` — channel messaging
- `User.Read`, `User.ReadBasic.All` — user profiles
- `Files.ReadWrite.All` — OneDrive access
- And other Teams-related scopes

### Which Token for Which Operation

| Operation | Token | Why |
|-----------|-------|-----|
| List/read messages | Graph | Standard Graph endpoint |
| Send email | Outlook | Outlook REST API for reliable delivery |
| Reply to email | Outlook | Outlook REST API for threading |
| Calendar CRUD | Graph | Standard Graph endpoint |
| List Teams/channels | Graph | Standard Graph endpoint |
| Send channel message | Graph | Standard Graph endpoint |
| List/read chats | Outlook | Outlook REST API for chat access |
| Send chat message | Graph | Graph API for chat messaging |
| User profile/search | Graph | Standard Graph endpoint |

## Why Not MSAL Device Code Flow?

The standard approach for CLI tools authenticating to Microsoft Graph is MSAL with device code flow. We don't use it because:

1. **Conditional Access blocks well-known app IDs** — many enterprise tenants have Conditional Access policies that block device code flow or block unregistered/well-known client application IDs. This makes device code auth unreliable in corporate environments.
2. **App registration required** — you need to register an app in Azure AD and configure redirect URIs, API permissions, and admin consent. This is a barrier for quick adoption.
3. **Scope limitations** — app registrations must explicitly request each scope. The browser-based approach gets whatever scopes the first-party apps (Outlook, Teams) already have.
4. **Admin consent** — many useful scopes require admin consent for custom app registrations, but the first-party Outlook/Teams apps already have these grants.

## Why Not Azure CLI (`az account get-access-token`)?

Azure CLI is another common auth approach. We don't use it because:

1. **Limited scopes** — Azure CLI tokens are scoped to Azure Resource Manager, not Microsoft Graph email/calendar/Teams operations.
2. **`--scope` parameter limitations** — even with `--scope https://graph.microsoft.com/.default`, the resulting token only gets the scopes consented to the Azure CLI app registration, which doesn't include Mail.Send, Calendars.ReadWrite, Chat.ReadWrite, etc.
3. **Separate install** — requires Azure CLI to be installed, adding a dependency.

## Profile Storage

The browser profile and auth data are stored under the user's home directory:

```
~/.microsoft-graph-skill/
  browser-profile/          # Playwright persistent context (Edge profile data)
    Default/
      Cookies               # Browser cookies database
      Local Storage/        # localStorage data
      Session Storage/      # sessionStorage data
      ...                   # Other browser state
  auth.json                 # Captured tokens and scopes
```

### auth.json Format

```json
{
  "GRAPH_TOKEN": "eyJ0eX...",
  "OUTLOOK_TOKEN": "eyJ0eX...",
  "GRAPH_SCOPES": ["User.Read", "Mail.ReadWrite", "..."],
  "OUTLOOK_SCOPES": ["Mail.ReadWrite", "Mail.Send", "..."]
}
```

### Profile Lifecycle

| Event | What Happens |
|-------|-------------|
| First run | Profile created, Edge opens for login, tokens captured |
| Subsequent runs | Profile reused, headless, tokens captured instantly |
| `--login` | Profile reused but Edge opens visibly for re-login |
| `--logout` | Profile directory + auth.json deleted entirely |
| Tokens expire | Auth script captures fresh tokens (profile still has valid cookies) |
| Cookies expire | Auth script detects login redirect, falls back to visible login |

### Token Lifetime

- **Bearer tokens:** ~1 hour (standard Microsoft identity platform token lifetime)
- **Browser cookies:** ~8–24 hours (varies by tenant policy)
- **Browser profile:** Indefinite (until `--logout` or manual deletion)

Re-running the auth script captures fresh tokens in ~5–10 seconds (headless Outlook + Teams page loads).

## Security Considerations

### Profile on Disk

- The browser profile is stored in your home directory with standard file permissions
- It contains the same data as your normal Edge profile — cookies, cached auth tokens, etc.
- Anyone with access to your home directory can read these tokens
- This is the same security model as using Edge normally
- The profile does NOT contain your password — only session tokens and cookies

### auth.json on Disk

- Contains raw bearer tokens (JWTs) — these grant API access until they expire
- Tokens are short-lived (~1 hour) but should still be protected
- The file is in your home directory, not in the repo (safe from accidental commits)

### Best Practices

- Don't log or commit token values
- Re-authenticate periodically (tokens expire after ~1 hour)
- Use `--logout` to clear the profile when switching accounts or machines
- The profile directory is excluded from version control by default (it's in your home directory, not the repo)
- Env vars (`GRAPH_TOKEN`, `OUTLOOK_TOKEN`) override `auth.json` if set — useful for CI/CD or ephemeral environments
