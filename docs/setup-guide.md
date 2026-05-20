# mg-api Setup Guide

This guide covers installing and authenticating the `mg-api` CLI. The repo also ships a thin skill router that agents can load after the CLI is installed.

---

## Install

```bash
cd <repo-root>
npm install
npm run build
npm link
```

Verify the CLI:

```bash
mg-api doctor
mg-api schema auth login
```

`npm install` installs Playwright, which is used only by `mg-api auth`. `npm run build` validates the CLI, generated skill router, and bin wiring before `npm link` exposes `mg-api` on your PATH.

Update a linked or git-clone install later with:

```bash
mg-api update
```

## Authenticate

```bash
mg-api auth login
```

On first run, Edge may open a visible browser window. Sign in with your Microsoft account. Once login completes, the browser closes and your session is saved to a local profile.

Subsequent runs reuse the saved profile headlessly and refresh tokens silently when possible. The flow navigates Outlook Web, Teams, a Teams chat URL, and an Office page so the browser issues bearer tokens for each audience.

## What Gets Saved

Tokens are saved to `~/.mg-api/auth.json`:

| Field | Description |
|-------|-------------|
| `GRAPH_TOKEN` | Bearer for `graph.microsoft.com` |
| `OUTLOOK_TOKEN` | Bearer for Outlook REST v2.0 (`outlook.office.com`) |
| `GRAPH_CHAT_TOKEN` | Graph bearer with Teams chat or channel-message scopes |
| `GRAPH_SCOPES`, `OUTLOOK_SCOPES`, `GRAPH_CHAT_SCOPES` | Scope arrays for diagnostics |

The auth file is an implementation detail consumed by `mg-api` and its internal helpers. Agents should not read or write it directly.

## Verify

```bash
mg-api auth status
mg-api users me --select displayName,mail
mg-api email list --top 5 --select subject,from,receivedDateTime
```

You should receive JSON envelopes on stdout. For command details, run:

```bash
mg-api --help
mg-api email --help
mg-api schema email list
```

## Login / Logout

Force visible re-login:

```bash
mg-api auth login --force
```

Clear the saved profile and auth file:

```bash
mg-api auth logout
```

The browser profile is stored at `~/.mg-api/browser-profile/`.

## Troubleshooting

### Edge not found

Playwright requires Microsoft Edge. Install Edge from [microsoft.com/edge](https://www.microsoft.com/edge).

### Login loop

Your saved session may have expired. Force a fresh login:

```bash
mg-api auth login --force
```

### HTTP 401

Tokens expired. Re-run `mg-api auth login`.

### HTTP 403 on Teams endpoints

The chat-scoped token did not refresh, or the tenant requires separate consent for `Chat.Read` / `Chat.ReadWrite`. Run `mg-api auth login --force` and re-try.

### HTTP 403 elsewhere

Your browser account does not have permission to the requested Microsoft Graph resource.

### Conditional Access / device compliance

Some tenants block tokens issued from non-compliant devices. Sign in once through a regular browser to satisfy the policy, then re-run `mg-api auth login`.
