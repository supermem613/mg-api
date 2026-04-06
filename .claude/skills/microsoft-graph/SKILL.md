---
name: microsoft-graph
description: "Interact with Microsoft 365 via Graph API — email, calendar, Teams, users, and more."
metadata:
  author: "Marcus Markiewicz"
  version: "1.0"
  license: "MIT"
  repo: "https://github.com/supermem613/microsoft-graph-skill"
---

# Microsoft Graph Skill

Interact with Microsoft 365 services directly via the Graph API. Email, calendar, Teams, users — authenticated HTTP calls, no browser needed after first login.

## Agent Usage (IMPORTANT — read first)

All scripts are in `src/cli/` under this skill's base directory. Set `$SD` once per session.

**Rules:**

1. **Set `SD`** — `SD="<base directory>/src/cli"` (from the `Base directory for this skill:` header above). Use `$SD` for all commands.
2. **Auth persists** via `~/.microsoft-graph-skill/auth.json` — written by `mg-auth-cli.js`, read automatically by `mg-get.js`/`mg-post.js`.
3. **Chain auth + query** with `&&` into a single call to minimize permission asks.
4. **Skip auth** if already authenticated — check `~/.microsoft-graph-skill/auth.json`.

**Single-call pattern** (auth + query combined):
```
SD="<base directory>/src/cli" && node "$SD/mg-auth-cli.js" && node "$SD/mg-get.js" "/me"
```

**Query-only pattern** (when already authenticated):
```
node "$SD/mg-get.js" "/me/messages?\$top=5"
```

## Setup

### Prerequisites

- **Node.js** (18+)
- **Microsoft Edge** — for Playwright persistent context auth
- Run `npm install` in the skill directory (one-time)

### Authenticate

```
node "$SD/mg-auth-cli.js"
```

First run opens Edge for login (one-time). After that, auth is automatic and instant.

### Login / Logout

```
node "$SD/mg-auth-cli.js" --login    # Force re-login
node "$SD/mg-auth-cli.js" --logout   # Clear saved profile + auth
```

## Helper Scripts

| Script | Purpose | Usage |
|--------|---------|-------|
| `mg-auth-cli.js` | Authenticate via Playwright | `node "$SD/mg-auth-cli.js"` |
| `mg-get.js` | Graph GET request | `node "$SD/mg-get.js" "/me/messages?\$top=5"` |
| `mg-post.js` | Graph POST/PATCH/DELETE | `node "$SD/mg-post.js" "/me/events" '{"subject":"Test"}'` |

All scripts auto-load auth from `~/.microsoft-graph-skill/auth.json`. Env vars (`GRAPH_TOKEN`) override the file if set.

## Quick Reference — 10 Most Common Operations

### 1. List recent emails
```
node "$SD/mg-get.js" "/me/messages?\$top=10&\$select=subject,from,receivedDateTime&\$orderby=receivedDateTime desc"
```

### 2. Read a specific message
```
node "$SD/mg-get.js" "/me/messages/{messageId}?\$select=subject,body,from,toRecipients"
```

### 3. Send an email
```
node "$SD/mg-post.js" "/me/sendMail" '{"message":{"subject":"Hello","body":{"contentType":"Text","content":"Hi there"},"toRecipients":[{"emailAddress":{"address":"user@example.com"}}]}}'
```

### 4. List today's calendar events
```
node "$SD/mg-get.js" "/me/calendarView?\$top=20&startDateTime=2025-01-01T00:00:00Z&endDateTime=2025-01-02T00:00:00Z&\$select=subject,start,end,location"
```

### 5. Create a calendar event
```
node "$SD/mg-post.js" "/me/events" '{"subject":"Meeting","start":{"dateTime":"2025-01-15T14:00:00","timeZone":"UTC"},"end":{"dateTime":"2025-01-15T15:00:00","timeZone":"UTC"}}'
```

### 6. List joined Teams
```
node "$SD/mg-get.js" "/me/joinedTeams?\$select=id,displayName"
```

### 7. Send a Teams channel message
```
node "$SD/mg-post.js" "/teams/{teamId}/channels/{channelId}/messages" '{"body":{"content":"Hello from Graph!"}}'
```

### 8. Get current user profile
```
node "$SD/mg-get.js" "/me?\$select=displayName,mail,jobTitle,department"
```

### 9. Search people
```
node "$SD/mg-get.js" "/me/people?\$search=\"John\"&\$top=5&\$select=displayName,emailAddresses,jobTitle"
```

### 10. List chats
```
node "$SD/mg-get.js" "/me/chats?\$top=20&\$select=topic,chatType,lastUpdatedDateTime&\$orderby=lastUpdatedDateTime desc"
```

## API Selection Rule

This skill uses the Microsoft Graph v1.0 API for all operations. All endpoints are under `https://graph.microsoft.com/v1.0/`.

| Use Case | Script |
|----------|--------|
| Email (read, send, search) | `mg-get.js` / `mg-post.js` |
| Calendar (events, scheduling) | `mg-get.js` / `mg-post.js` |
| Teams (teams, channels, messages, chats) | `mg-get.js` / `mg-post.js` |
| Users and people | `mg-get.js` |
| OneDrive files | `mg-get.js` / `mg-post.js` |

## Reference Files

Load these on demand for detailed API documentation:

| File | What It Covers |
|------|---------------|
| [`references/email.md`](references/email.md) | Messages, send, reply, attachments, folders |
| [`references/calendar.md`](references/calendar.md) | Events, calendar views, scheduling, reminders |
| [`references/teams.md`](references/teams.md) | Teams, channels, messages, chats, members |
| [`references/users.md`](references/users.md) | User profiles, people search, org hierarchy |
| [`references/api-patterns.md`](references/api-patterns.md) | Pagination, $select/$filter/$orderby, batching, error handling |

## What This Skill Can't Do (and alternatives)

| Capability | Why Not | Alternative |
|------------|---------|-------------|
| Admin operations (tenant config) | Requires admin consent / app-only tokens | Use Azure portal or admin PowerShell |
| SharePoint site operations | Separate API surface | Use the SharePoint API skill |
| Power Automate / Logic Apps | Workflow services, not REST | Use Power Platform directly |
| Real-time notifications (webhooks) | Requires a public endpoint | Use polling with `mg-get.js` |
| Large file upload (>4 MB) | Requires upload sessions | Compose upload session calls manually |
| Batch requests (JSON batching) | Not yet implemented | Make individual calls |
