# Token Model

Why mg-api carries Graph, Outlook, and Chat tokens and how the registry routes each verb.

## The Token Model

`mg-api` authenticates once and captures up to three tokens into `~/.mg-api/auth.json` for different audiences. The capability registry declares the right `token` per verb (`graph`, `outlook`, or `chat`), and `mg-api`'s REST module selects the cached token automatically — callers never set headers themselves.

### Graph Token

- **Audience:** `https://graph.microsoft.com`
- **Base:** `https://graph.microsoft.com/v1.0/...`
- **Used by:** Most read verbs — mail list/get/search/move/delete, calendar list/view/get/create/update/delete/accept/decline/find-times, users me/search/get, teams list-joined.

### Outlook Token

- **Audience:** `https://outlook.office.com` (Outlook REST v2.0)
- **Base:** `https://outlook.office.com/api/v2.0/...`
- **Used by:** `mg-api email send`, `mg-api email reply`, `mg-api chats list`, `mg-api chats messages`.
- **Note:** Outlook REST payloads use PascalCase field names — the registry body templates already match.

### Graph Chat Token

- **Audience:** `https://graph.microsoft.com` with `Chat.Read`/`Chat.ReadWrite` scopes.
- **Used by:** `mg-api teams list-channels`, `mg-api teams send-channel-message`, `mg-api chats send`.
- **Fallback:** If only a single Graph token was issued with chat scopes, it is reused as the chat token.

### Token routing rules

| Verb | Token | Base |
|------|-------|------|
| `email list|get|search|move|delete|attachments` | graph | graph |
| `email send|reply` | outlook | outlook |
| `calendar *` | graph | graph |
| `users me|search|get` | graph | graph |
| `teams list-joined` | graph | graph |
| `teams list-channels|send-channel-message` | chat | graph |
| `chats list|messages` | outlook | outlook |
| `chats send` | chat | graph |

> **Why separate tokens?** Teams chat permissions (`Chat.Read`, `Chat.ReadWrite`) are in a different permission set than mail/calendar. Some organizations grant these separately. `mg-api auth login` captures all three audiences in one Playwright session and stores them as distinct tokens so the correct one is always used.

> **Debugging 403 errors:** If a Teams endpoint returns 403 "Insufficient privileges", the most likely cause is that the cached chat token is missing or expired. Run `mg-api auth login --force` to re-capture.

---
