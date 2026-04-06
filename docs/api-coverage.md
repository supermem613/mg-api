# API Coverage

All supported operations, organized by domain. Each operation maps to a function in `src/core/mg-client.js` and is callable via the CLI scripts or MCP tools.

## Email (8 operations)

| Operation | Function | Method | Endpoint | Token | Reference |
|-----------|----------|--------|----------|-------|-----------|
| List messages | `listMessages` | GET | `/me/messages` | Graph | `references/email.md` |
| Read message | `getMessage` | GET | `/me/messages/{id}` | Graph | `references/email.md` |
| Send email | `sendEmail` | POST | `/me/sendmail` | Outlook | `references/email.md` |
| Reply to message | `replyToMessage` | POST | `/me/messages/{id}/reply` | Outlook | `references/email.md` |
| Search messages | `searchMessages` | GET | `/me/messages?$search=` | Graph | `references/email.md` |
| Move message | `moveMessage` | POST | `/me/messages/{id}/move` | Graph | `references/email.md` |
| Delete message | `deleteMessage` | DELETE | `/me/messages/{id}` | Graph | `references/email.md` |
| List attachments | `listAttachments` | GET | `/me/messages/{id}/attachments` | Graph | `references/email.md` |

## Calendar (8 operations)

| Operation | Function | Method | Endpoint | Token | Reference |
|-----------|----------|--------|----------|-------|-----------|
| List events | `listEvents` | GET | `/me/events` | Graph | `references/calendar.md` |
| Get event | `getEvent` | GET | `/me/events/{id}` | Graph | `references/calendar.md` |
| Create event | `createEvent` | POST | `/me/events` | Graph | `references/calendar.md` |
| Update event | `updateEvent` | PATCH | `/me/events/{id}` | Graph | `references/calendar.md` |
| Delete event | `deleteEvent` | DELETE | `/me/events/{id}` | Graph | `references/calendar.md` |
| Accept event | `acceptEvent` | POST | `/me/events/{id}/accept` | Graph | `references/calendar.md` |
| Decline event | `declineEvent` | POST | `/me/events/{id}/decline` | Graph | `references/calendar.md` |
| Find meeting times | `findMeetingTimes` | POST | `/me/findMeetingTimes` | Graph | `references/calendar.md` |

## Teams (6 operations)

| Operation | Function | Method | Endpoint | Token | Reference |
|-----------|----------|--------|----------|-------|-----------|
| List joined Teams | `listJoinedTeams` | GET | `/me/joinedTeams` | Graph | `references/teams.md` |
| List channels | `listChannels` | GET | `/teams/{id}/channels` | Graph | `references/teams.md` |
| Send channel message | `sendChannelMessage` | POST | `/teams/{id}/channels/{id}/messages` | Graph | `references/teams.md` |
| List chats | `listChats` | GET | `/me/chats` | Outlook | `references/teams.md` |
| Get chat messages | `getChatMessages` | GET | `/me/chats/{id}/messages` | Outlook | `references/teams.md` |
| Send chat message | `sendChatMessage` | POST | `/chats/{id}/messages` | Graph | `references/teams.md` |

## Users (3 operations)

| Operation | Function | Method | Endpoint | Token | Reference |
|-----------|----------|--------|----------|-------|-----------|
| Get my profile | `getMyProfile` | GET | `/me` | Graph | `references/users.md` |
| Search people | `searchPeople` | GET | `/me/people?$search=` | Graph | `references/users.md` |
| Get user | `getUser` | GET | `/users/{id}` | Graph | `references/users.md` |

## Summary

| Domain | Operations | Graph Token | Outlook Token |
|--------|-----------|-------------|---------------|
| Email | 8 | 6 | 2 (send, reply) |
| Calendar | 8 | 8 | 0 |
| Teams | 6 | 4 | 2 (list chats, get chat messages) |
| Users | 3 | 3 | 0 |
| **Total** | **25** | **21** | **4** |

## Not Supported

| Capability | Why | Alternative |
|-----------|-----|-------------|
| Admin operations (tenant config) | Requires admin consent / app-only tokens | Use Azure portal or admin PowerShell |
| SharePoint site operations | Separate API surface | Use the [SharePoint API skill](https://github.com/supermem613/sharepoint-api-skill) |
| Power Automate / Logic Apps | Workflow services, not REST | Use Power Platform directly |
| Real-time notifications (webhooks) | Requires a public endpoint | Use polling with `mg-get.js` |
| Large file upload (>4 MB) | Requires upload sessions | Compose upload session calls manually |
| Batch requests (JSON batching) | Not yet implemented | Make individual calls |
| OneDrive files | Not yet implemented in client | Use `mg-get.js`/`mg-post.js` with raw endpoints |
| Mail folders management | Not yet implemented in client | Use `mg-get.js`/`mg-post.js` with raw endpoints |
