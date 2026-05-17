---
name: mg-api
description: "Use when you need to interact with Microsoft Graph through the agentic mg-api CLI for mail, calendar, Teams chats and channels, users, auth, schema inspection, and other Graph capabilities."
metadata:
  author: "Marcus Markiewicz"
  version: "1.0"
  license: "MIT"
  repo: "https://github.com/supermem613/mg-api"
---

# mg-api

This bundled skill is a thin router for the `mg-api` CLI. Use the CLI for Microsoft Graph work. The CLI is agentic-only: stdout is JSON, progress and remediation go to stderr, and help/schema are generated from the same capability registry.

## Execution sequence

1. Run `mg-api doctor` if setup or auth is uncertain.
2. Use `mg-api schema` to inspect the full machine-readable contract, or `mg-api schema <capability> <verb>` for one command.
3. Run semantic commands such as `mg-api email list --top 10`, `mg-api calendar create --subject Standup --start 2026-01-15T09:00:00 --end 2026-01-15T09:30:00`, or `mg-api teams send-channel-message --team-id ... --channel-id ... --content "Build passed"`.
4. If a capability is not listed in `schema`, do not fall back to raw HTTP. Report the missing capability so a verb can be added.

## Capabilities

| Capability | Purpose | Details |
|------------|---------|---------|
| `auth` | Manage Microsoft Graph browser-session authentication | `mg-api auth --help` |
| `email` | Read and send Outlook mail | `mg-api email --help` |
| `calendar` | Read and manage Outlook calendar events | `mg-api calendar --help` |
| `teams` | Work with Microsoft Teams teams and channels | `mg-api teams --help` |
| `chats` | Work with Teams 1:1 and group chats | `mg-api chats --help` |
| `users` | Look up users and people | `mg-api users --help` |
| `schema` | Inspect the generated Microsoft Graph capability schema | `mg-api schema --help` |
| `doctor` | Run local health checks for the agentic CLI | `mg-api doctor --help` |
| `update` | Self-update this mg-api checkout | `mg-api update --help` |

## References

Load these only when you need deeper Microsoft Graph REST details behind a capability:

| File | Covers |
|------|--------|
| [`references/email/`](references/email/README.md) | Mailbox messages, search, send, reply, move, attachments |
| [`references/calendar/`](references/calendar/README.md) | Events, calendarView, RSVP, find meeting times |
| [`references/teams/`](references/teams/README.md) | Teams, channels, channel messages, plus 1:1 and group chats |
| [`references/users/`](references/users/README.md) | Me, people search, user profiles |
| [`references/api-patterns/`](references/api-patterns/README.md) | Graph vs Outlook REST, OData, token routing, pagination, errors |
