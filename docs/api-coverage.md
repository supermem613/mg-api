# API Coverage

`mg-api schema` is the source of truth for implemented commands. This document summarizes current and planned semantic coverage.

---

## Implemented Capability Groups

| Group | Verbs | Notes |
|-------|-------|-------|
| `auth` | `login`, `logout`, `status` | Playwright Edge persistent-context auth, captures Graph + Outlook + Chat tokens |
| `email` | `list`, `get`, `send`, `reply`, `search`, `move`, `delete`, `attachments` | Graph for reads and lifecycle, Outlook REST for `send` and `reply` |
| `calendar` | `list`, `view`, `get`, `create`, `update`, `delete`, `accept`, `decline`, `find-times` | `view` wraps `/me/calendarView`, `find-times` wraps `/me/findMeetingTimes` |
| `teams` | `list-joined`, `list-channels`, `send-channel-message` | `list-channels` and `send-channel-message` use the chat-scoped token |
| `chats` | `list`, `messages`, `send` | `list` and `messages` use Outlook REST, `send` uses chat-scoped Graph |
| `users` | `me`, `search`, `get` | `search` returns people with `scoredEmailAddresses` |
| `schema` | `show` | Machine-readable contract for all groups and verbs |
| `doctor` | `run` | Local health checks |
| `update` | `run` | Self-update for git-clone installs |

Inspect the live contract:

```bash
mg-api schema
mg-api schema email send
mg-api schema calendar view
mg-api schema teams send-channel-message
```

## Planned Capability Groups

These are exposed as planned groups in `mg-api schema` but are not implemented as commands yet:

| Group | Planned verbs |
|-------|---------------|
| `email` | `folders`, `large-attach`, `reply-all`, `forward` |
| `calendar` | `instances`, `occurrences` |
| `teams` | `members`, `list-team` |
| `chats` | `get`, `members` |
| `users` | `list`, `photo` |
| `files` | `list`, `get`, `upload`, `download` (OneDrive) |

## Adding Coverage

Do not add raw HTTP examples as the public interface. To add coverage:

1. Add a semantic verb in `src/registry.js` with the right `token` and `base`.
2. Include params, endpoint metadata, examples, output docs, and auth requirements.
3. Add command/schema/help tests.
4. Update this coverage summary.

## Unsupported by Design

| Operation | Why | Alternative |
|-----------|-----|-------------|
| Raw REST passthrough | Would bypass the semantic command contract | Add a capability verb |
| SharePoint REST | Different product surface | Use `sp-api` (sibling project) |
| Application-only Graph access | Persistent-context auth is browser-equivalent | Use an app-registration MSAL flow |
| Long-running daemons | The CLI is short-lived per command | Wrap `mg-api` calls in your own orchestrator |
