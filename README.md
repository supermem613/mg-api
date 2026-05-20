# mg-api

**An agentic Microsoft Graph CLI with a bundled thin skill for Claude Code, GitHub Copilot CLI, and other coding agents.**

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

---

## What It Does

This repo is first and foremost the `mg-api` CLI. It includes a bundled skill only as a thin router so agents know when and how to call the CLI. Agents use semantic Microsoft Graph capability commands instead of raw HTTP verbs. The CLI owns auth, command routing, generated help, generated schema, JSON envelopes, and the current email/calendar/teams/chats/users implementation.

`mg-api` is agentic by default:

- JSON stdout for non-help commands
- Progress and remediation on stderr
- `mg-api schema` as the machine-readable source of truth
- Help generated from the same capability registry as schema
- No raw HTTP passthrough
- Per-verb token routing (Graph, Outlook REST, Graph chat token) — callers never set headers

## Questions and Tasks It Can Handle

### Auth and Setup

- *"Log in to Microsoft 365 so future commands can use it"*
- *"Am I authenticated, and which tokens were captured?"*
- *"Clear the saved browser profile and auth state"*
- *"Check whether my local `mg-api` install is healthy"*
- *"Update this git-clone install of `mg-api`"*

### Email

- *"List my 10 most recent unread messages with subject, from, and received time"*
- *"Get message AAMkAGI... with the full body and sender"*
- *"Send an email to alice@example.com with subject Hello and body Hi there"*
- *"Reply to message AAMkAGI... with the comment Thanks, approved"*
- *"Search my mailbox for quarterly review"*
- *"Move message AAMkAGI... to the Archive folder"*
- *"Delete message AAMkAGI..."*
- *"List attachments on message AAMkAGI..."*

### Calendar

- *"List my next 20 events ordered by start time"*
- *"Show events on 2026-01-15 in calendar view"*
- *"Get event AAMkAGI... with subject, start, end, attendees"*
- *"Create a 30-minute standup for tomorrow at 09:00 with alice@example.com and bob@example.com"*
- *"Update event AAMkAGI... to change its subject"*
- *"Delete event AAMkAGI..."*
- *"Accept event AAMkAGI... with the comment See you there"*
- *"Find meeting times for alice@example.com and bob@example.com on Wednesday, 30 minutes"*

### Teams and Chats

- *"List the Teams I have joined"*
- *"List channels for team {team-id}"*
- *"Post a message to channel {channel-id} in team {team-id}"*
- *"List my Teams chats"*
- *"Read the last 50 messages in chat {chat-id}"*
- *"Send a message to chat {chat-id}"*

### Users and People

- *"Who am I — show display name, mail, job title, department"*
- *"Search for Martinez in the org directory"*
- *"Get user alice@example.com with display name and job title"*

### Schema, Help, and Agent Routing

- *"What Microsoft Graph capabilities does `mg-api` expose right now?"*
- *"Show the machine-readable schema for `email send`"*
- *"Show generated help for `calendar find-times`"*
- *"Tell me whether a capability is implemented or still planned without falling back to raw HTTP"*

## Current Command Surface

```text
mg-api auth     login | logout | status
mg-api email    list | get | send | reply | search | move | delete | attachments
mg-api calendar list | view | get | create | update | delete | accept | decline | find-times
mg-api teams    list-joined | list-channels | send-channel-message
mg-api chats    list | messages | send
mg-api users    me | search | get
mg-api schema   [capability] [verb]
mg-api doctor
mg-api update
```

Planned capability groups (email folders, calendar instances/occurrences, teams members, chats members, users list/photo, OneDrive files) are exposed in `mg-api schema` so agents can see what is not implemented yet without falling back to raw HTTP.

## Quick Start

```bash
git clone https://github.com/supermem613/mg-api
cd mg-api
npm install
npm run build
npm link
mg-api doctor
```

Authenticate once:

```bash
mg-api auth login
```

Then use semantic commands:

```bash
mg-api users me --select displayName,mail,jobTitle
mg-api email list --top 10 --select subject,from,receivedDateTime --orderby "receivedDateTime desc"
mg-api calendar view --start 2026-01-15T00:00:00Z --end 2026-01-16T00:00:00Z
mg-api schema email send
```

## Bundled Skill

The skill under `.claude/skills/mg-api` is not the product surface. It is a generated router plus lazy-loaded references. Install the CLI first, then install or copy the skill so agents route Microsoft Graph tasks to `mg-api`.

### Claude Code

```claude
/install supermem613/mg-api
```

### Copilot CLI / Other Agents

Copy `.claude/skills/mg-api` into the agent's skill directory and install the package dependencies from this repo. The skill routes agents to `mg-api`.

## Auth

`mg-api auth login` uses Playwright with Microsoft Edge persistent context. First run may open Edge for interactive login. Subsequent runs use the saved browser profile headlessly. The flow visits Outlook Web, Teams, a Teams chat URL, a Teams channel surface, and an Office page so the browser issues bearer tokens for each audience. If the Teams channel probe does not observe `ChannelMessage.Read.All`, login reports that explicitly because generic `Chat.*` scopes are not enough for channel-message ingest.

- Browser profile: `~/.mg-api/browser-profile/`
- Auth file: `~/.mg-api/auth.json`
- Force re-login: `mg-api auth login --force`
- Clear auth: `mg-api auth logout`

No app registration, client ID, tenant config, or secret is required.

## Tests

```bash
npm run build
npm test
npm run test:integration
```

`npm run build` validates generated artifacts and the `mg-api` bin before local linking or publishing. `npm link` and `npm run link:local` are supported for local development. For linked or git-clone installs, `mg-api update` pulls with `git pull --ff-only`, skips install/build when already current, and otherwise runs `npm install --no-audit --no-fund` plus `npm run build`.

`npm test` covers the `mg-api` registry, schema generation, help generation, JSON envelopes, Graph auth/REST internals, no raw fallback, auth isolation, package bin wiring, and SKILL router sync.

`npm run test:integration` is the live Microsoft Graph test suite and requires a cached login.

## Docs

- [`docs/AGENTIC_CONTRACT.md`](docs/AGENTIC_CONTRACT.md) — `mg-api` stdout/stderr, schema, help, and command contract
- [`docs/setup-guide.md`](docs/setup-guide.md) — install and auth
- [`docs/architecture.md`](docs/architecture.md) — registry, CLI, auth, and token-routing architecture
- [`docs/api-coverage.md`](docs/api-coverage.md) — current and planned capability coverage
- [`docs/auth-deep-dive.md`](docs/auth-deep-dive.md) — Playwright persistent-context auth details
- [`CONTRIBUTING.md`](CONTRIBUTING.md) — development workflow

## License

[MIT](LICENSE)
