# mg-api Agentic Contract

`mg-api` is the stable Microsoft Graph command surface. The repo includes a bundled skill router, but agents should treat the CLI as the product API and call semantic Graph capability commands instead of composing raw HTTP verbs.

## Source of truth

The capability registry in `src/registry.js` defines every command, option, endpoint, help string, schema entry, example, and output contract. Help and schema must be generated from that registry. Do not hand-write separate command help.

## Output contract

Non-help commands write one JSON object to stdout:

```json
{
  "ok": true,
  "command": "email.list",
  "data": {},
  "error": null,
  "meta": {
    "schemaVersion": "0.1.0",
    "endpoint": "/me/messages?$top=10",
    "method": "GET",
    "base": "graph",
    "token": "graph"
  }
}
```

Failures keep stdout machine-readable:

```json
{
  "ok": false,
  "command": "email.send",
  "data": null,
  "error": {
    "code": "VALIDATION_FAILED",
    "message": "Missing required option --to"
  },
  "meta": {
    "schemaVersion": "0.1.0"
  }
}
```

Progress, browser-login instructions, and remediation details go to stderr. Help text is the only human-oriented stdout output.

## Command model

Commands are Microsoft Graph capability groups with bounded verbs:

```text
mg-api auth login|logout|status
mg-api email list|get|send|reply|search|move|delete|attachments
mg-api calendar list|view|get|create|update|delete|accept|decline|find-times
mg-api teams list-joined|list-channels|send-channel-message
mg-api chats list|messages|send
mg-api users me|search|get
mg-api schema [capability] [verb]
mg-api doctor
mg-api update
```

There is no raw HTTP passthrough. If a Microsoft Graph action is missing, add a semantic verb to the registry with tests.

## Token routing

Each verb declares the token audience it needs in the registry: `graph`, `outlook`, or `chat`. The REST module selects the matching token from `~/.mg-api/auth.json`. Callers never set headers themselves.

| Token | Captured from | Used by |
|-------|---------------|---------|
| `graph` | Graph audience (`https://graph.microsoft.com`) | Most read verbs |
| `outlook` | Outlook REST audience (`https://outlook.office.com`) | `email send|reply`, `chats list|messages` |
| `chat` | Graph with chat scopes | `teams list-channels|send-channel-message`, `chats send` |

## Auth isolation

Only `mg-api auth` may load Playwright. REST capability commands must stay on the built-in REST client path and must not import Playwright directly or transitively.

## Self-update

`mg-api update` is for git-clone installs. It runs `git pull --ff-only`, skips install and build when already current, and runs `npm install --no-audit --no-fund` plus `npm run build` when changes arrive. It still returns the standard JSON envelope on stdout.

## Mutation safety

Mutating commands (`email send|reply|move|delete`, `calendar create|update|delete|accept|decline`, `teams send-channel-message`, `chats send`) must make required inputs explicit and return structured failures. Future mutations should add preview/apply semantics where that materially reduces risk.
