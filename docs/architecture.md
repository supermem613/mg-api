## Architecture & Design Decisions

### What This Is

This is a Microsoft Graph CLI with a bundled skill router. The `mg-api` CLI is the product surface: it maps capability verbs to Microsoft Graph and Outlook REST implementation details and returns stable JSON envelopes. Agents may load the bundled `SKILL.md`, but that skill only routes them to the CLI.

### Why a CLI with a Bundled Skill Instead of an MCP Server

- **Agentic command surface** — agents call `mg-api email list`, `mg-api calendar create`, and `mg-api schema`, not raw HTTP helpers.
- **No long-running server** — commands are short-lived and work in normal shells.
- **Cross-platform** — Node.js and Playwright work on Windows, macOS, and Linux.
- **Generated contract** — schema, help, and the skill router come from one registry.
- **Auth isolation** — Playwright is limited to `mg-api auth`; REST commands stay on the lightweight hot path.

### Command Architecture

```text
Agent loads SKILL.md
    |
    v
mg-api <capability> <verb>
    |
    +-- src/registry.js     capability specs, params, examples, endpoints, token+base per verb
    +-- src/renderers.js    generated help, schema, SKILL router
    +-- src/mg-api-core.js  dispatcher, JSON envelopes, doctor, self-update
    +-- src/graph-auth.js   Playwright auth and cached tokens
    +-- src/graph-rest.js   Graph and Outlook REST execution with token routing
    +-- src/graph-fetch.js  retry and diagnostics
    |
    v
graph.microsoft.com / outlook.office.com
```

The CLI owns the implementation. The skill directory contains only the router and lazy-loaded references.

### Auth Architecture

```text
mg-api auth login
    |
    v
src/graph-auth.js
    |
    v
Playwright persistent Edge context
    |
    +-- profile: ~/.mg-api/browser-profile/
    +-- auth:    ~/.mg-api/auth.json
```

A single login flow captures up to three tokens (Graph, Outlook, Chat) from Authorization headers on Outlook Web, Teams, and Office page loads. Only the auth path loads Playwright. Tests assert that REST capability files do not import Playwright.

### Token Routing

Each verb declares `token` (`graph`, `outlook`, `chat`) and `base` (`graph`, `outlook`) in the registry. `src/graph-rest.js` looks up the cached token for that audience and builds the request against the correct base URL. There is no per-request auto-detection — the registry is the source of truth.

| Verb group | Token | Base |
|------------|-------|------|
| `email list|get|search|move|delete|attachments` | graph | graph |
| `email send|reply` | outlook | outlook |
| `calendar *` | graph | graph |
| `users me|search|get` | graph | graph |
| `teams list-joined` | graph | graph |
| `teams list-channels|send-channel-message` | chat | graph |
| `chats list|messages` | outlook | outlook |
| `chats send` | chat | graph |

### Bundled Skill Architecture

```text
Agent loads SKILL.md
    |
    +-- command model
    +-- schema/help routing
    +-- reference file index
    |
    v
Agent calls mg-api schema <capability> <verb>
    |
    v
Agent calls semantic command
```

Reference files provide background REST details only after a capability has been selected. They should not route agents around `mg-api`.

### Design Decisions Log

| Decision | Chosen | Why |
|----------|--------|-----|
| Product form | CLI with bundled skill | Tested CLI owns behavior. Skill is a thin router |
| Primary surface | `mg-api` semantic commands | Capability-oriented and agentic |
| Source of truth | `src/registry.js` | Schema/help/SKILL/tests cannot drift |
| Auth method | Playwright persistent context | No app registration, browser-equivalent access |
| Auth boundary | `mg-api auth` only | Prevents Playwright from entering REST hot path |
| Token routing | Declared per verb in registry | No fragile path-based auto-detection |
| Raw HTTP fallback | Not supported | Missing coverage should become a semantic verb |
| Output | JSON envelope | Stable machine-readable agent contract |
| Self-update | `mg-api update` | Git-clone installs can pull, install, and rebuild in one command |
| MCP server | Removed | The CLI is the single product surface |

### What This CLI Cannot Do

| Capability | Why | Workaround |
|-----------|-----|-----------|
| Raw arbitrary HTTP passthrough | Deliberately excluded to keep the CLI semantic | Add a capability verb |
| OneDrive files | Not yet covered | Add a `files` capability when needed |
| SharePoint REST | Different product surface | Use `sp-api` (sibling project) |
| Bulk batch requests | Not exposed yet | Loop semantic verbs or add a `batch` capability |
| Application-only Graph access | Persistent-context auth is browser-equivalent | Use an app registration + MSAL flow if you need daemon access |
