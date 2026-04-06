# Contributing

Guide for engineers working on the Microsoft Graph Skill.

---

## Quick Start

```bash
git clone https://github.com/supermem613/microsoft-graph-skill
cd microsoft-graph-skill
npm install
```

**Prerequisites:** Node.js 18+, Microsoft Edge (for Playwright auth).

## Project Structure

```
src/
  core/
    mg-auth.js            # Playwright persistent-context auth (dual-token capture)
    mg-client.js          # Typed Graph API client (email, calendar, Teams, users)
    mg-env.js             # Shared auth loader (reads auth.json or env vars)
    mg-fetch.js           # Shared fetch with retry and diagnostic errors
  cli/
    mg-auth-cli.js        # CLI auth entrypoint
    mg-get.js             # Graph GET request (CLI)
    mg-post.js            # Graph POST/PATCH/DELETE (CLI)
  mcp/
    server.js             # MCP server (4 tools: graph_auth, graph_get, graph_post, graph_docs)
references/               # Domain-specific API docs (lazy-loaded by agent or MCP graph_docs)
  email.md                # Messages, send, reply, attachments, folders
  calendar.md             # Events, calendar views, scheduling, reminders
  teams.md                # Teams, channels, messages, chats, members
  users.md                # User profiles, people search, org hierarchy
  api-patterns.md         # Pagination, $select/$filter/$orderby, batching, error handling
.github/skills/microsoft-graph/
  SKILL.md                # GitHub Copilot skill definition
.claude/skills/microsoft-graph/
  SKILL.md                # Claude Code skill definition
docs/                     # Human-facing documentation
evals/
  results/                # Eval output
tests/
  test-scripts.js         # Static validation (no network)
  test-core.js            # Unit tests for core modules
```

## Architecture

This project is a **Skill + Lean MCP** — one domain core, two delivery surfaces:

1. **Skill** — the agent reads `SKILL.md`, learns the API patterns, and calls `mg-get.js`/`mg-post.js` directly. No runtime server needed.
2. **MCP server** — 4 tools (`graph_auth`, `graph_get`, `graph_post`, `graph_docs`) for clients that prefer tool-based interaction (VS Code, Cursor, Claude Desktop).

Both surfaces share the same core modules (`mg-auth.js`, `mg-client.js`, `mg-fetch.js`, `mg-env.js`).

**Auth flow:** Playwright launches Edge with a persistent browser profile (`~/.microsoft-graph-skill/browser-profile/`). It navigates to Outlook and Teams, intercepting network requests to capture Graph and Outlook bearer tokens. On first run, the user logs in visually. After that, auth is headless and instant. No app registration, no client IDs, no secrets.

**Token budget:** `SKILL.md` is kept small (~2K tokens). The 5 reference files in `references/` are loaded on demand by the agent only when needed, keeping the base cost low.

See [`docs/architecture.md`](docs/architecture.md) for diagrams and design decisions.

## Development Workflow

### Running Tests

```bash
npm test                    # Static validation — no network, no auth
npm run test:core           # Unit tests for core modules
npm run test:mcp            # MCP server tests
```

Static tests (`test-scripts.js`) validate file existence, shebangs, error messages, module structure, and that CLI scripts only require local modules (no npm deps in mg-get/mg-post).

### Running Evals

Evals test Graph API operations against your live Microsoft 365 account. They are defined in `evals/` as agent-executable specs — tell your AI agent:

```
Run evals/run-evals.md
```

Results are written to `evals/results/`.

> **Safe for your real account.** Evals are non-destructive — test emails go to yourself (auto-deleted), calendar events are auto-deleted, and Teams messages go to a private sandbox chat (only you). All test data is prefixed with `MICROSOFT_GRAPH_SKILL_EVAL_`.

### Authenticating for Development

```bash
npm run auth               # Standard auth (headless if profile exists)
npm run login              # Force visible browser for re-login
npm run logout             # Clear saved profile + tokens
```

First run opens Edge for login. After that, re-running the command captures fresh tokens headlessly. Tokens last ~1 hour. Use `--login` to force interactive login, `--logout` to clear the profile.

## Modifying Scripts

### CLI Scripts (`src/cli/`)

1. **No npm dependencies** in mg-get, mg-post. Only local `../core/` requires and Node built-ins. The test suite enforces this.
2. **mg-fetch.js** is the shared fetch layer. All HTTP calls go through `graphFetch()`, which handles retry on transient errors (ETIMEDOUT, ECONNRESET, 429) and produces diagnostic error messages. Don't bypass it.
3. **mg-env.js** is the auth resolver. It checks env vars first (`GRAPH_TOKEN`, `OUTLOOK_TOKEN`), then falls back to `~/.microsoft-graph-skill/auth.json`. Don't duplicate this logic.
4. **Cross-platform.** No shell dependencies, no OS-specific paths in scripts. Node.js only.

### Core Modules (`src/core/`)

- **mg-auth.js** — Playwright auth flow. Captures Graph + Outlook tokens by intercepting network requests from Outlook and Teams.
- **mg-client.js** — Typed API client with named functions for every operation. The MCP server and tests use this; CLI scripts use mg-get/mg-post.
- **mg-fetch.js** — Shared fetch with retry and enriched error diagnostics.
- **mg-env.js** — Auth resolver. Env vars override file-based auth.

### MCP Server (`src/mcp/`)

The MCP server exposes 4 tools and 5 resources. It lazy-loads core modules so auth is read fresh each call. When adding new Graph API operations, add them to `mg-client.js` — the MCP server's `graph_get`/`graph_post` tools are generic and don't need changes for new endpoints.

## Modifying SKILL.md

`SKILL.md` is what the agent reads. It's the most sensitive file in the repo — small changes affect every agent interaction.

- Keep it under ~2K tokens. Move detailed API docs to `references/`.
- Use `$SD` (not `$SKILL_DIR`) in all command examples. The agent sets `SD` once per session pointing to the scripts directory.
- Test changes by invoking the skill from Claude Code or Copilot CLI and observing agent behavior.
- The `.github/skills/` and `.claude/skills/` copies must stay in sync.

## Modifying Reference Files

The 5 files in `references/` are loaded on demand when the agent needs domain-specific knowledge (via skill reference loading or the MCP `graph_docs` tool).

| File | Covers |
|------|--------|
| `email.md` | Messages, send, reply, attachments, folders |
| `calendar.md` | Events, calendar views, scheduling, reminders |
| `teams.md` | Teams, channels, messages, chats, members |
| `users.md` | User profiles, people search, org hierarchy |
| `api-patterns.md` | Pagination, $select/$filter/$orderby, batching, error handling |

When adding a new operation: add it to the appropriate reference file, add a row to `docs/api-coverage.md`, and if it's common enough, add it to the quick reference in `SKILL.md`.

## Adding Evals

Evals live in `evals/`. Each eval has:

- A number and name
- The exact command or operation to run
- A pass condition
- Cleanup instructions (if the eval creates data)

All test data must be prefixed with `GRAPH_SKILL_EVAL_` so it's identifiable and cleanable.

## Code Style

- `'use strict'` at the top of every script
- Shebang line (`#!/usr/bin/env node`) on every script
- Errors go to stderr, data goes to stdout
- Silence means success — no verbose output by default
- No comments explaining obvious code. Comment intent, not mechanics.
