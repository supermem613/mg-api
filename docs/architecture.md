# Architecture & Design Decisions

## What This Is

A skill that teaches AI agents (Claude Code, GitHub Copilot, Codex, Cursor) to interact with Microsoft 365 via the Graph API. The project delivers two surfaces from one domain core: a skill (agent reads SKILL.md and calls CLI scripts) and a lean MCP server (4 tools for tool-based clients).

## Why Skill + MCP (Not One or the Other)

| Approach | Pros | Cons |
|----------|------|------|
| Skill only | Zero runtime deps, agent flexibility, cross-platform | Requires shell access, no structured tool schema |
| MCP only | Structured tools, discoverable via MCP protocol | Requires running server, fixed tool signatures |
| **Skill + Lean MCP** | **Best of both: flexible agents use the skill, tool-based clients use MCP** | **Two delivery surfaces to maintain** |

The lean MCP approach keeps the MCP server thin — 4 generic tools (`graph_auth`, `graph_get`, `graph_post`, `graph_docs`) that delegate to the same core modules. No operation-specific tools, no schema duplication.

## Auth Architecture

```
┌──────────────────────────────────────────────────────────┐
│                     Auth Layer                            │
│                                                           │
│  ┌──────────┐  Playwright Persistent  ┌──────────┐       │
│  │ mg-auth  │  Context (msedge)      │ Edge      │       │
│  │ .js      │───────────────────────▶│ Browser   │       │
│  │          │                        │ Profile   │       │
│  └──────────┘                        └──────────┘       │
│       │                                                   │
│       │  1. Navigate to Outlook → intercept Graph token   │
│       │  2. Navigate to Teams  → intercept Graph token    │
│       │  3. Capture best token per audience (most scopes) │
│       │                                                   │
│       │  Profile: ~/.microsoft-graph-skill/               │
│       │           browser-profile/                        │
└──────────────────────────────────────────────────────────┘
              │
              ▼ GRAPH_TOKEN + OUTLOOK_TOKEN
┌──────────────────────────────────────────────────────────┐
│                   Delivery Surfaces                       │
│                                                           │
│  ┌─────────────────┐       ┌──────────────────┐         │
│  │   Skill (CLI)   │       │   MCP Server     │         │
│  │ mg-get, mg-post │       │ 4 tools, stdio   │         │
│  └────────┬────────┘       └────────┬─────────┘         │
│           │                         │                     │
│           └────────┬────────────────┘                     │
│                    ▼                                      │
│         ┌──────────────────┐                              │
│         │   Core Modules   │                              │
│         │ mg-client.js     │                              │
│         │ mg-fetch.js      │                              │
│         │ mg-env.js        │                              │
│         └──────────────────┘                              │
└──────────────────────────────────────────────────────────┘
              │
    ┌─────────┼──────────┐
    ▼                    ▼
┌──────────┐     ┌─────────────┐
│ Graph    │     │ Outlook     │
│ v1.0 API │     │ v2.0 API    │
│ graph.   │     │ outlook.    │
│ microsoft│     │ office.com  │
│ .com     │     │             │
└──────────┘     └─────────────┘
```

## Skill Loading Architecture

```
Agent loads SKILL.md (~2K tokens)
    │
    ├─ Auth setup (one-time)
    ├─ Quick reference (10 common ops)
    ├─ Script usage patterns ($SD convention)
    └─ Reference file index
         │
         └─ On demand: agent loads specific reference file
            ├─ email.md (messages, send, reply, attachments)
            ├─ calendar.md (events, scheduling, RSVP)
            ├─ teams.md (teams, channels, messages, chats)
            ├─ users.md (profiles, people search)
            └─ api-patterns.md (pagination, $filter, batching)
```

## MCP Server Architecture

The MCP server (`src/mcp/server.js`) exposes 4 tools and 5 resources via stdio transport. Total tool description overhead: ~500 tokens.

| Tool | Purpose | Read-only |
|------|---------|-----------|
| `graph_auth` | Authenticate via browser sign-in | No |
| `graph_get` | Read data from any Graph GET endpoint | Yes |
| `graph_post` | Write data via POST/PATCH/DELETE | No |
| `graph_docs` | Load reference documentation by topic | Yes |

Resources mirror the reference files (`graph://docs/email`, `graph://docs/calendar`, etc.) for MCP clients that support resource access.

## Context Window Impact

| Component | Tokens | When Loaded |
|-----------|--------|-------------|
| SKILL.md | ~2,000 | Always (session start) |
| email.md | ~1,500 | On demand |
| calendar.md | ~1,500 | On demand |
| teams.md | ~1,200 | On demand |
| users.md | ~800 | On demand |
| api-patterns.md | ~1,500 | On demand |
| **Total (all loaded)** | **~8,500** | Rare |
| **Typical session** | **~3,500–5,000** | SKILL.md + 1–2 references |

## Scripts

| Script | Location | Purpose |
|--------|----------|---------|
| `mg-auth-cli.js` | `src/cli/` | CLI auth entrypoint — wraps mg-auth.js for shell use |
| `mg-get.js` | `src/cli/` | Graph GET request from command line |
| `mg-post.js` | `src/cli/` | Graph POST/PATCH/DELETE from command line |
| `mg-auth.js` | `src/core/` | Playwright auth — captures dual tokens from Outlook + Teams |
| `mg-client.js` | `src/core/` | Typed API client — 25+ named functions for Graph operations |
| `mg-env.js` | `src/core/` | Auth resolver — env vars > auth.json |
| `mg-fetch.js` | `src/core/` | Fetch with retry (transient errors + 429) and diagnostics |
| `server.js` | `src/mcp/` | MCP server — 4 tools, 5 resources, stdio transport |

## Design Decisions Log

| Decision | Chosen | Why |
|----------|--------|-----|
| Skill vs MCP | Both (Skill + Lean MCP) | Flexible agents use skill directly, tool-based clients use MCP |
| Auth method | Playwright persistent context | No app registration, no IT approval, captures real user tokens |
| Dual tokens | Graph + Outlook tokens | Some operations (send email, reply) require Outlook-audience tokens |
| Token capture | Network request interception | Intercept `Authorization: Bearer` headers from Outlook/Teams page loads |
| Script language | Node.js only | Cross-platform, no shell dependencies |
| API surface | Graph v1.0 + Outlook v2.0 | v1.0 for most ops; Outlook REST for send/reply/chats (token audience) |
| MCP tools | 4 generic tools | `graph_get`/`graph_post` handle any endpoint — no per-operation tools |
| Reference files | Lazy-loaded | Token efficiency (~2K base vs ~8.5K all) |
| Core vs CLI split | Separate directories | Core is testable + reusable by MCP; CLI is thin wrappers for shell |
| Error handling | Enriched diagnostics | `mg-fetch.js` adds hints (401 → re-auth, 429 → throttle, etc.) |
