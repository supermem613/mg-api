# Contributing

Guide for engineers working on the `mg-api` CLI and its bundled skill router.

---

## Quick Start

```bash
git clone https://github.com/supermem613/mg-api
cd mg-api
npm install
npm run build
npm link
npm test
```

**Prerequisites:** Node.js 24+, Microsoft Edge for Playwright auth.

## Project Structure

```text
bin/
  mg-api.js                 # Package bin entrypoint
src/
  registry.js               # Single source of truth for capabilities, schema, help
  renderers.js              # Generated help, schema, and SKILL router renderers
  mg-api-core.js            # Agentic CLI dispatcher and JSON envelopes
  graph-auth.js             # Playwright persistent-context auth
  graph-fetch.js            # Shared fetch with retry and diagnostics
  graph-rest.js             # Graph and Outlook REST execution with token routing
.claude/skills/mg-api/
  SKILL.md                  # Lean agent router generated from registry
  references/               # Lazy-loaded REST background references
docs/
  AGENTIC_CONTRACT.md       # mg-api command and output contract
  architecture.md
  setup-guide.md
  api-coverage.md
  auth-deep-dive.md
tests/
  test-scripts.js           # Built-in auth, REST, and fetch module tests
  test-mg-api.js            # mg-api registry/schema/help/envelope tests
  test-integration.js       # Live Microsoft Graph tests
```

## Architecture

This is a **CLI with a bundled skill**. The `mg-api` CLI is the product surface and owns all Microsoft Graph behavior. The bundled `SKILL.md` is a thin router that tells agents to call semantic `mg-api` commands instead of composing raw HTTP.

The registry in `src/registry.js` is the source of truth for:

- command groups and verbs
- parameters and examples
- endpoint/method/base/token metadata
- `mg-api schema`
- generated `--help`
- generated `SKILL.md` router
- test expectations

Implementation logic belongs under `src/`. The skill directory should stay a lean router plus references.

## Token Routing

Each verb declares `token` (`graph`, `outlook`, or `chat`) and `base` (`graph` or `outlook`). `src/graph-rest.js` picks the matching cached token from `~/.mg-api/auth.json` and builds the request against the right base URL. Do not implement path-based auto-detection. Add new audiences only by extending the registry and the auth capture flow.

## Development Workflow

### Running Tests

```bash
npm run build
npm test
npm run test:integration
```

`npm run build` validates that generated artifacts match the registry and that the `mg-api` bin is wired correctly. `npm run link:local` runs the build and then `npm link` for local CLI development.

Use `mg-api update` from linked or git-clone installs to self-update. It runs `git pull --ff-only`, skips install and build when already current, and otherwise runs `npm install --no-audit --no-fund` plus `npm run build`.

`npm test` must stay fast and offline. It validates the `mg-api` CLI contract and its built-in auth/REST modules.

### Linting the Skill

Run `lint-skill` after changing `SKILL.md` or reference files:

```bash
node C:\Users\marcusm\.copilot\skills\lint-skill\scripts\lint-skill.mjs --findings-only .claude\skills\mg-api
```

The linter should be clean before delivery. New errors or warnings must be fixed before shipping.

### Authenticating for Development

```bash
mg-api auth login
mg-api auth status
```

Use `--force` for interactive re-login and `mg-api auth logout` to clear the saved profile.

## Modifying Capabilities

Add or change commands in `src/registry.js` first. The registry change should drive schema, help, docs, and tests.

Rules:

1. **No raw passthrough.** Add semantic capability verbs instead of exposing arbitrary HTTP.
2. **Generated help.** Do not hand-write command help separate from the registry.
3. **Generated skill router.** `SKILL.md` must match `renderSkillRouter()`.
4. **Auth isolation.** Only `auth` may load Playwright. REST capability commands must not import Playwright directly or transitively.
5. **No environment variables.** Configuration lives in the registry and `~/.mg-api/auth.json`. Do not read `process.env`.
6. **Agentic envelopes.** Non-help commands write one JSON object to stdout. Remediation goes to stderr.
7. **Declare token + base.** Every REST verb must specify the audience and base so the router can pick the right token.

## Modifying Reference Files

References provide Microsoft Graph REST background for agents after they have selected a semantic `mg-api` capability. Prefer examples that start with `mg-api schema <capability> <verb>` or a semantic `mg-api` command.

When adding a new operation:

1. Add a semantic verb to `src/registry.js`.
2. Add or update tests in `tests/test-mg-api.js`.
3. Update `docs/api-coverage.md`.
4. Update references only for background details the schema cannot express.

## Code Style

- `'use strict'` at the top of CommonJS scripts
- Shebang line on executable bins
- JSON stdout for non-help commands
- stderr for progress and remediation
- Comments explain why, not mechanics
