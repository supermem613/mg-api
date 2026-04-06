# Microsoft Graph Skill

**AI-powered Microsoft Graph for Claude Code, GitHub Copilot CLI, and other AI coding agents**

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

---

## What It Does

This skill teaches AI coding agents to interact with Microsoft 365 through the Graph API: email, calendar, Teams, users — 25+ operations, zero app registration, cross-platform. Works as both a skill (agent reads SKILL.md and calls scripts) and an MCP server (4 tools, ~500 tokens).

## Capabilities

### Email
- *"Show me my 10 most recent emails"*
- *"Read the full body of that email from Sarah"*
- *"Send an email to the team about the outage"*
- *"Reply to John's message with 'Thanks, approved'"*
- *"Search my inbox for emails about quarterly review"*
- *"Move that email to the Archive folder"*
- *"Delete the spam message"*
- *"List attachments on that email"*

### Calendar
- *"What meetings do I have today?"*
- *"Get the details for my 2pm meeting"*
- *"Create a meeting with Alice and Bob next Tuesday at 3pm"*
- *"Update the project sync — move it to 4pm"*
- *"Cancel Friday's standup"*
- *"Accept the team lunch invitation"*
- *"Decline the 9am meeting with a note"*
- *"Find a free 30-minute slot with Sarah this week"*

### Teams
- *"List all my Teams"*
- *"Show the channels in the Engineering team"*
- *"Post 'Build passed ✅' to the #releases channel"*
- *"List my recent chats"*
- *"Show messages in my chat with Alice"*
- *"Send a chat message to Bob"*

### Users
- *"What's my profile info?"*
- *"Search for people named 'Martinez' in my org"*
- *"Look up user john@contoso.com"*

## Test Drive

Clone the repo and take it for a spin:

```bash
git clone https://github.com/supermem613/microsoft-graph-skill
cd microsoft-graph-skill && npm install
copilot   # or: claude
```

Try: *"Show me my 5 most recent emails"*

## Install

### Claude Code

```claude
/install supermem613/microsoft-graph-skill
```

### Copilot CLI / Other AI Coding Agents

Copy the skill directory into your project and install dependencies:

**macOS / Linux:**

```bash
git clone https://github.com/supermem613/microsoft-graph-skill /tmp/microsoft-graph-skill
cp -r /tmp/microsoft-graph-skill/.claude/skills/microsoft-graph .claude/skills/
npm install playwright
```

**Windows (PowerShell):**

```powershell
git clone https://github.com/supermem613/microsoft-graph-skill $env:TEMP\microsoft-graph-skill
Copy-Item -Recurse $env:TEMP\microsoft-graph-skill\.claude\skills\microsoft-graph .claude\skills\
npm install playwright
```

The skill is auto-discovered from `.claude/skills/`. Run `/skills list` in Copilot CLI to verify.

### MCP Server

The skill also ships as an MCP server with 4 tools (`graph_auth`, `graph_get`, `graph_post`, `graph_docs`). Configure it in your client:

**VS Code / Cursor (`.vscode/mcp.json`):**

```json
{
  "servers": {
    "microsoft-graph": {
      "command": "node",
      "args": ["/path/to/microsoft-graph-skill/src/mcp/server.js"]
    }
  }
}
```

**Claude Desktop (`claude_desktop_config.json`):**

```json
{
  "mcpServers": {
    "microsoft-graph": {
      "command": "node",
      "args": ["/path/to/microsoft-graph-skill/src/mcp/server.js"]
    }
  }
}
```

## Auth

The agent authenticates automatically when the skill is invoked. Playwright launches a persistent Edge browser context — first run opens Edge for login (one-time), then it's instant and headless. No app registration, no client ID, no tenant config, no secrets.

The auth flow navigates to Outlook and Teams to capture two bearer tokens:
- **Graph token** — for Graph API calls (`graph.microsoft.com`)
- **Outlook token** — for send/reply email and chat operations (`outlook.office.com`)

Credentials persist in `~/.microsoft-graph-skill/auth.json` (works across shell sessions). Use `--login` to force re-login, `--logout` to clear the profile.

## Evals

Evals covering auth, email, calendar, Teams, and user operations:

```claude
Run evals/run-evals.md
```

Results are written to `evals/results/`. See the eval files for the full list and scoring criteria.

> **Safe for your real account.** Evals are non-destructive: test emails are sent to yourself and auto-deleted, calendar events are created and auto-deleted, and Teams messages go to a private sandbox chat (only you as member). All test data uses the `MICROSOFT_GRAPH_SKILL_EVAL_` prefix and is cleaned up automatically.

## Tests

```bash
npm test                    # Static tests (no network)
npm run test:core           # Unit tests for core modules
npm run test:mcp            # MCP server tests
```

## Prerequisites

- **Node.js 18+**
- **Microsoft Edge** (Playwright uses your system Edge)
- `npm install` (one-time, installs Playwright + MCP SDK)

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for development setup, project structure, and how to modify scripts, evals, and reference docs.

## License

[MIT](LICENSE)
