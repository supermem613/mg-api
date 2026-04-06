#!/usr/bin/env node
'use strict';

const path = require('path');
const fs = require('fs');

const ROOT = path.resolve(__dirname, '..', '..');
const REFS = path.join(ROOT, 'references');

// Lazy-load core modules so auth file is read fresh each call
function env() { delete require.cache[require.resolve('../core/mg-env')]; return require('../core/mg-env'); }
function client() { return require('../core/mg-client'); }
function auth() { return require('../core/mg-auth'); }

function text(t) { return { content: [{ type: 'text', text: String(t) }] }; }
function json(obj) { return text(JSON.stringify(obj, null, 2)); }

const TOPICS = { email: 'email.md', calendar: 'calendar.md', teams: 'teams.md', users: 'users.md', patterns: 'api-patterns.md' };

async function main() {
  // Dynamic import for ESM-only MCP SDK
  const { McpServer } = await import('@modelcontextprotocol/sdk/server/mcp.js');
  const { StdioServerTransport } = await import('@modelcontextprotocol/sdk/server/stdio.js');
  // Use zod/v3 — the MCP SDK's JSON Schema conversion (toJsonSchemaCompat)
  // fully supports v3 but crashes on v4-full schemas containing z.record().
  const { z } = await import('zod/v3');

  const server = new McpServer({ name: 'microsoft-graph', version: '1.0.0' });

  // ── Tool 1: graph_auth ─────────────────────────────────────────────────────

  server.tool(
    'graph_auth',
    'Authenticate to Microsoft Graph via browser sign-in',
    { login: z.boolean().optional().describe('Force visible browser login'), logout: z.boolean().optional().describe('Clear saved tokens') },
    async ({ login, logout: doLogout }) => {
      if (doLogout) { auth().logout(); return text('Logged out'); }
      const result = await auth().authenticate({ forceLogin: !!login });
      const scopes = [...(result.GRAPH_SCOPES || []), ...(result.OUTLOOK_SCOPES || [])];
      return text(`Authenticated. Scopes: ${scopes.join(', ') || 'unknown'}`);
    },
  );

  // ── Tool 2: graph_get ──────────────────────────────────────────────────────

  server.tool(
    'graph_get',
    'Read data from Microsoft Graph API (any GET endpoint). Auto-selects the correct token for Teams chat endpoints (/me/chats, /teams/).',
    { endpoint: z.string().describe('Graph API path, e.g. /me/messages. Can include query params inline: /me/calendarView?startDateTime=...&endDateTime=...'), params: z.record(z.string()).optional().describe('Query params appended to URL. Keys are passed as-is — include $ for OData params ($filter, $select, $top, $orderby), omit $ for non-OData params (startDateTime, endDateTime).') },
    { readOnlyHint: true, destructiveHint: false },
    async ({ endpoint, params }) => {
      const { GRAPH_TOKEN, GRAPH_CHAT_TOKEN } = env();
      // Use chat token for Teams endpoints if available
      const isTeamsEndpoint = /^\/(me\/chats|teams\/)/.test(endpoint);
      const token = (isTeamsEndpoint && GRAPH_CHAT_TOKEN) ? GRAPH_CHAT_TOKEN : GRAPH_TOKEN;
      if (!token) return text('ERROR: Not authenticated. Run graph_auth first.');
      const { graphFetch, parseGraphError, formatHttpError } = require('../core/mg-fetch');
      // Build query string — pass param keys as-is (caller includes $ for OData params)
      let url = `https://graph.microsoft.com/v1.0${endpoint}`;
      if (params && Object.keys(params).length > 0) {
        const qsParts = Object.entries(params).map(([k, v]) => `${k}=${encodeURIComponent(v)}`);
        const separator = endpoint.includes('?') ? '&' : '?';
        url += separator + qsParts.join('&');
      }
      const res = await graphFetch(url, { headers: { Authorization: `Bearer ${token}`, Accept: 'application/json' } });
      if (!res.ok) { const ge = await parseGraphError(res); return text(formatHttpError(res, ge)); }
      return json(await res.json());
    },
  );

  // ── Tool 3: graph_post ─────────────────────────────────────────────────────

  server.tool(
    'graph_post',
    'Write data to Microsoft Graph API (POST/PATCH/DELETE). Auto-selects the correct token for Teams chat endpoints (/me/chats, /teams/).',
    {
      method: z.enum(['POST', 'PATCH', 'DELETE']).describe('HTTP method'),
      endpoint: z.string().describe('API path, e.g. /me/events'),
      body: z.record(z.unknown()).optional().describe('JSON request body'),
      useOutlookToken: z.boolean().optional().describe('Use Outlook token + outlook.office.com base URL'),
    },
    { readOnlyHint: false },
    async ({ method, endpoint, body, useOutlookToken }) => {
      const { GRAPH_TOKEN, GRAPH_CHAT_TOKEN, OUTLOOK_TOKEN } = env();
      // Use chat token for Teams endpoints if available
      const isTeamsEndpoint = /^\/(me\/chats|teams\/)/.test(endpoint);
      let token;
      if (useOutlookToken) token = OUTLOOK_TOKEN;
      else if (isTeamsEndpoint && GRAPH_CHAT_TOKEN) token = GRAPH_CHAT_TOKEN;
      else token = GRAPH_TOKEN;
      if (!token) return text(`ERROR: No ${useOutlookToken ? 'Outlook' : 'Graph'} token. Run graph_auth first.`);
      const { graphFetch, parseGraphError, formatHttpError } = require('../core/mg-fetch');
      const base = useOutlookToken ? 'https://outlook.office.com/api/v2.0' : 'https://graph.microsoft.com/v1.0';
      const opts = { method, headers: { Authorization: `Bearer ${token}`, Accept: 'application/json', 'Content-Type': 'application/json' } };
      if (body && method !== 'DELETE') opts.body = JSON.stringify(body);
      const res = await graphFetch(`${base}${endpoint}`, opts);
      if (res.status === 204) return text('Success (204 No Content)');
      if (!res.ok) { const ge = await parseGraphError(res); return text(formatHttpError(res, ge)); }
      return json(await res.json());
    },
  );

  // ── Tool 4: graph_docs ─────────────────────────────────────────────────────

  server.tool(
    'graph_docs',
    'Get Graph API reference documentation for a topic',
    { topic: z.enum(['email', 'calendar', 'teams', 'users', 'patterns']).describe('Documentation topic') },
    { readOnlyHint: true },
    async ({ topic }) => {
      const file = path.join(REFS, TOPICS[topic]);
      try { return text(fs.readFileSync(file, 'utf8')); } catch { return text(`ERROR: Reference file not found: ${file}`); }
    },
  );

  // ── Resources ──────────────────────────────────────────────────────────────

  for (const [topic, filename] of Object.entries(TOPICS)) {
    const uri = `graph://docs/${topic}`;
    server.resource(topic, uri, `Graph API ${topic} reference`, async () => ({
      contents: [{ uri, mimeType: 'text/markdown', text: fs.readFileSync(path.join(REFS, filename), 'utf8') }],
    }));
  }

  // ── Start ──────────────────────────────────────────────────────────────────

  const transport = new StdioServerTransport();
  await server.connect(transport);
}

main().catch(err => { process.stderr.write(`MCP server error: ${err.message}\n`); process.exit(1); });
