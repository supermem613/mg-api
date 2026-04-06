#!/usr/bin/env node
// ============================================================================
// mg-post.js — Authenticated Microsoft Graph POST/PATCH/DELETE request
// ============================================================================
// Usage:  node mg-post.js "/me/events" '{"subject":"Test"}'
//         node mg-post.js "/me/events/{id}" '{"subject":"Updated"}' PATCH
//         node mg-post.js "/me/events/{id}" '' DELETE
//         node mg-post.js "/me/sendMail" '{"message":{...}}' --outlook
//
// Optional 3rd argument overrides the HTTP method (PATCH, PUT, DELETE).
// Use --outlook flag to use Outlook token + outlook.office.com base URL
// (needed for Mail.Send, Chat.ReadWrite operations).
//
// Auth: Uses GRAPH_TOKEN or OUTLOOK_TOKEN from ~/.microsoft-graph-skill/auth.json.
// Outputs:  JSON response to stdout
// ============================================================================
'use strict';

const { GRAPH_TOKEN, GRAPH_CHAT_TOKEN, OUTLOOK_TOKEN } = require('../core/mg-env');
const { graphFetch } = require('../core/mg-fetch');

const args = process.argv.slice(2);
const forceOutlook = args.includes('--outlook');
const forceGraph = args.includes('--graph');
const positional = args.filter(a => !a.startsWith('--'));

let endpoint = positional[0];
const body = positional[1] ?? '';
const methodOverride = (positional[2] ?? '').toUpperCase();

if (!endpoint || positional.length < 2) {
  process.stderr.write('ERROR: Missing arguments.\n');
  process.stderr.write('Usage: node mg-post.js "/me/events" \'{"subject":"Test"}\' [PATCH|DELETE]\n');
  process.stderr.write('  Auto-routes: mail/calendar → Outlook, chats → Chat token, rest → Graph\n');
  process.stderr.write('  Override: --outlook | --graph\n');
  process.exit(1);
}

if (!endpoint.startsWith('/')) endpoint = '/' + endpoint;

// Auto-route based on endpoint path
const ep = endpoint.toLowerCase().split('?')[0];
const isMailOrCal = ep.startsWith('/me/messages') || ep.startsWith('/me/mailfolders')
  || ep.startsWith('/me/events') || ep.startsWith('/me/calendar')
  || ep.startsWith('/me/sendmail');
const isChat = ep.startsWith('/me/chats') || ep.startsWith('/chats/');

let token, base, tokenName;
if (forceOutlook) {
  token = OUTLOOK_TOKEN; base = 'https://outlook.office.com/api/v2.0'; tokenName = 'OUTLOOK_TOKEN';
} else if (forceGraph) {
  token = GRAPH_TOKEN; base = 'https://graph.microsoft.com/v1.0'; tokenName = 'GRAPH_TOKEN';
} else if (isMailOrCal && OUTLOOK_TOKEN) {
  token = OUTLOOK_TOKEN; base = 'https://outlook.office.com/api/v2.0'; tokenName = 'OUTLOOK_TOKEN';
} else if (isChat && GRAPH_CHAT_TOKEN) {
  token = GRAPH_CHAT_TOKEN; base = 'https://graph.microsoft.com/v1.0'; tokenName = 'GRAPH_CHAT_TOKEN';
} else {
  token = GRAPH_TOKEN; base = 'https://graph.microsoft.com/v1.0'; tokenName = 'GRAPH_TOKEN';
}

if (!token) {
  process.stderr.write(`ERROR: ${tokenName} is not set. Run: node mg-auth-cli.js --login\n`);
  process.exit(1);
}

const url = `${base}${endpoint}`;
const method = methodOverride || 'POST';
const headers = {
  'Accept': 'application/json',
  'Authorization': `Bearer ${token}`,
};

if (method !== 'DELETE') {
  headers['Content-Type'] = 'application/json';
}

const fetchOpts = { method, headers };
if (body) {
  fetchOpts.body = body;
}

(async () => {
  try {
    const res = await graphFetch(url, fetchOpts);
    const resBody = await res.text();

    if (res.ok) {
      if (resBody) {
        process.stdout.write(resBody + '\n');
      }
    } else {
      process.stderr.write(`ERROR: HTTP ${res.status} on ${method} ${url}\n`);
      process.stderr.write(resBody + '\n');
      process.exit(1);
    }
  } catch (err) {
    process.stderr.write(err.message + '\n');
    process.exit(1);
  }
})();
