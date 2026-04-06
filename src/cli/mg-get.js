#!/usr/bin/env node
// ============================================================================
// mg-get.js — Authenticated Microsoft Graph / Outlook GET request
// ============================================================================
// Usage:  node mg-get.js "/me/messages?$top=5"
//         node mg-get.js "/me/chats?$top=3"
//
// Auto-routes by endpoint:
//   /me/messages, /me/mailFolders, /me/sendMail, /me/events, /me/calendar*
//     → Outlook token + outlook.office.com (has Mail.ReadWrite, Calendars.ReadWrite)
//   /me/chats, /chats/*
//     → Graph Chat token (has Chat.Read, Chat.ReadWrite)
//   Everything else (/me/joinedTeams, /teams/*, /me/people, /users/*, etc.)
//     → Graph token + graph.microsoft.com
//
// Override with --outlook, --chat, or --graph flags.
// ============================================================================
'use strict';

const { GRAPH_TOKEN, GRAPH_CHAT_TOKEN, OUTLOOK_TOKEN } = require('../core/mg-env');
const { graphFetch } = require('../core/mg-fetch');

const args = process.argv.slice(2);
const forceOutlook = args.includes('--outlook');
const forceChat = args.includes('--chat');
const forceGraph = args.includes('--graph');
const positional = args.filter(a => !a.startsWith('--'));

let endpoint = positional[0];

if (!endpoint) {
  process.stderr.write('ERROR: Missing endpoint.\n');
  process.stderr.write('Usage: node mg-get.js "/me/messages?$top=5"\n');
  process.stderr.write('  Auto-routes: mail/calendar → Outlook, chats → Chat token, rest → Graph\n');
  process.stderr.write('  Override: --outlook | --chat | --graph\n');
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
} else if (forceChat) {
  token = GRAPH_CHAT_TOKEN; base = 'https://graph.microsoft.com/v1.0'; tokenName = 'GRAPH_CHAT_TOKEN';
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
const headers = {
  'Accept': 'application/json',
  'Authorization': `Bearer ${token}`,
};

(async () => {
  try {
    const res = await graphFetch(url, { method: 'GET', headers });
    const body = await res.text();

    if (res.ok) {
      process.stdout.write(body + '\n');
    } else {
      process.stderr.write(`ERROR: HTTP ${res.status} on GET ${url}\n`);
      process.stderr.write(body + '\n');
      process.exit(1);
    }
  } catch (err) {
    process.stderr.write(err.message + '\n');
    process.exit(1);
  }
})();
