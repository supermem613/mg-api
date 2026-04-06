#!/usr/bin/env node
// Tests for graph_get/graph_post handler logic:
// 1. Params passed as-is (no auto-$ prefix) — fixes calendarView 400 error
// 2. Teams token routing — auto-selects GRAPH_CHAT_TOKEN for /me/chats and /teams/ endpoints
//
// Run: node --test tests/test-mcp-handler-logic.js

'use strict';

const { describe, it } = require('node:test');
const assert = require('node:assert');

// ============================================================================
// Extract the handler logic from server.js into testable functions.
// These mirror the exact code paths in the graph_get and graph_post handlers.
// ============================================================================

function buildGetUrl(endpoint, params) {
  let url = `https://graph.microsoft.com/v1.0${endpoint}`;
  if (params && Object.keys(params).length > 0) {
    const qsParts = Object.entries(params).map(([k, v]) => `${k}=${encodeURIComponent(v)}`);
    const separator = endpoint.includes('?') ? '&' : '?';
    url += separator + qsParts.join('&');
  }
  return url;
}

function selectGetToken(endpoint, env) {
  const isTeamsEndpoint = /^\/(me\/chats|teams\/)/.test(endpoint);
  return (isTeamsEndpoint && env.GRAPH_CHAT_TOKEN) ? env.GRAPH_CHAT_TOKEN : env.GRAPH_TOKEN;
}

function selectPostToken(endpoint, useOutlookToken, env) {
  const isTeamsEndpoint = /^\/(me\/chats|teams\/)/.test(endpoint);
  if (useOutlookToken) return env.OUTLOOK_TOKEN;
  if (isTeamsEndpoint && env.GRAPH_CHAT_TOKEN) return env.GRAPH_CHAT_TOKEN;
  return env.GRAPH_TOKEN;
}

// ============================================================================
// 1. Params as-is (no auto-$ prefix)
// ============================================================================
describe('graph_get URL construction — params as-is', () => {

  it('no params → clean URL with no query string', () => {
    const url = buildGetUrl('/me/messages', undefined);
    assert.strictEqual(url, 'https://graph.microsoft.com/v1.0/me/messages');
  });

  it('empty params object → clean URL with no query string', () => {
    const url = buildGetUrl('/me/messages', {});
    assert.strictEqual(url, 'https://graph.microsoft.com/v1.0/me/messages');
  });

  it('OData params with $ prefix are passed through unchanged', () => {
    const url = buildGetUrl('/me/messages', { '$top': '10', '$select': 'subject,from' });
    assert.ok(url.includes('$top=10'), 'should contain $top=10');
    assert.ok(url.includes('$select=subject%2Cfrom'), 'should contain $select (comma encoded)');
    assert.ok(url.startsWith('https://graph.microsoft.com/v1.0/me/messages?'), 'should use ? separator');
  });

  it('non-OData params (no $) are passed through unchanged', () => {
    const url = buildGetUrl('/me/calendarView', {
      'startDateTime': '2026-01-06T00:00:00Z',
      'endDateTime': '2026-04-06T23:59:59Z',
    });
    assert.ok(url.includes('startDateTime=2026-01-06'), 'should contain startDateTime without $');
    assert.ok(url.includes('endDateTime=2026-04-06'), 'should contain endDateTime without $');
    assert.ok(!url.includes('$startDateTime'), 'should NOT have $ prefix on startDateTime');
    assert.ok(!url.includes('$endDateTime'), 'should NOT have $ prefix on endDateTime');
  });

  it('mixed OData and non-OData params both work', () => {
    const url = buildGetUrl('/me/calendarView', {
      '$select': 'subject,start,end',
      '$top': '100',
      'startDateTime': '2026-01-06T00:00:00Z',
      'endDateTime': '2026-04-06T23:59:59Z',
    });
    assert.ok(url.includes('$select='), 'should have $select');
    assert.ok(url.includes('$top=100'), 'should have $top');
    assert.ok(url.includes('startDateTime=2026-01-06'), 'should have startDateTime (no $)');
    assert.ok(url.includes('endDateTime=2026-04-06'), 'should have endDateTime (no $)');
  });

  it('$filter with spaces and special chars is properly encoded', () => {
    const url = buildGetUrl('/me/messages', {
      '$filter': "receivedDateTime ge 2026-01-06T00:00:00Z",
    });
    assert.ok(url.includes('$filter=receivedDateTime%20ge%202026-01-06'), 'should encode spaces in $filter');
  });

  it('endpoint already containing ? uses & as separator', () => {
    const url = buildGetUrl('/me/calendarView?startDateTime=2026-01-06T00:00:00Z&endDateTime=2026-04-06T23:59:59Z', {
      '$select': 'subject',
      '$top': '50',
    });
    assert.ok(!url.includes('??'), 'should not have double ?');
    const qsStart = url.indexOf('?');
    const afterQ = url.substring(qsStart + 1);
    // The params should be joined with &, not a second ?
    assert.ok(afterQ.includes('&$select=subject'), 'should append with & after existing ?');
    assert.ok(afterQ.includes('$top=50'), 'should include $top');
  });

  it('endpoint without ? uses ? as separator', () => {
    const url = buildGetUrl('/me/messages', { '$top': '10' });
    assert.match(url, /\/me\/messages\?\$top=10/, 'should use ? separator');
  });

  it('param values are URL-encoded', () => {
    const url = buildGetUrl('/me/messages', {
      '$filter': "subject eq 'Hello World'",
    });
    // encodeURIComponent encodes spaces but not single quotes
    assert.ok(url.includes('Hello%20World'), 'spaces should be encoded');
    assert.ok(url.includes('$filter='), 'should have $filter key as-is');
  });

  it('$orderby param is passed as-is', () => {
    const url = buildGetUrl('/me/messages', {
      '$orderby': 'receivedDateTime desc',
    });
    assert.ok(url.includes('$orderby=receivedDateTime%20desc'), 'should have $orderby with encoded space');
  });
});

// ============================================================================
// 2. Teams token routing — graph_get
// ============================================================================
describe('graph_get Teams token routing', () => {
  const fullEnv = { GRAPH_TOKEN: 'graph-tok', GRAPH_CHAT_TOKEN: 'chat-tok', OUTLOOK_TOKEN: 'outlook-tok' };
  const noChat = { GRAPH_TOKEN: 'graph-tok', GRAPH_CHAT_TOKEN: '', OUTLOOK_TOKEN: 'outlook-tok' };
  const noneSet = { GRAPH_TOKEN: '', GRAPH_CHAT_TOKEN: '', OUTLOOK_TOKEN: '' };

  it('/me/chats uses GRAPH_CHAT_TOKEN when available', () => {
    assert.strictEqual(selectGetToken('/me/chats', fullEnv), 'chat-tok');
  });

  it('/me/chats?$top=50 uses GRAPH_CHAT_TOKEN', () => {
    assert.strictEqual(selectGetToken('/me/chats?$top=50', fullEnv), 'chat-tok');
  });

  it('/me/chats/abc/messages uses GRAPH_CHAT_TOKEN', () => {
    assert.strictEqual(selectGetToken('/me/chats/abc/messages', fullEnv), 'chat-tok');
  });

  it('/teams/abc/channels uses GRAPH_CHAT_TOKEN', () => {
    assert.strictEqual(selectGetToken('/teams/abc/channels', fullEnv), 'chat-tok');
  });

  it('/teams/abc/channels/def/messages uses GRAPH_CHAT_TOKEN', () => {
    assert.strictEqual(selectGetToken('/teams/abc/channels/def/messages', fullEnv), 'chat-tok');
  });

  it('/me/messages uses GRAPH_TOKEN (not a Teams endpoint)', () => {
    assert.strictEqual(selectGetToken('/me/messages', fullEnv), 'graph-tok');
  });

  it('/me/events uses GRAPH_TOKEN', () => {
    assert.strictEqual(selectGetToken('/me/events', fullEnv), 'graph-tok');
  });

  it('/me/calendarView uses GRAPH_TOKEN', () => {
    assert.strictEqual(selectGetToken('/me/calendarView', fullEnv), 'graph-tok');
  });

  it('/me/joinedTeams uses GRAPH_TOKEN (not /teams/)', () => {
    assert.strictEqual(selectGetToken('/me/joinedTeams', fullEnv), 'graph-tok');
  });

  it('/me/chats falls back to GRAPH_TOKEN when GRAPH_CHAT_TOKEN is empty', () => {
    assert.strictEqual(selectGetToken('/me/chats', noChat), 'graph-tok');
  });

  it('/teams/abc falls back to GRAPH_TOKEN when GRAPH_CHAT_TOKEN is empty', () => {
    assert.strictEqual(selectGetToken('/teams/abc', noChat), 'graph-tok');
  });

  it('returns empty string when no tokens are set', () => {
    assert.strictEqual(selectGetToken('/me/messages', noneSet), '');
  });
});

// ============================================================================
// 3. Teams token routing — graph_post
// ============================================================================
describe('graph_post Teams token routing', () => {
  const fullEnv = { GRAPH_TOKEN: 'graph-tok', GRAPH_CHAT_TOKEN: 'chat-tok', OUTLOOK_TOKEN: 'outlook-tok' };
  const noChat = { GRAPH_TOKEN: 'graph-tok', GRAPH_CHAT_TOKEN: '', OUTLOOK_TOKEN: 'outlook-tok' };

  it('/me/chats/abc/messages uses GRAPH_CHAT_TOKEN', () => {
    assert.strictEqual(selectPostToken('/me/chats/abc/messages', false, fullEnv), 'chat-tok');
  });

  it('/teams/abc/channels/def/messages uses GRAPH_CHAT_TOKEN', () => {
    assert.strictEqual(selectPostToken('/teams/abc/channels/def/messages', false, fullEnv), 'chat-tok');
  });

  it('/me/events uses GRAPH_TOKEN', () => {
    assert.strictEqual(selectPostToken('/me/events', false, fullEnv), 'graph-tok');
  });

  it('useOutlookToken=true overrides Teams routing', () => {
    assert.strictEqual(selectPostToken('/me/chats/abc/messages', true, fullEnv), 'outlook-tok');
  });

  it('useOutlookToken=true on non-Teams endpoint uses OUTLOOK_TOKEN', () => {
    assert.strictEqual(selectPostToken('/me/messages', true, fullEnv), 'outlook-tok');
  });

  it('/me/chats falls back to GRAPH_TOKEN when GRAPH_CHAT_TOKEN is empty', () => {
    assert.strictEqual(selectPostToken('/me/chats/abc/messages', false, noChat), 'graph-tok');
  });

  it('non-Teams endpoint without useOutlookToken uses GRAPH_TOKEN', () => {
    assert.strictEqual(selectPostToken('/me/messages', false, fullEnv), 'graph-tok');
  });
});

// ============================================================================
// 4. Source-level verification — server.js handler logic matches expectations
// ============================================================================
describe('server.js source-level handler verification', () => {
  const { readFileSync } = require('node:fs');
  const { join } = require('node:path');
  const src = readFileSync(join(__dirname, '..', 'src', 'mcp', 'server.js'), 'utf8');

  it('graph_get does NOT auto-prefix $ on param keys', () => {
    // The old buggy code had: `$${k}=` which auto-prefixed $.
    // The fix uses: `${k}=` — param keys passed as-is.
    assert.ok(!src.includes('`$${k}='), 'should not contain `$${k}=` (auto-$ prefix)');
    // Verify the correct pattern is present
    assert.ok(src.includes('`${k}=${encodeURIComponent(v)}`'), 'should pass keys as-is');
  });

  it('graph_get handles endpoint with existing ? correctly', () => {
    assert.ok(src.includes("endpoint.includes('?')"), 'should check for existing ? in endpoint');
  });

  it('graph_get reads GRAPH_CHAT_TOKEN from env', () => {
    assert.ok(src.includes('GRAPH_CHAT_TOKEN'), 'should reference GRAPH_CHAT_TOKEN');
  });

  it('graph_get has Teams endpoint regex', () => {
    assert.ok(src.includes('me\\/chats|teams\\/'), 'should match /me/chats and /teams/ endpoints');
  });

  it('graph_post has Teams endpoint regex', () => {
    // Both graph_get and graph_post should have the regex
    const matches = src.match(/isTeamsEndpoint/g);
    assert.ok(matches && matches.length >= 2, 'isTeamsEndpoint should appear in both graph_get and graph_post');
  });

  it('graph_post respects useOutlookToken priority over Teams routing', () => {
    // In the post handler, useOutlookToken should be checked BEFORE isTeamsEndpoint
    const postHandler = src.substring(src.indexOf("'graph_post'"));
    const outlookIdx = postHandler.indexOf('useOutlookToken');
    const teamsIdx = postHandler.indexOf('isTeamsEndpoint && GRAPH_CHAT_TOKEN');
    assert.ok(outlookIdx > 0 && teamsIdx > 0, 'should have both token selection paths');
    assert.ok(outlookIdx < teamsIdx, 'useOutlookToken check should come before Teams check');
  });
});

// ============================================================================
// 5. Teams endpoint regex edge cases
// ============================================================================
describe('Teams endpoint regex matching', () => {
  const regex = /^\/(me\/chats|teams\/)/;

  // Should match
  for (const ep of ['/me/chats', '/me/chats/abc', '/me/chats/abc/messages', '/teams/abc', '/teams/abc/channels', '/teams/abc/channels/def/messages']) {
    it(`matches Teams endpoint: ${ep}`, () => {
      assert.ok(regex.test(ep), `${ep} should match Teams regex`);
    });
  }

  // Should NOT match
  for (const ep of ['/me/messages', '/me/events', '/me/calendarView', '/me/joinedTeams', '/me/teamwork', '/users/abc/chats', '/me/chatMessages']) {
    it(`does NOT match non-Teams endpoint: ${ep}`, () => {
      assert.ok(!regex.test(ep), `${ep} should NOT match Teams regex`);
    });
  }
});
