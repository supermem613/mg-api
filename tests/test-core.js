#!/usr/bin/env node
'use strict';

const { describe, it, beforeEach, afterEach } = require('node:test');
const assert = require('node:assert/strict');

const client = require('../src/core/mg-client');

const TOKEN = 'test-graph-token';
const OUTLOOK_TOKEN = 'test-outlook-token';
const GRAPH = 'https://graph.microsoft.com/v1.0';
const OUTLOOK = 'https://outlook.office.com/api/v2.0';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/** Capture the last fetch call's url + options and return a configurable response. */
function mockFetch(body = {}, status = 200) {
  const calls = [];
  globalThis.fetch = async (url, opts) => {
    calls.push({ url, opts });
    return {
      ok: status >= 200 && status < 300,
      status,
      statusText: status === 200 ? 'OK' : 'Error',
      url,
      headers: new Map(),
      json: async () => body,
    };
  };
  return calls;
}

/** Mock fetch that fails N times with a network error, then succeeds. */
function mockFetchTransient(failCount, code) {
  let attempt = 0;
  const calls = [];
  globalThis.fetch = async (url, opts) => {
    calls.push({ url, opts });
    if (attempt++ < failCount) {
      const err = new Error(`connect ${code}`);
      err.code = code;
      throw err;
    }
    return {
      ok: true, status: 200, statusText: 'OK', url,
      headers: new Map(),
      json: async () => ({ value: [] }),
    };
  };
  return calls;
}

/** Mock fetch that returns 429 N times, then succeeds. */
function mockFetch429(failCount) {
  let attempt = 0;
  const calls = [];
  globalThis.fetch = async (url, opts) => {
    calls.push({ url, opts });
    if (attempt++ < failCount) {
      return {
        ok: false, status: 429, statusText: 'Too Many Requests', url,
        headers: new Map([['Retry-After', '0']]),
        json: async () => ({ error: { code: 'TooManyRequests', message: 'Throttled' } }),
      };
    }
    return {
      ok: true, status: 200, statusText: 'OK', url,
      headers: new Map(),
      json: async () => ({ value: [] }),
    };
  };
  return calls;
}

let originalFetch;

beforeEach(() => { originalFetch = globalThis.fetch; });
afterEach(() => { globalThis.fetch = originalFetch; });

// ===========================================================================
// Email
// ===========================================================================

describe('Email', () => {
  it('listMessages builds correct URL with query params', async () => {
    const calls = mockFetch({ value: [] });
    await client.listMessages(TOKEN, { top: 5, filter: "isRead eq false", select: 'subject' });
    assert.equal(calls.length, 1);
    const url = calls[0].url;
    assert.ok(url.startsWith(`${GRAPH}/me/messages?`), `URL should start with graph/me/messages, got ${url}`);
    assert.ok(url.includes('$top=5'), 'should contain $top');
    assert.ok(url.includes('$filter='), 'should contain $filter');
    assert.ok(url.includes('$select=subject'), 'should contain $select');
    assert.equal(calls[0].opts.method, undefined, 'GET has no explicit method');
    assert.ok(calls[0].opts.headers.Authorization.includes(TOKEN));
  });

  it('sendEmail uses Outlook endpoint and formats recipients', async () => {
    const calls = mockFetch({});
    await client.sendEmail(OUTLOOK_TOKEN, {
      to: ['alice@example.com', 'bob@example.com'],
      cc: ['carol@example.com'],
      subject: 'Test',
      body: 'Hello',
    });
    assert.equal(calls.length, 1);
    const { url, opts } = calls[0];
    assert.ok(url.startsWith(OUTLOOK), `should use Outlook base, got ${url}`);
    assert.ok(url.includes('/me/sendmail'), 'should hit /me/sendmail');
    assert.equal(opts.method, 'POST');

    const payload = JSON.parse(opts.body);
    assert.equal(payload.Message.ToRecipients.length, 2);
    assert.equal(payload.Message.ToRecipients[0].EmailAddress.Address, 'alice@example.com');
    assert.equal(payload.Message.CcRecipients.length, 1);
    assert.equal(payload.Message.Subject, 'Test');
    assert.equal(payload.SaveToSentItems, true);
  });

  it('searchMessages URL-encodes the search query', async () => {
    const calls = mockFetch({ value: [] });
    await client.searchMessages(TOKEN, { query: 'hello world' });
    assert.equal(calls.length, 1);
    const url = calls[0].url;
    assert.ok(url.includes('$search='), 'should have $search param');
    assert.ok(url.includes(encodeURIComponent('"hello world"')), 'query should be URL-encoded with quotes');
  });

  it('deleteMessage uses DELETE method', async () => {
    const calls = mockFetch(null, 204);
    await client.deleteMessage(TOKEN, { messageId: 'msg-123' });
    assert.equal(calls.length, 1);
    assert.equal(calls[0].opts.method, 'DELETE');
    assert.ok(calls[0].url.includes('/me/messages/msg-123'));
  });

  it('getMessage includes messageId in path', async () => {
    const calls = mockFetch({ subject: 'hi' });
    await client.getMessage(TOKEN, { messageId: 'abc', select: 'subject,body' });
    assert.ok(calls[0].url.includes('/me/messages/abc'));
    assert.ok(calls[0].url.includes('$select='));
  });

  it('replyToMessage POSTs to Outlook reply endpoint', async () => {
    const calls = mockFetch({});
    await client.replyToMessage(OUTLOOK_TOKEN, { messageId: 'msg-1', comment: 'Thanks' });
    assert.ok(calls[0].url.startsWith(OUTLOOK));
    assert.ok(calls[0].url.includes('/me/messages/msg-1/reply'));
    const body = JSON.parse(calls[0].opts.body);
    assert.equal(body.Comment, 'Thanks');
  });
});

// ===========================================================================
// Calendar
// ===========================================================================

describe('Calendar', () => {
  it('createEvent sends correct JSON body with start/end/attendees', async () => {
    const calls = mockFetch({ id: 'evt-1' });
    await client.createEvent(TOKEN, {
      subject: 'Standup',
      body: 'Daily standup',
      start: '2025-01-15T09:00:00',
      end: '2025-01-15T09:30:00',
      attendees: ['alice@example.com'],
      isOnlineMeeting: true,
    });
    assert.equal(calls.length, 1);
    const { url, opts } = calls[0];
    assert.ok(url.includes('/me/events'));
    assert.equal(opts.method, 'POST');

    const body = JSON.parse(opts.body);
    assert.equal(body.subject, 'Standup');
    assert.deepEqual(body.start, { dateTime: '2025-01-15T09:00:00', timeZone: 'UTC' });
    assert.deepEqual(body.end, { dateTime: '2025-01-15T09:30:00', timeZone: 'UTC' });
    assert.equal(body.attendees.length, 1);
    assert.equal(body.attendees[0].emailAddress.address, 'alice@example.com');
    assert.equal(body.attendees[0].type, 'required');
    assert.equal(body.isOnlineMeeting, true);
  });

  it('listEvents with filter builds correct OData query', async () => {
    const calls = mockFetch({ value: [] });
    await client.listEvents(TOKEN, {
      top: 10,
      filter: "start/dateTime ge '2025-01-01'",
      orderby: 'start/dateTime',
    });
    const url = calls[0].url;
    assert.ok(url.startsWith(`${GRAPH}/me/events?`));
    assert.ok(url.includes('$top=10'));
    assert.ok(url.includes('$filter='));
    assert.ok(url.includes('$orderby='));
  });

  it('acceptEvent POSTs to correct endpoint', async () => {
    const calls = mockFetch({});
    await client.acceptEvent(TOKEN, { eventId: 'evt-42', comment: 'Sure!', sendResponse: true });
    assert.equal(calls[0].opts.method, 'POST');
    assert.ok(calls[0].url.includes('/me/events/evt-42/accept'));
    const body = JSON.parse(calls[0].opts.body);
    assert.equal(body.comment, 'Sure!');
    assert.equal(body.sendResponse, true);
  });

  it('declineEvent POSTs to decline endpoint', async () => {
    const calls = mockFetch({});
    await client.declineEvent(TOKEN, { eventId: 'evt-7', comment: 'Busy' });
    assert.ok(calls[0].url.includes('/me/events/evt-7/decline'));
  });

  it('deleteEvent uses DELETE method', async () => {
    const calls = mockFetch(null, 204);
    await client.deleteEvent(TOKEN, { eventId: 'evt-99' });
    assert.equal(calls[0].opts.method, 'DELETE');
    assert.ok(calls[0].url.includes('/me/events/evt-99'));
  });
});

// ===========================================================================
// Teams
// ===========================================================================

describe('Teams', () => {
  it('listJoinedTeams hits correct endpoint', async () => {
    const calls = mockFetch({ value: [] });
    await client.listJoinedTeams(TOKEN);
    assert.equal(calls.length, 1);
    assert.ok(calls[0].url.includes('/me/joinedTeams'));
    assert.ok(calls[0].opts.headers.Authorization.includes(TOKEN));
  });

  it('sendChannelMessage formats body correctly', async () => {
    const calls = mockFetch({ id: 'msg-1' });
    await client.sendChannelMessage(TOKEN, {
      teamId: 'team-1',
      channelId: 'chan-1',
      content: 'Hello team!',
    });
    const { url, opts } = calls[0];
    assert.ok(url.includes('/teams/team-1/channels/chan-1/messages'));
    assert.equal(opts.method, 'POST');
    const body = JSON.parse(opts.body);
    assert.equal(body.body.content, 'Hello team!');
    assert.equal(body.body.contentType, 'text');
  });

  it('listChats uses Outlook base URL', async () => {
    const calls = mockFetch({ value: [] });
    await client.listChats(OUTLOOK_TOKEN, { top: 5 });
    assert.equal(calls.length, 1);
    assert.ok(calls[0].url.startsWith(OUTLOOK), `should use Outlook, got ${calls[0].url}`);
    assert.ok(calls[0].url.includes('/me/chats'));
    assert.ok(calls[0].url.includes('$top=5'));
  });

  it('listChannels hits correct path', async () => {
    const calls = mockFetch({ value: [] });
    await client.listChannels(TOKEN, { teamId: 'team-x' });
    assert.ok(calls[0].url.includes('/teams/team-x/channels'));
  });
});

// ===========================================================================
// Users
// ===========================================================================

describe('Users', () => {
  it('getMyProfile with $select param', async () => {
    const calls = mockFetch({ displayName: 'Alice' });
    await client.getMyProfile(TOKEN, { select: 'displayName,mail' });
    assert.equal(calls.length, 1);
    const url = calls[0].url;
    assert.ok(url.includes('/me?'));
    assert.ok(url.includes('$select='));
    assert.ok(url.includes('displayName'));
  });

  it('searchPeople URL-encodes search', async () => {
    const calls = mockFetch({ value: [] });
    await client.searchPeople(TOKEN, { query: 'John Doe', top: 3 });
    const url = calls[0].url;
    assert.ok(url.includes('/me/people'));
    assert.ok(url.includes('$search='), 'should include $search');
    assert.ok(url.includes(encodeURIComponent('"John Doe"')), 'query should be encoded');
    assert.ok(url.includes('$top=3'));
  });

  it('getUser includes userId in path', async () => {
    const calls = mockFetch({ displayName: 'Bob' });
    await client.getUser(TOKEN, { userId: 'user-42', select: 'displayName' });
    assert.ok(calls[0].url.includes('/users/user-42'));
  });
});

// ===========================================================================
// Error handling
// ===========================================================================

describe('Error handling', () => {
  it('graphFetch retries on 429 then succeeds', async () => {
    const calls = mockFetch429(1);
    const res = await client.listMessages(TOKEN);
    // First call gets 429, second succeeds
    assert.equal(calls.length, 2, 'should retry once after 429');
    assert.equal(res.status, 200);
  });

  it('graphFetch retries on transient network errors', async () => {
    const calls = mockFetchTransient(1, 'ETIMEDOUT');
    const res = await client.listJoinedTeams(TOKEN);
    assert.equal(calls.length, 2, 'should retry after ETIMEDOUT');
    assert.equal(res.status, 200);
  });

  it('graphFetch throws after exhausting retries on network error', async () => {
    mockFetchTransient(10, 'ECONNREFUSED');
    await assert.rejects(
      () => client.listJoinedTeams(TOKEN),
      (err) => {
        assert.ok(err.message.includes('ECONNREFUSED'), 'error message should mention the code');
        return true;
      }
    );
  });

  it('graphFetch returns error response (non-429) without retry', async () => {
    const calls = mockFetch(
      { error: { code: 'MailboxNotFound', message: 'Mailbox not found' } },
      404,
    );
    const res = await client.getMessage(TOKEN, { messageId: 'bad-id' });
    assert.equal(calls.length, 1, 'should NOT retry on 404');
    assert.equal(res.status, 404);
  });
});

// ===========================================================================
// Headers & auth
// ===========================================================================

describe('Headers', () => {
  it('GET requests include Bearer token and Accept header', async () => {
    const calls = mockFetch({ value: [] });
    await client.listMessages(TOKEN);
    const h = calls[0].opts.headers;
    assert.equal(h.Authorization, `Bearer ${TOKEN}`);
    assert.equal(h.Accept, 'application/json');
    assert.equal(h['Content-Type'], undefined, 'GET should not have Content-Type');
  });

  it('POST requests include Content-Type application/json', async () => {
    const calls = mockFetch({});
    await client.acceptEvent(TOKEN, { eventId: 'e1' });
    const h = calls[0].opts.headers;
    assert.equal(h['Content-Type'], 'application/json');
  });
});
