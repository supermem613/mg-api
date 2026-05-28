#!/usr/bin/env node
'use strict';

const { describe, it } = require('node:test');
const assert = require('node:assert');
const { existsSync, mkdtempSync, rmSync, writeFileSync, mkdirSync } = require('node:fs');
const { tmpdir } = require('node:os');
const { join } = require('node:path');
const { pathToFileURL } = require('node:url');
const {
  authenticate,
  authStatus,
  classifyToken,
  compareGraphTokenScopes,
  decodeJwtPayload,
  hasCalendarScopes,
  hasChatScopes,
  hasChannelMessageScopes,
  hasMailScopes,
  isLoginUrl,
  logout,
  readAuthFile,
} = require('../src/graph-auth');
const {
  executeGraphRequest,
  loadAuth,
  parseResponseBody,
  resolveBase,
  resolveToken,
} = require('../src/graph-rest');
const { extractCode, formatError, graphFetch } = require('../src/graph-fetch');

function encodeJwt(payload) {
  const b64 = Buffer.from(JSON.stringify(payload)).toString('base64')
    .replace(/=+$/, '').replace(/\+/g, '-').replace(/\//g, '_');
  return `header.${b64}.signature`;
}

describe('mg-api source ownership', () => {
  it('keeps implementation logic in src instead of the skill directory', () => {
    assert.ok(existsSync(join(__dirname, '..', 'src', 'graph-auth.js')));
    assert.ok(existsSync(join(__dirname, '..', 'src', 'graph-rest.js')));
    assert.ok(existsSync(join(__dirname, '..', 'src', 'graph-fetch.js')));
    assert.strictEqual(existsSync(join(__dirname, '..', '.claude', 'skills', 'mg-api', 'scripts')), false);
    assert.strictEqual(existsSync(join(__dirname, '..', 'src', 'mcp')), false);
    assert.strictEqual(existsSync(join(__dirname, '..', 'src', 'cli')), false);
    assert.strictEqual(existsSync(join(__dirname, '..', 'src', 'core')), false);
  });
});

describe('Graph auth module', () => {
  it('decodes JWT payloads and classifies token audience', () => {
    const graph = encodeJwt({ aud: 'https://graph.microsoft.com', scp: 'Mail.ReadWrite User.Read' });
    assert.deepStrictEqual(decodeJwtPayload(graph), { aud: 'https://graph.microsoft.com', scp: 'Mail.ReadWrite User.Read' });
    assert.deepStrictEqual(classifyToken(graph), { type: 'graph', scopes: ['Mail.ReadWrite', 'User.Read'] });

    const outlook = encodeJwt({ aud: 'https://outlook.office.com', scp: 'Mail.Send' });
    assert.deepStrictEqual(classifyToken(outlook), { type: 'outlook', scopes: ['Mail.Send'] });

    const unknown = encodeJwt({ aud: 'https://example.com', scp: '' });
    assert.strictEqual(classifyToken(unknown), null);
    assert.strictEqual(decodeJwtPayload('not-a-jwt'), null);
  });

  it('detects mail and chat scopes case-insensitively', () => {
    assert.strictEqual(hasMailScopes(['Mail.ReadWrite', 'User.Read']), true);
    assert.strictEqual(hasMailScopes(['mail.read']), true);
    assert.strictEqual(hasMailScopes(['Chat.Read']), false);
    assert.strictEqual(hasChatScopes(['Chat.Read']), true);
    assert.strictEqual(hasChatScopes(['CHAT.READWRITE']), true);
    assert.strictEqual(hasChatScopes(['Chat.ReadWrite.All']), true);
    assert.strictEqual(hasChatScopes(['ChannelMessage.Read.All']), true);
    assert.strictEqual(hasChannelMessageScopes(['ChannelMessage.Read.All']), true);
    assert.strictEqual(hasChannelMessageScopes(['Chat.ReadWrite.All']), false);
    assert.strictEqual(hasChatScopes(['Mail.Read']), false);
  });

  it('detects calendar scopes including shared variants', () => {
    assert.strictEqual(hasCalendarScopes(['Calendars.Read']), true);
    assert.strictEqual(hasCalendarScopes(['calendars.readwrite']), true);
    assert.strictEqual(hasCalendarScopes(['Calendars.Read.Shared']), true);
    assert.strictEqual(hasCalendarScopes(['Calendars.ReadWrite.Shared']), true);
    assert.strictEqual(hasCalendarScopes(['CalendarsView.Read']), false);
    assert.strictEqual(hasCalendarScopes(['Mail.Read']), false);
  });

  it('prefers graph tokens with mail then calendar then more scopes', () => {
    // mail-scoped beats larger no-mail token (regression: OWA shell graph token
    // had 26 scopes including chat/files but no Mail.Read or Calendars.Read).
    const owaShell = ['Chat.Read', 'Files.ReadWrite.All', 'User.Read', 'People.Read', 'Group.ReadWrite.All', 'OnlineMeetings.Read', 'Directory.Read.All', 'Channel.ReadBasic.All'];
    const teamsWeb = ['Mail.Read', 'Calendars.Read', 'User.Read'];
    assert.ok(compareGraphTokenScopes(teamsWeb, owaShell) > 0);
    assert.ok(compareGraphTokenScopes(owaShell, teamsWeb) < 0);

    // both have mail: calendar tie-breaker wins
    const mailOnly = ['Mail.Read', 'User.Read'];
    const mailAndCal = ['Mail.Read', 'Calendars.Read'];
    assert.ok(compareGraphTokenScopes(mailAndCal, mailOnly) > 0);

    // identical mail+calendar profile: more total scopes wins
    const small = ['Mail.Read', 'Calendars.Read'];
    const big = ['Mail.Read', 'Calendars.Read', 'Files.Read'];
    assert.ok(compareGraphTokenScopes(big, small) > 0);

    // starts from empty captured set: anything beats nothing
    assert.ok(compareGraphTokenScopes(['User.Read'], []) > 0);
  });

  it('detects Microsoft login URLs', () => {
    assert.strictEqual(isLoginUrl('https://login.microsoftonline.com/common'), true);
    assert.strictEqual(isLoginUrl('https://login.microsoft.com/oauth2'), true);
    assert.strictEqual(isLoginUrl('https://login.live.com/'), true);
    assert.strictEqual(isLoginUrl('https://outlook.office.com/mail/'), false);
    assert.strictEqual(isLoginUrl('not-a-url'), false);
  });

  it('reports auth status from the auth file', () => {
    const dir = mkdtempSync(join(tmpdir(), 'mg-api-auth-'));
    try {
      const authFile = join(dir, 'auth.json');
      writeFileSync(authFile, JSON.stringify({
        GRAPH_TOKEN: 'g',
        OUTLOOK_TOKEN: 'o',
        GRAPH_SCOPES: ['Mail.Read'],
      }));
      assert.deepStrictEqual(authStatus(authFile), {
        authFile,
        exists: true,
        hasGraphToken: true,
        hasOutlookToken: true,
        hasChatToken: false,
        hasChannelMessageToken: false,
        channelMessageScopeObserved: false,
        teamsChannelProbe: null,
        graphScopes: ['Mail.Read'],
        outlookScopes: [],
        chatScopes: [],
      });
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });

  it('reports missing auth as not existing', () => {
    const dir = mkdtempSync(join(tmpdir(), 'mg-api-auth-missing-'));
    try {
      const status = authStatus(join(dir, 'auth.json'));
      assert.strictEqual(status.exists, false);
      assert.strictEqual(status.hasGraphToken, false);
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });

  it('logs out by clearing auth state paths', () => {
    const dir = mkdtempSync(join(tmpdir(), 'mg-api-logout-'));
    try {
      const profileDir = join(dir, 'profile');
      const authFile = join(dir, 'auth.json');
      writeFileSync(authFile, '{}');
      mkdirSync(profileDir);
      const result = logout({ profileDir, authFile });
      assert.strictEqual(result.cleared, true);
      assert.strictEqual(existsSync(authFile), false);
      assert.strictEqual(existsSync(profileDir), false);
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });

  it('authenticates with injectable Playwright and writes auth.json', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'mg-api-playwright-'));
    try {
      const authFile = join(dir, 'auth.json');
      const graphToken = encodeJwt({ aud: 'https://graph.microsoft.com', scp: 'Mail.ReadWrite User.Read' });
      const outlookToken = encodeJwt({ aud: 'https://outlook.office.com', scp: 'Mail.Send' });
      const chatToken = encodeJwt({ aud: 'https://graph.microsoft.com', scp: 'ChannelMessage.Read.All' });

      function makePage(authHeaders) {
        let handler;
        return {
          on: (evt, fn) => { if (evt === 'request') handler = fn; },
          goto: async () => {
            for (const auth of authHeaders) {
              if (handler) handler({ headers: () => ({ authorization: auth }) });
            }
          },
          url: () => 'https://outlook.office.com/mail/',
          waitForLoadState: async () => {},
          waitForURL: async () => {},
        };
      }
      const outlookPage = makePage([`Bearer ${graphToken}`, `Bearer ${outlookToken}`]);
      const teamsPage = makePage([`Bearer ${chatToken}`]);
      const officePage = makePage([]);
      let nextPage = 0;
      const newPages = [teamsPage, officePage];

      const playwright = {
        chromium: {
          launchPersistentContext: async () => ({
            pages: () => [outlookPage],
            newPage: async () => newPages[nextPage++],
            cookies: async () => [],
            close: async () => {},
          }),
        },
      };

      const result = await authenticate({ playwright, authFile, profileDir: join(dir, 'profile') });
      assert.strictEqual(result.GRAPH_TOKEN, graphToken);
      assert.strictEqual(result.OUTLOOK_TOKEN, outlookToken);
      assert.strictEqual(result.GRAPH_CHAT_TOKEN, chatToken);
      const persisted = readAuthFile(authFile);
      assert.strictEqual(persisted.GRAPH_TOKEN, graphToken);
      assert.deepStrictEqual(persisted.GRAPH_SCOPES, ['Mail.ReadWrite', 'User.Read']);
      assert.deepStrictEqual(persisted.OUTLOOK_SCOPES, ['Mail.Send']);
      assert.deepStrictEqual(persisted.GRAPH_CHAT_SCOPES, ['ChannelMessage.Read.All']);
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });

  it('preserves the mail-scoped graph token when a larger OWA shell token arrives later', async () => {
    // Regression: OWA emits a graph token with many scopes (chat, files,
    // meetings) but no Mail.Read / Calendars.Read. Without the priority
    // comparator the shell token clobbers an earlier mail-scoped capture
    // and email.list / calendar.list start returning 403.
    const dir = mkdtempSync(join(tmpdir(), 'mg-api-graph-priority-'));
    try {
      const authFile = join(dir, 'auth.json');
      const mailScoped = encodeJwt({ aud: 'https://graph.microsoft.com', scp: 'Mail.Read Calendars.Read User.Read' });
      const owaShell = encodeJwt({ aud: 'https://graph.microsoft.com', scp: 'Chat.Read Files.ReadWrite.All User.Read People.Read Group.ReadWrite.All OnlineMeetings.Read Directory.Read.All Channel.ReadBasic.All' });
      const outlookToken = encodeJwt({ aud: 'https://outlook.office.com', scp: 'Mail.Send' });

      function makePage(authHeaders) {
        let handler;
        return {
          on: (evt, fn) => { if (evt === 'request') handler = fn; },
          goto: async () => {
            for (const auth of authHeaders) {
              if (handler) handler({ headers: () => ({ authorization: auth }) });
            }
          },
          url: () => 'https://outlook.office.com/mail/',
          waitForLoadState: async () => {},
          waitForURL: async () => {},
        };
      }
      const outlookPage = makePage([`Bearer ${mailScoped}`, `Bearer ${owaShell}`, `Bearer ${outlookToken}`]);
      const teamsPage = makePage([]);
      const officePage = makePage([]);
      let nextPage = 0;
      const newPages = [teamsPage, officePage];

      const playwright = {
        chromium: {
          launchPersistentContext: async () => ({
            pages: () => [outlookPage],
            newPage: async () => newPages[nextPage++],
            cookies: async () => [],
            close: async () => {},
          }),
        },
      };

      const result = await authenticate({ playwright, authFile, profileDir: join(dir, 'profile') });
      assert.strictEqual(result.GRAPH_TOKEN, mailScoped, 'mail-scoped graph token must win over larger OWA shell token');
      assert.deepStrictEqual(result.GRAPH_SCOPES, ['Mail.Read', 'Calendars.Read', 'User.Read']);
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });

  it('probes a Teams channel surface and persists the channel-message token', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'mg-api-channel-probe-'));
    try {
      const authFile = join(dir, 'auth.json');
      const graphToken = encodeJwt({ aud: 'https://graph.microsoft.com', scp: 'Mail.ReadWrite User.Read' });
      const chatToken = encodeJwt({ aud: 'https://graph.microsoft.com', scp: 'Chat.Read Chat.ReadWrite' });
      const channelMessageToken = encodeJwt({ aud: 'https://graph.microsoft.com', scp: 'ChannelMessage.Read.All' });
      const teamsUrls = [];
      let channelClicked = false;

      function emit(handler, token) {
        if (handler) handler({ headers: () => ({ authorization: `Bearer ${token}` }) });
      }

      function makePage(tokensByUrl = new Map()) {
        let handler;
        return {
          on: (evt, fn) => { if (evt === 'request') handler = fn; },
          goto: async url => {
            if (url.includes('teams.cloud.microsoft')) teamsUrls.push(url);
            for (const token of tokensByUrl.get(url) || []) emit(handler, token);
          },
          locator: selector => ({
            count: async () => (selector.includes('/l/channel/') ? 1 : 0),
            nth: () => ({
              isVisible: async () => true,
              click: async () => {
                channelClicked = true;
                emit(handler, channelMessageToken);
              },
            }),
          }),
          url: () => 'https://outlook.office.com/mail/',
          waitForLoadState: async () => {},
          waitForRequest: async () => {},
          waitForURL: async () => {},
        };
      }

      const outlookPage = makePage(new Map([['https://outlook.office.com/mail/', [graphToken]]]));
      const teamsPage = makePage(new Map([['https://teams.cloud.microsoft/v2/chat', [chatToken]]]));
      const officePage = makePage();
      let nextPage = 0;
      const newPages = [teamsPage, officePage];
      const playwright = {
        chromium: {
          launchPersistentContext: async () => ({
            pages: () => [outlookPage],
            newPage: async () => newPages[nextPage++],
            cookies: async () => [],
            close: async () => {},
          }),
        },
      };

      const result = await authenticate({ playwright, authFile, profileDir: join(dir, 'profile') });
      assert.ok(teamsUrls.includes('https://teams.cloud.microsoft/v2/teams'));
      assert.strictEqual(channelClicked, true);
      assert.strictEqual(result.GRAPH_CHAT_TOKEN, channelMessageToken);
      assert.strictEqual(result.CHANNEL_MESSAGE_SCOPE_OBSERVED, true);
      assert.deepStrictEqual(result.TEAMS_CHANNEL_PROBE, { attempted: true, opened: true, observed: true });
      const persisted = readAuthFile(authFile);
      assert.strictEqual(persisted.GRAPH_CHAT_TOKEN, channelMessageToken);
      assert.deepStrictEqual(persisted.GRAPH_CHAT_SCOPES, ['ChannelMessage.Read.All']);
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });

  it('discovers and opens a joined Teams channel when channel links are not rendered', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'mg-api-channel-discovery-'));
    try {
      const authFile = join(dir, 'auth.json');
      const graphToken = encodeJwt({ aud: 'https://graph.microsoft.com', scp: 'Mail.ReadWrite User.Read Channel.ReadBasic.All' });
      const chatToken = encodeJwt({ aud: 'https://graph.microsoft.com', scp: 'Chat.Read' });
      const channelMessageToken = encodeJwt({ aud: 'https://graph.microsoft.com', scp: 'ChannelMessage.Read.All' });
      const channelUrl = 'https://teams.microsoft.com/l/channel/19%3Ageneral/General?groupId=team-1';
      const visitedUrls = [];
      const fetchUrls = [];

      function emit(handler, token) {
        if (handler) handler({ headers: () => ({ authorization: `Bearer ${token}` }) });
      }

      function makePage(tokensByUrl = new Map()) {
        let handler;
        return {
          on: (evt, fn) => { if (evt === 'request') handler = fn; },
          goto: async url => {
            visitedUrls.push(url);
            for (const token of tokensByUrl.get(url) || []) emit(handler, token);
          },
          locator: () => ({ count: async () => 0 }),
          url: () => 'https://outlook.office.com/mail/',
          waitForLoadState: async () => {},
          waitForRequest: async () => {},
          waitForURL: async () => {},
        };
      }

      const fetch = async url => {
        fetchUrls.push(url);
        if (url.includes('/me/joinedTeams')) {
          return { ok: true, status: 200, json: async () => ({ value: [{ id: 'team-1', displayName: 'Team 1' }] }) };
        }
        if (url.includes('/teams/team-1/channels')) {
          return { ok: true, status: 200, json: async () => ({ value: [{ id: '19:general', displayName: 'General', webUrl: channelUrl }] }) };
        }
        return { ok: false, status: 404, json: async () => ({}) };
      };

      const outlookPage = makePage(new Map([['https://outlook.office.com/mail/', [graphToken]]]));
      const teamsPage = makePage(new Map([
        ['https://teams.cloud.microsoft/v2/chat', [chatToken]],
        [channelUrl, [channelMessageToken]],
      ]));
      const officePage = makePage();
      let nextPage = 0;
      const newPages = [teamsPage, officePage];
      const playwright = {
        chromium: {
          launchPersistentContext: async () => ({
            pages: () => [outlookPage],
            newPage: async () => newPages[nextPage++],
            cookies: async () => [],
            close: async () => {},
          }),
        },
      };

      const result = await authenticate({ playwright, fetch, authFile, profileDir: join(dir, 'profile') });
      assert.ok(fetchUrls.some(url => url.includes('/me/joinedTeams')));
      assert.ok(fetchUrls.some(url => url.includes('/teams/team-1/channels')));
      assert.ok(visitedUrls.includes(channelUrl));
      assert.strictEqual(result.GRAPH_CHAT_TOKEN, channelMessageToken);
      assert.strictEqual(result.CHANNEL_MESSAGE_SCOPE_OBSERVED, true);
      assert.deepStrictEqual(result.TEAMS_CHANNEL_PROBE, { attempted: true, opened: true, observed: true });
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });

  it('uses the captured Teams Graph token for channel discovery before the generic Graph token is available', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'mg-api-chat-discovery-token-'));
    try {
      const authFile = join(dir, 'auth.json');
      const graphToken = encodeJwt({ aud: 'https://graph.microsoft.com', scp: 'Mail.ReadWrite User.Read' });
      const chatToken = encodeJwt({ aud: 'https://graph.microsoft.com', scp: 'Chat.Read Channel.ReadBasic.All' });
      const channelMessageToken = encodeJwt({ aud: 'https://graph.microsoft.com', scp: 'ChannelMessage.Read.All' });
      const channelUrl = 'https://teams.microsoft.com/l/channel/19%3Ageneral/General?groupId=team-1';
      const fetchAuthorizations = [];

      function emit(handler, token) {
        if (handler) handler({ headers: () => ({ authorization: `Bearer ${token}` }) });
      }

      function makePage(tokensByUrl = new Map()) {
        let handler;
        return {
          on: (evt, fn) => { if (evt === 'request') handler = fn; },
          goto: async url => {
            for (const token of tokensByUrl.get(url) || []) emit(handler, token);
          },
          locator: () => ({ count: async () => 0 }),
          url: () => 'https://outlook.office.com/mail/',
          waitForLoadState: async () => {},
          waitForRequest: async () => {},
          waitForURL: async () => {},
        };
      }

      const fetch = async (url, options) => {
        fetchAuthorizations.push(options.headers.Authorization);
        if (url.includes('/me/joinedTeams')) {
          return { ok: true, status: 200, json: async () => ({ value: [{ id: 'team-1', displayName: 'Team 1' }] }) };
        }
        if (url.includes('/teams/team-1/channels')) {
          return { ok: true, status: 200, json: async () => ({ value: [{ id: '19:general', displayName: 'General', webUrl: channelUrl }] }) };
        }
        return { ok: false, status: 404, json: async () => ({}) };
      };

      const outlookPage = makePage();
      const teamsPage = makePage(new Map([
        ['https://teams.cloud.microsoft/v2/chat', [chatToken]],
        [channelUrl, [channelMessageToken]],
      ]));
      const officePage = makePage(new Map([['https://www.office.com/', [graphToken]]]));
      let nextPage = 0;
      const newPages = [teamsPage, officePage];
      const playwright = {
        chromium: {
          launchPersistentContext: async () => ({
            pages: () => [outlookPage],
            newPage: async () => newPages[nextPage++],
            cookies: async () => [],
            close: async () => {},
          }),
        },
      };

      const result = await authenticate({ playwright, fetch, authFile, profileDir: join(dir, 'profile') });
      assert.strictEqual(fetchAuthorizations[0], `Bearer ${chatToken}`);
      assert.strictEqual(result.GRAPH_TOKEN, graphToken);
      assert.strictEqual(result.GRAPH_CHAT_TOKEN, channelMessageToken);
      assert.strictEqual(result.CHANNEL_MESSAGE_SCOPE_OBSERVED, true);
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });

  it('makes missing ChannelMessage.Read.All visible after the Teams channel probe', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'mg-api-missing-channel-scope-'));
    const originalWrite = process.stderr.write;
    const stderr = [];
    process.stderr.write = chunk => { stderr.push(String(chunk)); return true; };
    try {
      const authFile = join(dir, 'auth.json');
      const graphToken = encodeJwt({ aud: 'https://graph.microsoft.com', scp: 'Mail.ReadWrite User.Read' });
      const chatToken = encodeJwt({ aud: 'https://graph.microsoft.com', scp: 'Chat.Read' });

      function makePage(tokens = []) {
        let handler;
        return {
          on: (evt, fn) => { if (evt === 'request') handler = fn; },
          goto: async () => {
            for (const token of tokens) {
              emit(handler, token);
            }
          },
          locator: () => ({
            count: async () => 1,
            nth: () => ({ isVisible: async () => true, click: async () => {} }),
          }),
          url: () => 'https://outlook.office.com/mail/',
          waitForLoadState: async () => {},
          waitForRequest: async () => { throw new Error('timeout'); },
          waitForURL: async () => {},
        };
      }

      function emit(handler, token) {
        if (handler) handler({ headers: () => ({ authorization: `Bearer ${token}` }) });
      }

      const outlookPage = makePage([graphToken]);
      const teamsPage = makePage([chatToken]);
      const officePage = makePage([]);
      let nextPage = 0;
      const newPages = [teamsPage, officePage];
      const playwright = {
        chromium: {
          launchPersistentContext: async () => ({
            pages: () => [outlookPage],
            newPage: async () => newPages[nextPage++],
            cookies: async () => [],
            close: async () => {},
          }),
        },
      };

      const result = await authenticate({ playwright, authFile, profileDir: join(dir, 'profile') });
      assert.strictEqual(result.GRAPH_CHAT_TOKEN, chatToken);
      assert.strictEqual(result.CHANNEL_MESSAGE_SCOPE_OBSERVED, false);
      assert.deepStrictEqual(result.TEAMS_CHANNEL_PROBE, { attempted: true, opened: true, observed: false });
      assert.match(stderr.join(''), /ChannelMessage\.Read\.All was not observed during Teams channel probe/);
      const status = authStatus(authFile);
      assert.strictEqual(status.hasChatToken, true);
      assert.strictEqual(status.hasChannelMessageToken, false);
      assert.strictEqual(status.channelMessageScopeObserved, false);
    } finally {
      process.stderr.write = originalWrite;
      rmSync(dir, { recursive: true, force: true });
    }
  });

  it('prefers channel-message Graph tokens over broader chat tokens', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'mg-api-teams-token-'));
    try {
      const authFile = join(dir, 'auth.json');
      const graphToken = encodeJwt({ aud: 'https://graph.microsoft.com', scp: 'Mail.ReadWrite User.Read' });
      const broadChatToken = encodeJwt({ aud: 'https://graph.microsoft.com', scp: 'Chat.Read User.Read Mail.Read' });
      const channelMessageToken = encodeJwt({ aud: 'https://graph.microsoft.com', scp: 'ChannelMessage.Read.All' });

      function makePage(authHeaders) {
        let handler;
        return {
          on: (evt, fn) => { if (evt === 'request') handler = fn; },
          goto: async () => {
            for (const auth of authHeaders) {
              if (handler) handler({ headers: () => ({ authorization: auth }) });
            }
          },
          url: () => 'https://outlook.office.com/mail/',
          waitForLoadState: async () => {},
          waitForURL: async () => {},
        };
      }
      const outlookPage = makePage([`Bearer ${graphToken}`]);
      const teamsPage = makePage([`Bearer ${broadChatToken}`, `Bearer ${channelMessageToken}`]);
      const officePage = makePage([]);
      let nextPage = 0;
      const newPages = [teamsPage, officePage];

      const playwright = {
        chromium: {
          launchPersistentContext: async () => ({
            pages: () => [outlookPage],
            newPage: async () => newPages[nextPage++],
            cookies: async () => [],
            close: async () => {},
          }),
        },
      };

      const result = await authenticate({ playwright, authFile, profileDir: join(dir, 'profile') });
      assert.strictEqual(result.GRAPH_CHAT_TOKEN, channelMessageToken);
      assert.deepStrictEqual(result.GRAPH_CHAT_SCOPES, ['ChannelMessage.Read.All']);
      const persisted = readAuthFile(authFile);
      assert.strictEqual(persisted.GRAPH_CHAT_TOKEN, channelMessageToken);
      assert.deepStrictEqual(persisted.GRAPH_CHAT_SCOPES, ['ChannelMessage.Read.All']);
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });

  it('throws a clear error when no tokens are captured', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'mg-api-no-token-'));
    try {
      const authFile = join(dir, 'auth.json');
      const emptyPage = {
        on: () => {},
        goto: async () => {},
        url: () => 'https://outlook.office.com/mail/',
        waitForLoadState: async () => {},
        waitForURL: async () => {},
      };
      const playwright = {
        chromium: {
          launchPersistentContext: async () => ({
            pages: () => [emptyPage],
            newPage: async () => emptyPage,
            cookies: async () => [],
            close: async () => {},
          }),
        },
      };
      await assert.rejects(
        () => authenticate({ playwright, authFile, profileDir: join(dir, 'profile') }),
        /No tokens captured/,
      );
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });
});

describe('Graph REST module', () => {
  it('rejects missing or empty auth files', () => {
    const dir = mkdtempSync(join(tmpdir(), 'mg-api-rest-'));
    try {
      const authFile = join(dir, 'auth.json');
      assert.throws(() => loadAuth(authFile), /Run mg-api auth login/);
      writeFileSync(authFile, JSON.stringify({}));
      assert.throws(() => loadAuth(authFile), /Graph or Outlook token/);
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });

  it('parses JSON, text, and empty response bodies', () => {
    assert.deepStrictEqual(parseResponseBody('{"value":1}'), { value: 1 });
    assert.strictEqual(parseResponseBody('plain'), 'plain');
    assert.strictEqual(parseResponseBody(''), null);
  });

  it('resolves token and base from the verb spec', () => {
    const auth = { GRAPH_TOKEN: 'g', OUTLOOK_TOKEN: 'o', GRAPH_CHAT_TOKEN: 'c' };
    assert.strictEqual(resolveToken({ id: 'x.y', token: 'graph' }, auth), 'g');
    assert.strictEqual(resolveToken({ id: 'x.y', token: 'outlook' }, auth), 'o');
    assert.strictEqual(resolveToken({ id: 'x.y', token: 'chat' }, auth), 'c');
    assert.strictEqual(resolveBase({ base: 'graph' }), 'https://graph.microsoft.com/v1.0');
    assert.strictEqual(resolveBase({ base: 'outlook' }), 'https://outlook.office.com/api/v2.0');
  });

  it('falls back to the graph token when chat is requested but not cached', () => {
    const auth = { GRAPH_TOKEN: 'g', OUTLOOK_TOKEN: 'o' };
    assert.strictEqual(resolveToken({ id: 'x.y', token: 'chat' }, auth), 'g');
  });

  it('throws when the requested token is not present', () => {
    assert.throws(() => resolveToken({ id: 'x.y', token: 'outlook' }, { GRAPH_TOKEN: 'g' }), /Outlook token/);
    assert.throws(() => resolveToken({ id: 'x.y', token: 'graph' }, { OUTLOOK_TOKEN: 'o' }), /Graph token/);
  });

  it('executes GET requests through the built-in REST client with the right token', async () => {
    const calls = [];
    const result = await executeGraphRequest(
      { id: 'email.list', method: 'GET', token: 'graph', base: 'graph' },
      '/me/messages?%24top=5',
      '',
      {
        auth: { GRAPH_TOKEN: 'graph-tok', OUTLOOK_TOKEN: 'out-tok' },
        fetch: async (url, options) => {
          calls.push({ url, options });
          return { ok: true, status: 200, text: async () => '{"value":[]}' };
        },
      },
    );
    assert.deepStrictEqual(result.data, { value: [] });
    assert.strictEqual(calls[0].url, 'https://graph.microsoft.com/v1.0/me/messages?%24top=5');
    assert.strictEqual(calls[0].options.method, 'GET');
    assert.strictEqual(calls[0].options.headers.Authorization, 'Bearer graph-tok');
    assert.strictEqual(result.token, 'graph');
    assert.strictEqual(result.base, 'graph');
  });

  it('sends POST bodies with content type and the outlook base when configured', async () => {
    const calls = [];
    await executeGraphRequest(
      { id: 'email.send', method: 'POST', token: 'outlook', base: 'outlook' },
      '/me/sendmail',
      '{"Message":{"Subject":"hi"}}',
      {
        auth: { OUTLOOK_TOKEN: 'out-tok' },
        fetch: async (url, options) => {
          calls.push({ url, options });
          return { ok: true, status: 202, text: async () => '' };
        },
      },
    );
    assert.strictEqual(calls[0].url, 'https://outlook.office.com/api/v2.0/me/sendmail');
    assert.strictEqual(calls[0].options.method, 'POST');
    assert.strictEqual(calls[0].options.headers['Content-Type'], 'application/json');
    assert.strictEqual(calls[0].options.headers.Authorization, 'Bearer out-tok');
    assert.strictEqual(calls[0].options.body, '{"Message":{"Subject":"hi"}}');
  });

  it('selects the chat token for chat-scoped verbs', async () => {
    const calls = [];
    await executeGraphRequest(
      { id: 'chats.send', method: 'POST', token: 'chat', base: 'graph' },
      '/chats/19%3Aabc/messages',
      '{"body":{"content":"hi"}}',
      {
        auth: { GRAPH_TOKEN: 'g', GRAPH_CHAT_TOKEN: 'c' },
        fetch: async (url, options) => {
          calls.push({ url, options });
          return { ok: true, status: 201, text: async () => '{}' };
        },
      },
    );
    assert.strictEqual(calls[0].options.headers.Authorization, 'Bearer c');
  });

  it('throws on non-2xx responses with the status and method', async () => {
    await assert.rejects(
      () => executeGraphRequest(
        { id: 'email.list', method: 'GET', token: 'graph', base: 'graph' },
        '/me/messages',
        '',
        {
          auth: { GRAPH_TOKEN: 'g' },
          fetch: async () => ({ ok: false, status: 403, text: async () => 'forbidden' }),
        },
      ),
      /HTTP 403 on GET/,
    );
  });
});

describe('Graph fetch module', () => {
  it('walks nested error causes', () => {
    assert.strictEqual(extractCode({ cause: { code: 'ETIMEDOUT' } }), 'ETIMEDOUT');
    assert.strictEqual(extractCode({ errors: [{ code: 'ECONNRESET' }] }), 'ECONNRESET');
    assert.strictEqual(extractCode({ message: 'plain' }), null);
  });

  it('formats actionable network errors', () => {
    assert.match(formatError({ message: 'failed', code: 'ENOTFOUND' }), /Hint: DNS lookup failed/);
    assert.match(formatError({ message: 'down', code: 'ECONNREFUSED' }), /Hint: Connection refused/);
    assert.match(formatError({ message: 'slow', code: 'ETIMEDOUT' }), /Hint: Connection timed out/);
  });

  it('retries retryable fetch failures without shelling out', async () => {
    let attempts = 0;
    const response = await graphFetch(pathToFileURL(__filename).toString(), {}, {
      fetch: async () => {
        attempts++;
        if (attempts === 1) {
          const err = new Error('timeout');
          err.code = 'ETIMEDOUT';
          throw err;
        }
        return { ok: true };
      },
    });
    assert.strictEqual(response.ok, true);
    assert.strictEqual(attempts, 2);
  });

  it('retries on HTTP 429 honoring Retry-After', async () => {
    let attempts = 0;
    const response = await graphFetch('https://example.com', {}, {
      fetch: async () => {
        attempts++;
        if (attempts === 1) {
          return {
            ok: false,
            status: 429,
            headers: { get: name => (name === 'Retry-After' ? '0' : null) },
            text: async () => '',
          };
        }
        return { ok: true, status: 200, text: async () => '{}' };
      },
    });
    assert.strictEqual(response.ok, true);
    assert.strictEqual(attempts, 2);
  });
});
