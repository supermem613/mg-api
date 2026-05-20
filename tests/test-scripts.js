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
  decodeJwtPayload,
  hasChatScopes,
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
    assert.strictEqual(hasChatScopes(['Mail.Read']), false);
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
