'use strict';

const fs = require('fs');
const os = require('os');
const path = require('path');

const DATA_DIR = path.join(os.homedir(), '.mg-api');
const PROFILE_DIR = path.join(DATA_DIR, 'browser-profile');
const AUTH_FILE = path.join(DATA_DIR, 'auth.json');
const LOGIN_TIMEOUT_MS = 300_000;
const HEADLESS_PROBE_MS = 5_000;

const OUTLOOK_URL = 'https://outlook.office.com/mail/';
const TEAMS_URL = 'https://teams.cloud.microsoft/';
const TEAMS_CHAT_URL = 'https://teams.cloud.microsoft/v2/chat';
const OFFICE_URL = 'https://www.office.com/';
const PAGE_TIMEOUT_MS = 20_000;
const NETWORK_IDLE_MS = 8_000;

function ensureDir(dir) {
  fs.mkdirSync(dir, { recursive: true });
}

function readAuthFile(authFile = AUTH_FILE) {
  try {
    return JSON.parse(fs.readFileSync(authFile, 'utf8'));
  } catch {
    return null;
  }
}

function authStatus(authFile = AUTH_FILE) {
  const parsed = readAuthFile(authFile);
  return {
    authFile,
    exists: !!parsed,
    hasGraphToken: !!parsed?.GRAPH_TOKEN,
    hasOutlookToken: !!parsed?.OUTLOOK_TOKEN,
    hasChatToken: !!parsed?.GRAPH_CHAT_TOKEN,
    graphScopes: parsed?.GRAPH_SCOPES || [],
    outlookScopes: parsed?.OUTLOOK_SCOPES || [],
    chatScopes: parsed?.GRAPH_CHAT_SCOPES || [],
  };
}

function requirePlaywright() {
  try {
    return require('playwright');
  } catch {
    throw new Error('playwright is not installed. Run npm install in the mg-api repo.');
  }
}

function isLoginUrl(url) {
  try {
    const host = new URL(url).hostname.toLowerCase();
    return host.includes('login.microsoftonline.com')
      || host.includes('login.microsoft.com')
      || host.includes('login.live.com');
  } catch {
    return false;
  }
}

function decodeJwtPayload(token) {
  try {
    return JSON.parse(Buffer.from(token.split('.')[1], 'base64').toString());
  } catch {
    return null;
  }
}

function classifyToken(token) {
  const payload = decodeJwtPayload(token);
  if (!payload) return null;
  const aud = (payload.aud || '').toLowerCase();
  const scopes = (payload.scp || '').split(' ').filter(Boolean);
  if (aud.includes('graph.microsoft.com')) return { type: 'graph', scopes };
  if (aud.includes('outlook.office.com') || aud.includes('outlook.office365.com')) return { type: 'outlook', scopes };
  return null;
}

function hasMailScopes(scopes) {
  return scopes.some(s => /^mail\.(read|readwrite)$/i.test(s));
}

function hasChatScopes(scopes) {
  return scopes.some(s => /^(chat\.read|chat\.readwrite|chat\.readwrite\.all|channelmessage\.read\.all)$/i.test(s));
}

function hasChannelMessageScopes(scopes) {
  return scopes.some(s => /^channelmessage\.read\.all$/i.test(s));
}

function compareTeamsTokenScopes(candidateScopes, currentScopes) {
  const candidateChannel = hasChannelMessageScopes(candidateScopes);
  const currentChannel = hasChannelMessageScopes(currentScopes);
  if (candidateChannel !== currentChannel) return candidateChannel ? 1 : -1;
  return candidateScopes.length - currentScopes.length;
}

async function authenticate(options = {}) {
  const {
    forceLogin = false,
    verbose = false,
    playwright = requirePlaywright(),
    authFile = AUTH_FILE,
    profileDir = PROFILE_DIR,
  } = options;

  ensureDir(profileDir);

  const captured = {
    graph: null, graphScopes: [],
    graphChat: null, graphChatScopes: [],
    outlook: null, outlookScopes: [],
  };

  function installTokenInterceptor(p) {
    p.on('request', request => {
      const auth = request.headers()['authorization'];
      if (!auth?.startsWith('Bearer ')) return;
      const token = auth.substring(7);
      const info = classifyToken(token);
      if (!info) return;
      if (info.type === 'graph') {
        if (hasMailScopes(info.scopes) && !hasMailScopes(captured.graphScopes)) {
          captured.graph = token;
          captured.graphScopes = info.scopes;
        } else if (info.scopes.length > captured.graphScopes.length) {
          captured.graph = token;
          captured.graphScopes = info.scopes;
        }
        if (hasChatScopes(info.scopes) && compareTeamsTokenScopes(info.scopes, captured.graphChatScopes) > 0) {
          captured.graphChat = token;
          captured.graphChatScopes = info.scopes;
        }
      } else if (info.type === 'outlook' && info.scopes.length > captured.outlookScopes.length) {
        captured.outlook = token;
        captured.outlookScopes = info.scopes;
      }
    });
  }

  let headless = !forceLogin;
  const log = (msg) => { if (verbose) process.stderr.write(`[auth] ${msg}\n`); };

  log(`launching Edge (${headless ? 'headless probe' : 'visible'}, profile=${profileDir})`);
  let context = await playwright.chromium.launchPersistentContext(profileDir, {
    channel: 'msedge',
    headless,
    args: ['--disable-blink-features=AutomationControlled'],
    viewport: { width: 1280, height: 800 },
  });
  let page = context.pages()[0] || await context.newPage();
  installTokenInterceptor(page);

  log(`loading ${OUTLOOK_URL}`);
  await page.goto(OUTLOOK_URL, { waitUntil: 'domcontentloaded', timeout: PAGE_TIMEOUT_MS }).catch(err => log(`Outlook goto failed: ${err.message}`));

  if (headless && isLoginUrl(page.url())) {
    log('login redirect detected — relaunching visible Edge for sign-in');
    try {
      await page.waitForURL(url => !isLoginUrl(url.toString()), { timeout: HEADLESS_PROBE_MS });
    } catch {
      await context.close();
      process.stderr.write('Opening Edge for login. Complete sign-in in the browser window.\n');
      context = await playwright.chromium.launchPersistentContext(profileDir, {
        channel: 'msedge',
        headless: false,
        args: ['--disable-blink-features=AutomationControlled'],
        viewport: { width: 1280, height: 800 },
      });
      page = context.pages()[0] || await context.newPage();
      installTokenInterceptor(page);
      await page.goto(OUTLOOK_URL, { waitUntil: 'domcontentloaded', timeout: PAGE_TIMEOUT_MS }).catch(err => log(`Outlook goto failed: ${err.message}`));
      headless = false;
    }
  }

  if (!headless && isLoginUrl(page.url())) {
    log(`waiting up to ${LOGIN_TIMEOUT_MS / 1000}s for sign-in to complete`);
    try {
      await page.waitForURL(url => !isLoginUrl(url.toString()), { timeout: LOGIN_TIMEOUT_MS });
      log('sign-in complete');
    } catch {
      await context.close();
      throw new Error('Login timed out or browser was closed.');
    }
  }

  log('settling Outlook page');
  await page.waitForLoadState('networkidle', { timeout: NETWORK_IDLE_MS }).catch(() => {});

  log('visiting Teams to capture chat scopes');
  const teamsPage = await context.newPage();
  installTokenInterceptor(teamsPage);
  await teamsPage.goto(TEAMS_URL, { waitUntil: 'domcontentloaded', timeout: PAGE_TIMEOUT_MS }).catch(err => log(`Teams goto failed (continuing): ${err.message}`));
  await teamsPage.waitForLoadState('networkidle', { timeout: NETWORK_IDLE_MS }).catch(() => {});
  await teamsPage.goto(TEAMS_CHAT_URL, { waitUntil: 'domcontentloaded', timeout: PAGE_TIMEOUT_MS }).catch(err => log(`Teams chat goto failed (continuing): ${err.message}`));
  await teamsPage.waitForLoadState('networkidle', { timeout: NETWORK_IDLE_MS }).catch(() => {});

  log('visiting office.com to capture Outlook scopes');
  const officePage = await context.newPage();
  installTokenInterceptor(officePage);
  await officePage.goto(OFFICE_URL, { waitUntil: 'domcontentloaded', timeout: PAGE_TIMEOUT_MS }).catch(err => log(`Office goto failed (continuing): ${err.message}`));
  await officePage.waitForLoadState('networkidle', { timeout: NETWORK_IDLE_MS }).catch(() => {});

  log(`captured: graph=${!!captured.graph} chat=${!!captured.graphChat} outlook=${!!captured.outlook}`);
  await context.close();

  if (!captured.graph && !captured.outlook) {
    throw new Error('No tokens captured. Run mg-api auth login --force to force a fresh login.');
  }

  const authData = {
    ...(captured.graph && { GRAPH_TOKEN: captured.graph }),
    ...(captured.graphChat && { GRAPH_CHAT_TOKEN: captured.graphChat }),
    ...(captured.outlook && { OUTLOOK_TOKEN: captured.outlook }),
    ...(captured.graphScopes.length && { GRAPH_SCOPES: captured.graphScopes }),
    ...(captured.graphChatScopes.length && { GRAPH_CHAT_SCOPES: captured.graphChatScopes }),
    ...(captured.outlookScopes.length && { OUTLOOK_SCOPES: captured.outlookScopes }),
  };

  ensureDir(path.dirname(authFile));
  fs.writeFileSync(authFile, JSON.stringify(authData, null, 2) + '\n');
  log(`wrote ${authFile}`);
  return authData;
}

function logout(options = {}) {
  const profileDir = options.profileDir || PROFILE_DIR;
  const authFile = options.authFile || AUTH_FILE;
  let cleared = false;
  if (fs.existsSync(profileDir)) {
    fs.rmSync(profileDir, { recursive: true, force: true });
    cleared = true;
  }
  if (fs.existsSync(authFile)) {
    fs.rmSync(authFile);
    cleared = true;
  }
  return { cleared, authFile, profileDir };
}

module.exports = {
  AUTH_FILE,
  PROFILE_DIR,
  DATA_DIR,
  authenticate,
  authStatus,
  classifyToken,
  decodeJwtPayload,
  hasChatScopes,
  hasMailScopes,
  isLoginUrl,
  logout,
  readAuthFile,
};
