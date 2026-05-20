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
const TEAMS_CHANNELS_URL = 'https://teams.cloud.microsoft/v2/teams';
const OFFICE_URL = 'https://www.office.com/';
const PAGE_TIMEOUT_MS = 20_000;
const NETWORK_IDLE_MS = 8_000;
const CHANNEL_PROBE_TIMEOUT_MS = 15_000;
const CHANNEL_MESSAGE_SCOPE_WARNING = 'ChannelMessage.Read.All was not observed during Teams channel probe. Teams channel ingest may fail.';
const TEAMS_CHANNEL_LINK_SELECTORS = [
  'a[href*="/l/channel/"]',
  'a[href*="/v2/channel/"]',
  'a[href*="/v2/channels/"]',
  '[data-tid*="channel"] a[href]',
];

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
  const chatScopes = parsed?.GRAPH_CHAT_SCOPES || [];
  const outlookChannelMessageScopes = parsed?.OUTLOOK_CHANNEL_MESSAGE_SCOPES || [];
  const channelMessageScopeObserved = hasChannelMessageScopes(chatScopes);
  const outlookChannelMessageScopeObserved = hasChannelMessageScopes(outlookChannelMessageScopes);
  return {
    authFile,
    exists: !!parsed,
    hasGraphToken: !!parsed?.GRAPH_TOKEN,
    hasOutlookToken: !!parsed?.OUTLOOK_TOKEN,
    hasChatToken: !!parsed?.GRAPH_CHAT_TOKEN,
    hasChannelMessageToken: !!parsed?.GRAPH_CHAT_TOKEN && channelMessageScopeObserved,
    hasOutlookChannelMessageToken: !!parsed?.OUTLOOK_CHANNEL_MESSAGE_TOKEN && outlookChannelMessageScopeObserved,
    channelMessageScopeObserved,
    outlookChannelMessageScopeObserved,
    teamsChannelProbe: parsed?.TEAMS_CHANNEL_PROBE || null,
    graphScopes: parsed?.GRAPH_SCOPES || [],
    outlookScopes: parsed?.OUTLOOK_SCOPES || [],
    outlookChannelMessageScopes,
    chatScopes,
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

async function settlePage(page) {
  await page.waitForLoadState('networkidle', { timeout: NETWORK_IDLE_MS }).catch(() => {});
}

async function clickFirstVisible(locator) {
  const count = await locator.count();
  for (let i = 0; i < count; i++) {
    const candidate = locator.nth(i);
    if (!candidate.isVisible || await candidate.isVisible().catch(() => false)) {
      await candidate.click();
      return true;
    }
  }
  return false;
}

async function waitForChannelMessageScope(page, captureTokenFromAuthorization, log) {
  if (typeof page.waitForRequest !== 'function') return false;
  try {
    await page.waitForRequest(request => {
      const auth = request.headers()['authorization'];
      const info = captureTokenFromAuthorization(auth);
      return hasChannelMessageScopes(info?.scopes || []);
    }, { timeout: CHANNEL_PROBE_TIMEOUT_MS });
    return true;
  } catch {
    log('Teams channel probe did not observe ChannelMessage.Read.All before timeout');
    return false;
  }
}

function buildTeamsChannelUrl(team, channel) {
  if (channel.webUrl) return channel.webUrl;
  const channelName = encodeURIComponent(channel.displayName || 'General');
  return `https://teams.microsoft.com/l/channel/${encodeURIComponent(channel.id)}/${channelName}?groupId=${encodeURIComponent(team.id)}`;
}

async function graphGetJson(fetchImpl, token, url, log) {
  const response = await fetchImpl(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!response.ok) {
    log(`Graph discovery request failed with HTTP ${response.status}: ${url}`);
    return null;
  }
  return response.json();
}

async function discoverTeamsChannel(fetchImpl, graphToken, log) {
  if (!fetchImpl || !graphToken) return null;
  const teams = await graphGetJson(fetchImpl, graphToken, 'https://graph.microsoft.com/v1.0/me/joinedTeams', log).catch(err => {
    log(`Teams discovery failed (continuing): ${err.message}`);
    return null;
  });
  for (const team of teams?.value || []) {
    if (!team.id) continue;
    const channelsUrl = `https://graph.microsoft.com/v1.0/teams/${encodeURIComponent(team.id)}/channels`;
    const channels = await graphGetJson(fetchImpl, graphToken, channelsUrl, log).catch(err => {
      log(`Teams channel discovery failed for a joined team (continuing): ${err.message}`);
      return null;
    });
    const channel = (channels?.value || []).find(candidate => candidate.id && (candidate.webUrl || candidate.displayName));
    if (channel) return buildTeamsChannelUrl(team, channel);
  }
  return null;
}

async function openDiscoveredTeamsChannel(page, fetchImpl, graphToken, log) {
  const channelUrl = await discoverTeamsChannel(fetchImpl, graphToken, log);
  if (!channelUrl) {
    log('Teams channel probe could not discover a joined team channel to open');
    return false;
  }
  await page.goto(channelUrl, { waitUntil: 'domcontentloaded', timeout: PAGE_TIMEOUT_MS }).catch(err => log(`Teams discovered channel goto failed (continuing): ${err.message}`));
  return true;
}

async function probeTeamsChannelMessages(page, options) {
  const {
    captureTokenFromAuthorization,
    fetchImpl,
    graphToken,
    hasChannelMessageToken,
    log,
  } = options;
  log('visiting Teams channels surface to capture channel message scopes');
  await page.goto(TEAMS_CHANNELS_URL, { waitUntil: 'domcontentloaded', timeout: PAGE_TIMEOUT_MS }).catch(err => log(`Teams channels goto failed (continuing): ${err.message}`));
  await settlePage(page);
  if (hasChannelMessageToken()) return { attempted: true, opened: false, observed: true };

  if (typeof page.locator === 'function') {
    for (const selector of TEAMS_CHANNEL_LINK_SELECTORS) {
      const locator = page.locator(selector);
      const opened = await clickFirstVisible(locator).catch(err => {
        log(`Teams channel selector ${selector} failed (continuing): ${err.message}`);
        return false;
      });
      if (!opened) continue;
      if (!hasChannelMessageToken()) await waitForChannelMessageScope(page, captureTokenFromAuthorization, log);
      await settlePage(page);
      return { attempted: true, opened: true, observed: hasChannelMessageToken() };
    }
    log('Teams channel probe did not find a channel link to open');
  } else {
    log('Teams channel probe selector pass skipped: Playwright locator API unavailable');
  }

  const openedDiscoveredChannel = await openDiscoveredTeamsChannel(page, fetchImpl, graphToken, log);
  if (openedDiscoveredChannel) {
    if (!hasChannelMessageToken()) await waitForChannelMessageScope(page, captureTokenFromAuthorization, log);
    await settlePage(page);
    return { attempted: true, opened: true, observed: hasChannelMessageToken() };
  }

  return { attempted: true, opened: false, observed: false };
}

async function authenticate(options = {}) {
  const {
    forceLogin = false,
    verbose = false,
    playwright = requirePlaywright(),
    fetch = globalThis.fetch,
    authFile = AUTH_FILE,
    profileDir = PROFILE_DIR,
  } = options;

  ensureDir(profileDir);

  const captured = {
    graph: null, graphScopes: [],
    graphChat: null, graphChatScopes: [],
    outlook: null, outlookScopes: [],
    outlookChannelMessage: null, outlookChannelMessageScopes: [],
    channelProbe: { attempted: false, opened: false, observed: false },
  };

  function captureTokenFromAuthorization(auth) {
    if (!auth?.startsWith('Bearer ')) return null;
    const token = auth.substring(7);
    const info = classifyToken(token);
    if (!info) return null;
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
    if (info.type === 'outlook'
      && hasChannelMessageScopes(info.scopes)
      && compareTeamsTokenScopes(info.scopes, captured.outlookChannelMessageScopes) > 0) {
      captured.outlookChannelMessage = token;
      captured.outlookChannelMessageScopes = info.scopes;
    }
    return info;
  }

  function installTokenInterceptor(p) {
    p.on('request', request => captureTokenFromAuthorization(request.headers()['authorization']));
  }

  function installCapture(context) {
    if (typeof context.on === 'function') {
      installTokenInterceptor(context);
      return true;
    }
    for (const existingPage of context.pages()) installTokenInterceptor(existingPage);
    return false;
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
  let captureInstalledOnContext = installCapture(context);
  let page = context.pages()[0] || await context.newPage();
  if (!captureInstalledOnContext && context.pages()[0] !== page) installTokenInterceptor(page);

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
      captureInstalledOnContext = installCapture(context);
      page = context.pages()[0] || await context.newPage();
      if (!captureInstalledOnContext && context.pages()[0] !== page) installTokenInterceptor(page);
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
  await settlePage(page);

  log('visiting Teams to capture chat scopes');
  const teamsPage = await context.newPage();
  if (!captureInstalledOnContext) installTokenInterceptor(teamsPage);
  await teamsPage.goto(TEAMS_URL, { waitUntil: 'domcontentloaded', timeout: PAGE_TIMEOUT_MS }).catch(err => log(`Teams goto failed (continuing): ${err.message}`));
  await settlePage(teamsPage);
  await teamsPage.goto(TEAMS_CHAT_URL, { waitUntil: 'domcontentloaded', timeout: PAGE_TIMEOUT_MS }).catch(err => log(`Teams chat goto failed (continuing): ${err.message}`));
  await settlePage(teamsPage);
  captured.channelProbe = await probeTeamsChannelMessages(teamsPage, {
    captureTokenFromAuthorization,
    fetchImpl: fetch,
    graphToken: captured.graph || captured.graphChat,
    hasChannelMessageToken: () => hasChannelMessageScopes(captured.graphChatScopes),
    log,
  });

  log('visiting office.com to capture Outlook scopes');
  const officePage = await context.newPage();
  if (!captureInstalledOnContext) installTokenInterceptor(officePage);
  await officePage.goto(OFFICE_URL, { waitUntil: 'domcontentloaded', timeout: PAGE_TIMEOUT_MS }).catch(err => log(`Office goto failed (continuing): ${err.message}`));
  await settlePage(officePage);

  const channelMessageScopeObserved = hasChannelMessageScopes(captured.graphChatScopes);
  const outlookChannelMessageScopeObserved = hasChannelMessageScopes(captured.outlookChannelMessageScopes);
  log(`captured: graph=${!!captured.graph} chat=${!!captured.graphChat} channelMessage=${channelMessageScopeObserved} outlook=${!!captured.outlook} outlookChannelMessage=${outlookChannelMessageScopeObserved}`);
  await context.close();

  if (!captured.graph && !captured.outlook) {
    throw new Error('No tokens captured. Run mg-api auth login --force to force a fresh login.');
  }

  const authData = {
    ...(captured.graph && { GRAPH_TOKEN: captured.graph }),
    ...(captured.graphChat && { GRAPH_CHAT_TOKEN: captured.graphChat }),
    ...(captured.outlook && { OUTLOOK_TOKEN: captured.outlook }),
    ...(captured.outlookChannelMessage && { OUTLOOK_CHANNEL_MESSAGE_TOKEN: captured.outlookChannelMessage }),
    ...(captured.graphScopes.length && { GRAPH_SCOPES: captured.graphScopes }),
    ...(captured.graphChatScopes.length && { GRAPH_CHAT_SCOPES: captured.graphChatScopes }),
    ...(captured.outlookScopes.length && { OUTLOOK_SCOPES: captured.outlookScopes }),
    ...(captured.outlookChannelMessageScopes.length && { OUTLOOK_CHANNEL_MESSAGE_SCOPES: captured.outlookChannelMessageScopes }),
    CHANNEL_MESSAGE_SCOPE_OBSERVED: channelMessageScopeObserved,
    OUTLOOK_CHANNEL_MESSAGE_SCOPE_OBSERVED: outlookChannelMessageScopeObserved,
    TEAMS_CHANNEL_PROBE: captured.channelProbe,
  };

  ensureDir(path.dirname(authFile));
  fs.writeFileSync(authFile, JSON.stringify(authData, null, 2) + '\n');
  if (!channelMessageScopeObserved) process.stderr.write(`${CHANNEL_MESSAGE_SCOPE_WARNING}\n`);
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
  hasChannelMessageScopes,
  hasMailScopes,
  isLoginUrl,
  logout,
  readAuthFile,
};
