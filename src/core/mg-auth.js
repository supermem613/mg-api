#!/usr/bin/env node
// ============================================================================
// mg-auth.js — Playwright-based Microsoft Graph authentication
// ============================================================================
// Usage:
//   node mg-auth.js [--login] [--logout] [--help]
//
// First run: opens Edge for login (one-time)
// Subsequent runs: headless, uses cached profile (instant)
//
// Navigates to Outlook and Teams to capture Graph + Outlook bearer tokens.
// Auth is saved to ~/.microsoft-graph-skill/auth.json.
// ============================================================================
'use strict';

let chromium;
try {
  ({ chromium } = require('playwright'));
} catch {
  process.stderr.write('ERROR: playwright is not installed.\nRun: npm install    (in the skill directory)\n');
  process.exit(1);
}
const fs = require('fs');
const path = require('path');
const os = require('os');

const DATA_DIR = path.join(os.homedir(), '.microsoft-graph-skill');
const PROFILE_DIR = path.join(DATA_DIR, 'browser-profile');
const AUTH_FILE = path.join(DATA_DIR, 'auth.json');
const LOGIN_TIMEOUT_MS = 300_000;
const HEADLESS_PROBE_MS = 5_000;

const OUTLOOK_URL = 'https://outlook.office.com/mail/';
const TEAMS_URL = 'https://teams.microsoft.com/';

// ── Helpers ──────────────────────────────────────────────────────────────────

function ensureDir(dir) {
  fs.mkdirSync(dir, { recursive: true });
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

// ── Core auth flow ───────────────────────────────────────────────────────────

async function authenticate({ forceLogin = false } = {}) {
  ensureDir(PROFILE_DIR);

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
        const hasMail = info.scopes.some(s => s.toLowerCase() === 'mail.read' || s.toLowerCase() === 'mail.readwrite');
        const hasChat = info.scopes.some(s => s.toLowerCase() === 'chat.read' || s.toLowerCase() === 'chat.readwrite');
        // Prefer token with Mail scopes for general use
        if (hasMail && !captured.graphScopes.some(s => s.toLowerCase() === 'mail.read' || s.toLowerCase() === 'mail.readwrite')) {
          captured.graph = token;
          captured.graphScopes = info.scopes;
        } else if (info.scopes.length > captured.graphScopes.length) {
          captured.graph = token;
          captured.graphScopes = info.scopes;
        }
        // Keep separate chat token
        if (hasChat && info.scopes.length > captured.graphChatScopes.length) {
          captured.graphChat = token;
          captured.graphChatScopes = info.scopes;
        }
      }
      if (info.type === 'outlook' && info.scopes.length > captured.outlookScopes.length) {
        captured.outlook = token;
        captured.outlookScopes = info.scopes;
      }
    });
  }

  // First attempt: headless (unless --login forces visible)
  let headless = !forceLogin;
  let context = await chromium.launchPersistentContext(PROFILE_DIR, {
    channel: 'msedge',
    headless,
    args: ['--disable-blink-features=AutomationControlled'],
    viewport: { width: 1280, height: 800 },
  });

  let page = context.pages()[0] || await context.newPage();
  installTokenInterceptor(page);

  await page.goto(OUTLOOK_URL, { waitUntil: 'domcontentloaded' });

  // Check if we landed on a login page
  if (headless && isLoginUrl(page.url())) {
    try {
      await page.waitForURL(url => !isLoginUrl(url.toString()), { timeout: HEADLESS_PROBE_MS });
    } catch {
      await context.close();
      process.stderr.write('🔑 Opening Edge for login (one-time)...\n');
      process.stderr.write('   Complete the login in the browser window.\n');

      context = await chromium.launchPersistentContext(PROFILE_DIR, {
        channel: 'msedge',
        headless: false,
        args: ['--disable-blink-features=AutomationControlled'],
        viewport: { width: 1280, height: 800 },
      });

      page = context.pages()[0] || await context.newPage();
      installTokenInterceptor(page);
      await page.goto(OUTLOOK_URL, { waitUntil: 'domcontentloaded' });
      headless = false;
    }
  }

  // If visible, wait for user to complete login
  if (!headless && isLoginUrl(page.url())) {
    try {
      await page.waitForURL(url => !isLoginUrl(url.toString()), { timeout: LOGIN_TIMEOUT_MS });
      await page.waitForTimeout(2000);
      process.stderr.write('✅ Login successful. Profile saved.\n');
    } catch {
      process.stderr.write('⚠️  Login timed out or browser was closed.\n');
      await context.close();
      process.exit(1);
    }
  }

  if (forceLogin && !isLoginUrl(page.url())) {
    await page.waitForTimeout(1000);
    process.stderr.write('✅ Login successful. Profile saved.\n');
  }

  // Wait for Outlook to fully load and fire API calls
  await page.waitForLoadState('networkidle').catch(() => {});
  await page.waitForTimeout(5000);

  // Open Teams in a new tab to capture Graph tokens with Mail/Calendar scopes
  const teamsPage = await context.newPage();
  installTokenInterceptor(teamsPage);
  await teamsPage.goto(TEAMS_URL, { waitUntil: 'domcontentloaded' });
  await teamsPage.waitForLoadState('networkidle').catch(() => {});
  await teamsPage.waitForTimeout(5000);

  // Navigate Teams to chat section to trigger Chat-scoped Graph tokens
  await teamsPage.goto('https://teams.microsoft.com/v2/chat', { waitUntil: 'domcontentloaded' }).catch(() => {});
  await teamsPage.waitForLoadState('networkidle').catch(() => {});
  await teamsPage.waitForTimeout(3000);

  // Navigate to office.com — its Graph token has Chat.Read
  const officePage = await context.newPage();
  installTokenInterceptor(officePage);
  await officePage.goto('https://www.office.com/', { waitUntil: 'domcontentloaded' });
  await officePage.waitForLoadState('networkidle').catch(() => {});
  await officePage.waitForTimeout(3000);

  await context.close();

  if (!captured.graph && !captured.outlook) {
    process.stderr.write('ERROR: No tokens captured.\n');
    process.stderr.write('Try running with --login to force a fresh login.\n');
    process.exit(1);
  }

  // Persist auth
  const authData = {
    ...(captured.graph && { GRAPH_TOKEN: captured.graph }),
    ...(captured.graphChat && { GRAPH_CHAT_TOKEN: captured.graphChat }),
    ...(captured.outlook && { OUTLOOK_TOKEN: captured.outlook }),
    ...(captured.graphScopes.length && { GRAPH_SCOPES: captured.graphScopes }),
    ...(captured.graphChatScopes.length && { GRAPH_CHAT_SCOPES: captured.graphChatScopes }),
    ...(captured.outlookScopes.length && { OUTLOOK_SCOPES: captured.outlookScopes }),
  };
  ensureDir(DATA_DIR);
  fs.writeFileSync(AUTH_FILE, JSON.stringify(authData, null, 2) + '\n');

  return authData;
}

// ── Logout ────────────────────────────────────────────────────────────────────

function logout() {
  let cleared = false;
  if (fs.existsSync(PROFILE_DIR)) {
    fs.rmSync(PROFILE_DIR, { recursive: true, force: true });
    cleared = true;
  }
  if (fs.existsSync(AUTH_FILE)) {
    fs.rmSync(AUTH_FILE);
    cleared = true;
  }
  process.stderr.write(cleared ? '🗑️  Browser profile and auth cleared.\n' : 'ℹ️  No profile found.\n');
}

// ── Exports ──────────────────────────────────────────────────────────────────

module.exports = { authenticate, logout, AUTH_FILE, DATA_DIR, PROFILE_DIR };

// ── CLI entrypoint (only when run directly) ──────────────────────────────────

if (require.main === module) {
  (async () => {
    const args = process.argv.slice(2);
    const flags = new Set(args.filter(a => a.startsWith('--')));

    if (flags.has('--help')) {
      process.stderr.write(`Usage: node mg-auth.js [--login] [--logout]\n`);
      process.stderr.write(`\n`);
      process.stderr.write(`  --login     Force visible browser for re-login\n`);
      process.stderr.write(`  --logout    Clear saved browser profile and tokens\n`);
      process.stderr.write(`\n`);
      process.stderr.write(`Navigates to Outlook and Teams to capture Graph + Outlook tokens.\n`);
      process.stderr.write(`Auth is saved to ~/.microsoft-graph-skill/auth.json.\n`);
      process.exit(0);
    }

    if (flags.has('--logout')) {
      logout();
      process.exit(0);
    }

    await authenticate({ forceLogin: flags.has('--login') });

    process.stderr.write(`✅ Graph authentication complete\n`);
    process.stderr.write(`   Auth saved to ${AUTH_FILE}\n`);
  })().catch(err => {
    process.stderr.write(`ERROR: ${err.message}\n`);
    process.exit(1);
  });
}
