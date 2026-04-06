#!/usr/bin/env node
// ============================================================================
// mg-auth-cli.js — CLI entrypoint for Microsoft Graph authentication
// ============================================================================
// Usage:  node mg-auth-cli.js [--login] [--logout] [--help]
//
// Thin wrapper — delegates to ../core/mg-auth.js for the actual auth flow.
// ============================================================================
'use strict';

const { authenticate, logout, AUTH_FILE } = require('../core/mg-auth');

const args = process.argv.slice(2);
const flags = new Set(args.filter(a => a.startsWith('--')));

if (flags.has('--help')) {
  process.stderr.write('Usage: node mg-auth-cli.js [--login] [--logout]\n');
  process.stderr.write('\n');
  process.stderr.write('  --login     Force visible browser for re-login\n');
  process.stderr.write('  --logout    Clear saved browser profile and tokens\n');
  process.stderr.write('\n');
  process.stderr.write('Captures Graph + Outlook tokens via Playwright.\n');
  process.stderr.write(`Auth saved to ${AUTH_FILE}\n`);
  process.exit(0);
}

(async () => {
  try {
    if (flags.has('--logout')) {
      logout();
      process.exit(0);
    }

    await authenticate({ forceLogin: flags.has('--login') });
    process.stderr.write(`✅ Auth saved to ${AUTH_FILE}\n`);
  } catch (err) {
    process.stderr.write(`ERROR: ${err.message}\n`);
    process.exit(1);
  }
})();
