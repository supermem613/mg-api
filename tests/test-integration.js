#!/usr/bin/env node
'use strict';

const { describe, it, before } = require('node:test');
const assert = require('node:assert');
const { execFileSync } = require('node:child_process');
const { existsSync } = require('node:fs');
const { join } = require('node:path');
const { AUTH_FILE } = require('../src/graph-auth');

const repoRoot = join(__dirname, '..');
const cliPath = join(repoRoot, 'bin', 'mg-api.js');
const forceLogin = process.argv.includes('--login');

function runMgApi(args, options = {}) {
  const stdout = execFileSync(process.execPath, [cliPath, ...args], {
    cwd: repoRoot,
    encoding: 'utf8',
    timeout: options.timeout || 120000,
    stdio: ['pipe', 'pipe', options.inheritStderr ? 'inherit' : 'pipe'],
  });
  return JSON.parse(stdout);
}

if (!existsSync(AUTH_FILE) && !forceLogin) {
  console.log('Skipping integration tests because no Microsoft Graph auth is cached.');
  console.log('Authenticate first with: mg-api auth login');
  console.log('Or include --login to authenticate as part of the run.');
  process.exit(1);
}

describe('mg-api live Microsoft Graph integration', () => {
  before(() => {
    if (forceLogin) {
      runMgApi(['auth', 'login', '--force'], { timeout: 300000, inheritStderr: true });
    }
  });

  it('reports auth status through mg-api', () => {
    const status = runMgApi(['auth', 'status']);
    assert.strictEqual(status.ok, true);
    assert.strictEqual(status.data.exists, true);
    assert.strictEqual(status.data.hasGraphToken, true);
  });

  it('runs doctor without helper script checks', () => {
    const doctor = runMgApi(['doctor']);
    assert.strictEqual(doctor.ok, true);
    assert.ok(doctor.data.checks.some(check => check.name === 'auth-file'));
    assert.ok(!doctor.data.checks.some(check => check.name.includes('.js')));
  });

  it('reads the signed-in user profile through the semantic command surface', () => {
    const me = runMgApi(['users', 'me', '--select', 'displayName,mail']);
    assert.strictEqual(me.ok, true);
    assert.ok(me.data.displayName);
    assert.strictEqual(me.meta.token, 'graph');
    assert.strictEqual(me.meta.base, 'graph');
  });

  it('lists at least one mail message envelope shape', () => {
    const list = runMgApi(['email', 'list', '--top', '1', '--select', 'subject,from,receivedDateTime']);
    assert.strictEqual(list.ok, true);
    assert.ok(Array.isArray(list.data.value));
    assert.strictEqual(list.meta.token, 'outlook');
    assert.strictEqual(list.meta.base, 'outlook');
  });
});
