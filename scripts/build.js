#!/usr/bin/env node
'use strict';

const assert = require('node:assert');
const fs = require('node:fs');
const path = require('node:path');
const { renderSkillRouter } = require('../src/renderers');
const { capabilities } = require('../src/registry');

const repoRoot = path.join(__dirname, '..');
const packageJsonPath = path.join(repoRoot, 'package.json');
const skillPath = path.join(repoRoot, '.claude', 'skills', 'mg-api', 'SKILL.md');
const binPath = path.join(repoRoot, 'bin', 'mg-api.js');

function normalizeNewlines(text) {
  return text.replace(/\r\n/g, '\n');
}

function main() {
  const pkg = JSON.parse(fs.readFileSync(packageJsonPath, 'utf8'));
  assert.match(pkg.version, /^\d+\.\d+\.\d+(?:-[\w.]+)?$/, 'package version must be valid semver');
  assert.strictEqual(pkg.bin?.['mg-api'], 'bin/mg-api.js', 'package bin must expose mg-api');

  const bin = fs.readFileSync(binPath, 'utf8');
  assert.match(bin.split(/\r?\n/)[0], /node/, 'mg-api bin must have a node shebang');

  const skill = normalizeNewlines(fs.readFileSync(skillPath, 'utf8'));
  assert.strictEqual(skill, renderSkillRouter(), 'SKILL.md must match the registry-rendered router');

  for (const capability of ['auth', 'email', 'calendar', 'teams', 'chats', 'users', 'schema', 'doctor', 'update']) {
    assert.ok(capabilities[capability], `missing capability: ${capability}`);
  }

  process.stdout.write('Build validation passed.\n');
}

main();
