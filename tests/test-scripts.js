#!/usr/bin/env node
// Dry-run script validation — no network calls
// Run: node --test tests/test-scripts.js

const { describe, it } = require('node:test');
const assert = require('node:assert');
const { execSync } = require('node:child_process');
const { existsSync, readFileSync, readdirSync } = require('node:fs');
const { join } = require('node:path');

const ROOT = join(__dirname, '..');
const coreDir = join(ROOT, 'src', 'core');
const cliDir = join(ROOT, 'src', 'cli');
const srcDir = join(ROOT, 'src');

const CORE_SCRIPTS = ['mg-auth.js', 'mg-client.js', 'mg-env.js', 'mg-fetch.js'];
const CLI_SCRIPTS = ['mg-auth-cli.js', 'mg-get.js', 'mg-post.js'];
const ALL_SCRIPTS = [
  ...CORE_SCRIPTS.map(s => ({ name: s, dir: coreDir })),
  ...CLI_SCRIPTS.map(s => ({ name: s, dir: cliDir })),
];

/**
 * Run a Node.js script in a child process with a clean env.
 * Returns { exitCode, stdout, stderr }.
 */
function runScript(dir, scriptName, args = [], env = {}) {
  const scriptPath = join(dir, scriptName);
  const fakeHome = join(__dirname, '.test-home');
  const cleanEnv = {
    PATH: process.env.PATH,
    HOME: fakeHome,
    USERPROFILE: fakeHome,
    HOMEDRIVE: fakeHome.slice(0, 2),
    HOMEPATH: fakeHome.slice(2),
    SystemRoot: process.env.SystemRoot || '',
    ...env
  };
  try {
    const stdout = execSync(`node "${scriptPath}" ${args.map(a => `"${a}"`).join(' ')}`, {
      env: cleanEnv, encoding: 'utf8', timeout: 10000, stdio: ['pipe', 'pipe', 'pipe']
    });
    return { exitCode: 0, stdout, stderr: '' };
  } catch (e) {
    return { exitCode: e.status ?? 1, stdout: e.stdout ?? '', stderr: e.stderr ?? '' };
  }
}

/**
 * Read script source for static analysis tests.
 */
function readScript(dir, scriptName) {
  return readFileSync(join(dir, scriptName), 'utf8');
}

/** Recursively collect all file paths under a directory. */
function getAllFiles(dir) {
  const results = [];
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    const full = join(dir, entry.name);
    if (entry.isDirectory()) results.push(...getAllFiles(full));
    else results.push(full);
  }
  return results;
}

// ============================================================================
// 1. File existence
// ============================================================================
describe('File existence', () => {
  for (const s of CORE_SCRIPTS) {
    it(`src/core/${s} exists`, () => {
      assert.ok(existsSync(join(coreDir, s)), `${s} not found`);
    });
  }
  for (const s of CLI_SCRIPTS) {
    it(`src/cli/${s} exists`, () => {
      assert.ok(existsSync(join(cliDir, s)), `${s} not found`);
    });
  }
});

// ============================================================================
// 2. Shebang line
// ============================================================================
describe('Has shebang line', () => {
  for (const { name, dir } of ALL_SCRIPTS) {
    it(`${name} has node shebang`, () => {
      const first = readScript(dir, name).split(/\r?\n/)[0];
      assert.match(first, /^#!.*node/, `${name} should have node shebang`);
    });
  }
});

// ============================================================================
// 3. 'use strict'
// ============================================================================
describe('Uses strict mode', () => {
  for (const { name, dir } of ALL_SCRIPTS) {
    it(`${name} has 'use strict'`, () => {
      const content = readScript(dir, name);
      assert.match(content, /['"]use strict['"]/, `${name} should use strict mode`);
    });
  }
});

// ============================================================================
// 4. Error on missing arguments
// ============================================================================
describe('Error on missing arguments', () => {
  it('mg-get.js fails with no args', () => {
    const r = runScript(cliDir, 'mg-get.js');
    assert.notStrictEqual(r.exitCode, 0, 'Expected non-zero exit code');
  });

  it('mg-post.js fails with no args', () => {
    const r = runScript(cliDir, 'mg-post.js');
    assert.notStrictEqual(r.exitCode, 0, 'Expected non-zero exit code');
  });
});

// ============================================================================
// 5. Error on missing auth
// ============================================================================
describe('Error on missing auth', () => {
  it('mg-get.js with endpoint but no auth mentions mg-auth', () => {
    const r = runScript(cliDir, 'mg-get.js', ['/me/messages']);
    assert.notStrictEqual(r.exitCode, 0);
    assert.match(r.stderr, /mg-auth/, 'Error should mention mg-auth');
  });

  it('mg-post.js with endpoint but no auth mentions mg-auth', () => {
    const r = runScript(cliDir, 'mg-post.js', ['/me/events', '{}']);
    assert.notStrictEqual(r.exitCode, 0);
    assert.match(r.stderr, /mg-auth/, 'Error should mention mg-auth');
  });
});

// ============================================================================
// 6. mg-auth.js validation
// ============================================================================
describe('mg-auth.js validation', () => {
  const content = readScript(coreDir, 'mg-auth.js');

  it('uses playwright', () => {
    assert.match(content, /playwright/, 'mg-auth.js should use playwright');
  });

  it('handles --login flag', () => {
    assert.match(content, /--login/, 'mg-auth.js should handle --login flag');
  });

  it('handles --logout flag', () => {
    assert.match(content, /--logout/, 'mg-auth.js should handle --logout flag');
  });

  it('writes auth.json', () => {
    assert.match(content, /auth\.json/, 'mg-auth.js should write auth.json');
  });
});

// ============================================================================
// 7. mg-env.js validation
// ============================================================================
describe('mg-env.js validation', () => {
  const content = readScript(coreDir, 'mg-env.js');

  it('reads from auth.json', () => {
    assert.match(content, /auth\.json/, 'mg-env.js should read auth.json');
  });

  it('checks process.env', () => {
    assert.match(content, /process\.env/, 'mg-env.js should check process.env');
  });
});

// ============================================================================
// 8. mg-fetch.js validation
// ============================================================================
describe('mg-fetch.js validation', () => {
  const content = readScript(coreDir, 'mg-fetch.js');

  it('exports graphFetch function', () => {
    assert.match(content, /graphFetch/, 'mg-fetch.js should export graphFetch');
  });

  it('has retry logic', () => {
    assert.match(content, /RETRYABLE|retry/i, 'mg-fetch.js should have retry logic');
  });

  it('walks error cause chain', () => {
    assert.match(content, /\.cause/, 'mg-fetch.js should walk error cause chain');
  });
});

// ============================================================================
// 9. No shell scripts in src/
// ============================================================================
describe('No shell scripts in src/', () => {
  const allFiles = getAllFiles(srcDir);

  it('no .sh files', () => {
    const shFiles = allFiles.filter(f => f.endsWith('.sh'));
    assert.deepStrictEqual(shFiles, [], `Unexpected .sh files: ${shFiles.join(', ')}`);
  });

  it('no .ps1 files', () => {
    const ps1Files = allFiles.filter(f => f.endsWith('.ps1'));
    assert.deepStrictEqual(ps1Files, [], `Unexpected .ps1 files: ${ps1Files.join(', ')}`);
  });
});

// ============================================================================
// 10. All scripts are Node.js
// ============================================================================
describe('All scripts are Node.js', () => {
  it('every file in src/ is a .js file', () => {
    const allFiles = getAllFiles(srcDir);
    for (const f of allFiles) {
      assert.ok(f.endsWith('.js'), `Non-JS file found: ${f}`);
    }
  });
});

// ============================================================================
// 11. No external npm deps in CLI scripts
// ============================================================================
describe('No external npm dependencies in CLI scripts', () => {
  for (const s of ['mg-get.js', 'mg-post.js']) {
    it(`${s} only requires local modules`, () => {
      const content = readScript(cliDir, s);
      const requires = content.match(/require\s*\(\s*['"]([^'"]+)['"]\s*\)/g) || [];
      for (const r of requires) {
        assert.match(r, /['"]\.\.?\//, `${s} should only require local modules, found: ${r}`);
      }
    });
  }
});

// ============================================================================
// 12. SKILL.md exists
// ============================================================================
describe('SKILL.md exists', () => {
  it('SKILL.md in .github/skills/microsoft-graph/', () => {
    const skillMd = join(ROOT, '.github', 'skills', 'microsoft-graph', 'SKILL.md');
    assert.ok(existsSync(skillMd), 'SKILL.md not found in .github/skills/microsoft-graph/');
  });
});

// ============================================================================
// 13. Reference files exist
// ============================================================================
describe('Reference files exist', () => {
  const refDir = join(ROOT, 'references');
  const refFiles = ['email.md', 'calendar.md', 'teams.md', 'users.md', 'api-patterns.md'];

  if (existsSync(refDir)) {
    for (const f of refFiles) {
      it(`references/${f} exists`, () => {
        assert.ok(existsSync(join(refDir, f)), `references/${f} not found`);
      });
    }
  }
});
