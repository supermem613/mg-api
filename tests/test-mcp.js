#!/usr/bin/env node
// MCP server static validation — no transport connection needed
// Run: node --test tests/test-mcp.js

const { describe, it } = require('node:test');
const assert = require('node:assert');
const { existsSync, readFileSync, readdirSync } = require('node:fs');
const { join } = require('node:path');

const ROOT = join(__dirname, '..');
const REFS = join(ROOT, 'references');
const SERVER = join(ROOT, 'src', 'mcp', 'server.js');
const PKG = JSON.parse(readFileSync(join(ROOT, 'package.json'), 'utf8'));

const EXPECTED_REF_FILES = ['email.md', 'calendar.md', 'teams.md', 'users.md', 'api-patterns.md'];
const TOOL_NAMES = ['graph_auth', 'graph_get', 'graph_post', 'graph_docs'];
const REF_KEYWORDS = {
  'email.md': ['messages', 'mail'],
  'calendar.md': ['calendar', 'events'],
  'teams.md': ['teams', 'chat'],
  'users.md': ['users', '/me'],
  'api-patterns.md': ['OData', 'filter'],
};

// ============================================================================
// 1. Server module loads without errors
// ============================================================================
describe('Server module', () => {
  it('server.js exists', () => {
    assert.ok(existsSync(SERVER), 'src/mcp/server.js not found');
  });

  it('server.js is valid JavaScript (syntax check)', () => {
    // Use Node's built-in syntax check — cheaper than require() which starts the server
    const { execSync } = require('node:child_process');
    const result = execSync(`node --check "${SERVER}"`, { encoding: 'utf8', stdio: ['pipe', 'pipe', 'pipe'] });
    assert.strictEqual(typeof result, 'string'); // no throw = pass
  });
});

// ============================================================================
// 2. References directory has all 5 expected files
// ============================================================================
describe('Reference files', () => {
  it('references/ directory exists', () => {
    assert.ok(existsSync(REFS), 'references/ directory not found');
  });

  it('has exactly 5 expected files', () => {
    const files = readdirSync(REFS).sort();
    assert.deepStrictEqual(files, [...EXPECTED_REF_FILES].sort(), `Expected ${EXPECTED_REF_FILES.join(', ')}`);
  });

  for (const f of EXPECTED_REF_FILES) {
    it(`references/${f} exists`, () => {
      assert.ok(existsSync(join(REFS, f)), `${f} not found`);
    });
  }
});

// ============================================================================
// 3. Each reference file is non-empty and contains expected keywords
// ============================================================================
describe('Reference file content', () => {
  for (const f of EXPECTED_REF_FILES) {
    const filePath = join(REFS, f);

    it(`${f} is non-empty`, () => {
      const content = readFileSync(filePath, 'utf8');
      assert.ok(content.trim().length > 0, `${f} is empty`);
    });

    const keywords = REF_KEYWORDS[f] || [];
    for (const kw of keywords) {
      it(`${f} contains keyword "${kw}"`, () => {
        const content = readFileSync(filePath, 'utf8');
        assert.ok(content.toLowerCase().includes(kw.toLowerCase()), `${f} should contain "${kw}"`);
      });
    }
  }
});

// ============================================================================
// 4. package.json has microsoft-graph-mcp bin entry
// ============================================================================
describe('package.json bin entry', () => {
  it('has "microsoft-graph-mcp" bin', () => {
    assert.ok(PKG.bin, 'package.json missing bin field');
    assert.ok(PKG.bin['microsoft-graph-mcp'], 'missing microsoft-graph-mcp bin entry');
  });

  it('bin points to src/mcp/server.js', () => {
    assert.match(PKG.bin['microsoft-graph-mcp'], /server\.js$/, 'bin should point to server.js');
  });

  it('test script runs all test files', () => {
    assert.ok(PKG.scripts['test'], 'missing test script');
    assert.match(PKG.scripts['test'], /test-mcp/, 'test script should include MCP tests');
  });
});

// ============================================================================
// 5. Server file has shebang line
// ============================================================================
describe('Server shebang', () => {
  const content = readFileSync(SERVER, 'utf8');

  it('has node shebang on first line', () => {
    const first = content.split(/\r?\n/)[0];
    assert.match(first, /^#!.*node/, 'server.js should have node shebang');
  });

  it('uses strict mode', () => {
    assert.match(content, /['"]use strict['"]/, 'server.js should use strict mode');
  });
});

// ============================================================================
// 6. Server file references all 4 tool names
// ============================================================================
describe('Tool registration', () => {
  const content = readFileSync(SERVER, 'utf8');

  for (const tool of TOOL_NAMES) {
    it(`registers tool "${tool}"`, () => {
      assert.ok(content.includes(`'${tool}'`), `server.js should register tool ${tool}`);
    });
  }

  it('uses McpServer from SDK', () => {
    assert.match(content, /McpServer/, 'should import McpServer');
  });

  it('uses StdioServerTransport', () => {
    assert.match(content, /StdioServerTransport/, 'should import StdioServerTransport');
  });
});

// ============================================================================
// 7. MCP resource registration
// ============================================================================
describe('Resource registration', () => {
  const content = readFileSync(SERVER, 'utf8');

  it('registers resources via server.resource()', () => {
    assert.match(content, /server\.resource\(/, 'server.js should register MCP resources');
  });

  it('uses graph://docs/ URI scheme', () => {
    assert.match(content, /graph:\/\/docs\//, 'resources should use graph://docs/ URI');
  });
});
