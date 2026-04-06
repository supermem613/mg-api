#!/usr/bin/env node
// MCP tool integration tests — validates tool registration, schemas, and handlers.
// These tests catch the class of bug where server.tool() arg order is wrong,
// causing empty inputSchemas and non-callable handlers ("typedHandler is not a function").
//
// Run: node --test tests/test-mcp-tools.js

'use strict';

const { describe, it, before } = require('node:test');
const assert = require('node:assert');
const { join } = require('node:path');
const { readFileSync } = require('node:fs');

const ROOT = join(__dirname, '..');
const SERVER = join(ROOT, 'src', 'mcp', 'server.js');

// ============================================================================
// Helpers: build McpServer instances with tools registered exactly as in
// server.js — but WITHOUT connecting a transport, so we can inspect internals.
// ============================================================================

async function buildServer() {
  const { McpServer } = await import('@modelcontextprotocol/sdk/server/mcp.js');
  const { z } = await import('zod/v3');

  const server = new McpServer({ name: 'microsoft-graph', version: '1.0.0' });
  const TOPICS = { email: 'email.md', calendar: 'calendar.md', teams: 'teams.md', users: 'users.md', patterns: 'api-patterns.md' };
  const REFS = join(ROOT, 'references');

  server.tool('graph_auth', 'Authenticate to Microsoft Graph via browser sign-in',
    { login: z.boolean().optional().describe('Force visible browser login'), logout: z.boolean().optional().describe('Clear saved tokens') },
    async ({ login, logout: doLogout }) => ({ content: [{ type: 'text', text: 'stub' }] }));

  // CORRECT order: schema, annotations, handler
  server.tool('graph_get', 'Read data from Microsoft Graph API (any GET endpoint)',
    { endpoint: z.string().describe('Graph API path, e.g. /me/messages'), params: z.record(z.string()).optional().describe('OData query params (top, filter, select, orderby)') },
    { readOnlyHint: true, destructiveHint: false },
    async ({ endpoint, params }) => ({ content: [{ type: 'text', text: 'stub' }] }));

  server.tool('graph_post', 'Write data to Microsoft Graph API (POST/PATCH/DELETE)',
    { method: z.enum(['POST', 'PATCH', 'DELETE']).describe('HTTP method'), endpoint: z.string().describe('API path, e.g. /me/events'), body: z.record(z.unknown()).optional().describe('JSON request body'), useOutlookToken: z.boolean().optional().describe('Use Outlook token + outlook.office.com base URL') },
    { readOnlyHint: false },
    async ({ method, endpoint, body, useOutlookToken }) => ({ content: [{ type: 'text', text: 'stub' }] }));

  server.tool('graph_docs', 'Get Graph API reference documentation for a topic',
    { topic: z.enum(['email', 'calendar', 'teams', 'users', 'patterns']).describe('Documentation topic') },
    { readOnlyHint: true },
    async ({ topic }) => {
      const file = join(REFS, TOPICS[topic]);
      try { return { content: [{ type: 'text', text: readFileSync(file, 'utf8') }] }; }
      catch { return { content: [{ type: 'text', text: `ERROR: ${file} not found` }] }; }
    });

  return server;
}

// Build a server with the BUGGY arg order (annotations before schema).
async function buildBuggyServer() {
  const { McpServer } = await import('@modelcontextprotocol/sdk/server/mcp.js');
  const { z } = await import('zod/v3');

  const server = new McpServer({ name: 'buggy', version: '1.0.0' });

  // BAD order: annotations, schema, handler
  server.tool('graph_get_bad', 'Read data',
    { readOnlyHint: true, destructiveHint: false },
    { endpoint: z.string(), params: z.record(z.string()).optional() },
    async ({ endpoint }) => ({ content: [{ type: 'text', text: 'ok' }] }));

  server.tool('graph_docs_bad', 'Get docs',
    { readOnlyHint: true },
    { topic: z.enum(['email', 'calendar']) },
    async ({ topic }) => ({ content: [{ type: 'text', text: 'ok' }] }));

  return server;
}

let server;
let buggyServer;

before(async () => {
  server = await buildServer();
  buggyServer = await buildBuggyServer();
});

// ============================================================================
// 1. All 4 tools are registered
// ============================================================================
describe('Tool discovery', () => {
  it('registers exactly 4 tools', () => {
    const names = Object.keys(server._registeredTools);
    assert.strictEqual(names.length, 4);
  });

  for (const name of ['graph_auth', 'graph_get', 'graph_post', 'graph_docs']) {
    it(`registers tool "${name}"`, () => {
      assert.ok(server._registeredTools[name], `tool ${name} not found`);
    });
  }
});

// ============================================================================
// 2. Every tool handler is a callable function (not a Zod schema object).
//    This is the primary regression test for the arg-swap bug:
//    when annotations come before the schema, the SDK silently assigns the
//    Zod schema as the handler and drops the real async function.
// ============================================================================
describe('Tool handlers are callable functions', () => {
  for (const name of ['graph_auth', 'graph_get', 'graph_post', 'graph_docs']) {
    it(`${name} handler is a function`, () => {
      const tool = server._registeredTools[name];
      assert.strictEqual(typeof tool.handler, 'function',
        `${name} handler should be a function, got ${typeof tool.handler}. ` +
        'This usually means server.tool() args are misordered (annotations before schema).');
    });
  }
});

// ============================================================================
// 3. Input schemas are properly registered (not undefined or empty).
//    When args are swapped, inputSchema becomes undefined because the SDK
//    treats the annotations as "not a Zod shape" and skips the schema entirely.
// ============================================================================
describe('Input schemas are registered', () => {
  for (const name of ['graph_auth', 'graph_get', 'graph_post', 'graph_docs']) {
    it(`${name} has a non-null inputSchema`, () => {
      const tool = server._registeredTools[name];
      assert.ok(tool.inputSchema, `${name} inputSchema is missing — args may be misordered`);
    });
  }
});

// ============================================================================
// 4. Specific schema properties for each tool
// ============================================================================
describe('graph_auth schema', () => {
  it('has "login" and "logout" fields', async () => {
    const { getObjectShape } = await import('@modelcontextprotocol/sdk/server/zod-compat.js');
    const shape = getObjectShape(server._registeredTools['graph_auth'].inputSchema);
    assert.ok(shape, 'could not extract shape');
    assert.ok(shape.login, 'missing "login" field');
    assert.ok(shape.logout, 'missing "logout" field');
  });
});

describe('graph_get schema', () => {
  it('has "endpoint" and "params" fields', async () => {
    const { getObjectShape } = await import('@modelcontextprotocol/sdk/server/zod-compat.js');
    const shape = getObjectShape(server._registeredTools['graph_get'].inputSchema);
    assert.ok(shape, 'could not extract shape');
    assert.ok(shape.endpoint, 'missing "endpoint" field');
    assert.ok(shape.params, 'missing "params" field');
  });
});

describe('graph_post schema', () => {
  it('has "method", "endpoint", "body", and "useOutlookToken" fields', async () => {
    const { getObjectShape } = await import('@modelcontextprotocol/sdk/server/zod-compat.js');
    const shape = getObjectShape(server._registeredTools['graph_post'].inputSchema);
    assert.ok(shape, 'could not extract shape');
    assert.ok(shape.method, 'missing "method" field');
    assert.ok(shape.endpoint, 'missing "endpoint" field');
    assert.ok(shape.body, 'missing "body" field');
    assert.ok(shape.useOutlookToken, 'missing "useOutlookToken" field');
  });
});

describe('graph_docs schema', () => {
  it('has "topic" field', async () => {
    const { getObjectShape } = await import('@modelcontextprotocol/sdk/server/zod-compat.js');
    const shape = getObjectShape(server._registeredTools['graph_docs'].inputSchema);
    assert.ok(shape, 'could not extract shape');
    assert.ok(shape.topic, 'missing "topic" field');
  });
});

// ============================================================================
// 5. Annotations are correctly attached
// ============================================================================
describe('Tool annotations', () => {
  it('graph_get has readOnlyHint: true, destructiveHint: false', () => {
    const ann = server._registeredTools['graph_get'].annotations;
    assert.ok(ann, 'annotations missing');
    assert.strictEqual(ann.readOnlyHint, true);
    assert.strictEqual(ann.destructiveHint, false);
  });

  it('graph_post has readOnlyHint: false', () => {
    const ann = server._registeredTools['graph_post'].annotations;
    assert.ok(ann, 'annotations missing');
    assert.strictEqual(ann.readOnlyHint, false);
  });

  it('graph_docs has readOnlyHint: true', () => {
    const ann = server._registeredTools['graph_docs'].annotations;
    assert.ok(ann, 'annotations missing');
    assert.strictEqual(ann.readOnlyHint, true);
  });

  it('graph_auth has no annotations', () => {
    const ann = server._registeredTools['graph_auth'].annotations;
    assert.ok(!ann, 'graph_auth should have no annotations');
  });
});

// ============================================================================
// 6. Tool handler invocation — call the handler and verify it returns content.
//    The arg-swap bug causes "typedHandler is not a function" because the
//    Zod schema object is assigned as the handler.
// ============================================================================
describe('graph_docs handler returns docs', () => {
  it('returns non-empty markdown for "email"', async () => {
    const tool = server._registeredTools['graph_docs'];
    const result = await tool.handler({ topic: 'email' }, {});
    assert.ok(result.content, 'result should have content');
    assert.ok(result.content[0].text.length > 100, 'should return substantial docs');
    assert.ok(result.content[0].text.toLowerCase().includes('mail'), 'email docs should mention mail');
  });

  for (const topic of ['email', 'calendar', 'teams', 'users', 'patterns']) {
    it(`returns docs for "${topic}"`, async () => {
      const tool = server._registeredTools['graph_docs'];
      const result = await tool.handler({ topic }, {});
      assert.ok(result.content[0].text.length > 0, `${topic} docs should be non-empty`);
    });
  }
});

// ============================================================================
// 7. Regression: buggy arg order produces broken tools
//    These tests verify that the WRONG order causes exactly the symptoms
//    the user reported: handler is not a function, inputSchema is undefined.
//    If these tests ever pass with the buggy server, the SDK changed its
//    parsing logic and these guard-tests should be updated.
// ============================================================================
describe('Buggy arg order (annotations before schema) produces broken tools', () => {
  it('buggy handler is NOT a function (it is the Zod schema object)', () => {
    const tool = buggyServer._registeredTools['graph_get_bad'];
    assert.notStrictEqual(typeof tool.handler, 'function',
      'With the buggy arg order, the handler should NOT be a function — ' +
      'it should be the Zod schema object. If this test fails, the SDK ' +
      'may have changed its arg-parsing logic.');
  });

  it('buggy inputSchema is undefined', () => {
    const tool = buggyServer._registeredTools['graph_get_bad'];
    assert.ok(!tool.inputSchema,
      'With the buggy arg order, inputSchema should be missing. ' +
      'If this test fails, the SDK may have changed its arg-parsing logic.');
  });

  it('buggy handler object has schema keys instead of being callable', () => {
    const tool = buggyServer._registeredTools['graph_get_bad'];
    assert.ok(typeof tool.handler === 'object' && tool.handler !== null,
      'handler should be an object (the schema)');
    assert.ok('endpoint' in tool.handler, 'handler should contain "endpoint" key from the schema');
  });
});

// ============================================================================
// 8. Source-level check: server.js arg order matches the SDK's expected order.
//    The SDK expects: server.tool(name, description, schema, annotations, cb)
//    Verify that for each tool with annotations, the schema arg appears BEFORE
//    the annotations arg (not after).
// ============================================================================
describe('server.js source-level arg order', () => {
  const src = readFileSync(SERVER, 'utf8');

  it('graph_get: schema before annotations', () => {
    const schemaIdx = src.indexOf("{ endpoint: z.string()");
    const annotIdx = src.indexOf("{ readOnlyHint: true, destructiveHint: false }");
    assert.ok(schemaIdx > 0 && annotIdx > 0, 'could not find graph_get registration');
    assert.ok(schemaIdx < annotIdx, 'graph_get: schema must appear before annotations');
  });

  it('graph_post: schema before annotations', () => {
    const schemaIdx = src.indexOf("method: z.enum(['POST', 'PATCH', 'DELETE'])");
    const annotIdx = src.indexOf("{ readOnlyHint: false }");
    assert.ok(schemaIdx > 0 && annotIdx > 0, 'could not find graph_post registration');
    assert.ok(schemaIdx < annotIdx, 'graph_post: schema must appear before annotations');
  });

  it('graph_docs: schema before annotations', () => {
    const toolIdx = src.indexOf("'graph_docs'");
    const schemaIdx = src.indexOf("z.enum(['email', 'calendar', 'teams', 'users', 'patterns'])", toolIdx);
    const annotIdxStart = src.indexOf("{ readOnlyHint: true }", toolIdx);
    assert.ok(toolIdx > 0 && schemaIdx > 0 && annotIdxStart > 0, 'could not find graph_docs registration');
    assert.ok(schemaIdx < annotIdxStart, 'graph_docs: schema must appear before annotations');
  });
});

// ============================================================================
// 9. Schema-to-JSON conversion doesn't crash
//    When the bug was present, normalizeObjectSchema() on the stored
//    inputSchema would fail because the schema was undefined.
// ============================================================================
describe('Schema serialization to JSON Schema', () => {
  for (const name of ['graph_auth', 'graph_get', 'graph_post', 'graph_docs']) {
    it(`${name} inputSchema converts to JSON Schema without error`, async () => {
      const { normalizeObjectSchema } = await import('@modelcontextprotocol/sdk/server/zod-compat.js');
      const { toJsonSchemaCompat } = await import('@modelcontextprotocol/sdk/server/zod-json-schema-compat.js');
      const tool = server._registeredTools[name];
      const obj = normalizeObjectSchema(tool.inputSchema);
      assert.ok(obj, `${name}: normalizeObjectSchema returned falsy`);
      const jsonSchema = toJsonSchemaCompat(obj, { strictUnions: true, pipeStrategy: 'input' });
      assert.ok(jsonSchema, `${name}: toJsonSchemaCompat returned falsy`);
      assert.strictEqual(jsonSchema.type, 'object', 'JSON Schema type should be "object"');
      assert.ok(jsonSchema.properties, 'JSON Schema should have properties');
    });
  }

  it('graph_get JSON Schema has endpoint and params properties', async () => {
    const { normalizeObjectSchema } = await import('@modelcontextprotocol/sdk/server/zod-compat.js');
    const { toJsonSchemaCompat } = await import('@modelcontextprotocol/sdk/server/zod-json-schema-compat.js');
    const obj = normalizeObjectSchema(server._registeredTools['graph_get'].inputSchema);
    const jsonSchema = toJsonSchemaCompat(obj, { strictUnions: true, pipeStrategy: 'input' });
    assert.ok(jsonSchema.properties.endpoint, 'should have endpoint property');
    assert.ok(jsonSchema.properties.params, 'should have params property');
    assert.ok(jsonSchema.required.includes('endpoint'), 'endpoint should be required');
    assert.ok(!jsonSchema.required.includes('params'), 'params should not be required');
  });

  it('graph_post JSON Schema has method, endpoint, body, useOutlookToken', async () => {
    const { normalizeObjectSchema } = await import('@modelcontextprotocol/sdk/server/zod-compat.js');
    const { toJsonSchemaCompat } = await import('@modelcontextprotocol/sdk/server/zod-json-schema-compat.js');
    const obj = normalizeObjectSchema(server._registeredTools['graph_post'].inputSchema);
    const jsonSchema = toJsonSchemaCompat(obj, { strictUnions: true, pipeStrategy: 'input' });
    for (const prop of ['method', 'endpoint', 'body', 'useOutlookToken']) {
      assert.ok(jsonSchema.properties[prop], `should have ${prop} property`);
    }
    assert.ok(jsonSchema.required.includes('method'), 'method should be required');
    assert.ok(jsonSchema.required.includes('endpoint'), 'endpoint should be required');
    assert.ok(!jsonSchema.required.includes('body'), 'body should not be required');
  });

  it('graph_docs JSON Schema has topic enum', async () => {
    const { normalizeObjectSchema } = await import('@modelcontextprotocol/sdk/server/zod-compat.js');
    const { toJsonSchemaCompat } = await import('@modelcontextprotocol/sdk/server/zod-json-schema-compat.js');
    const obj = normalizeObjectSchema(server._registeredTools['graph_docs'].inputSchema);
    const jsonSchema = toJsonSchemaCompat(obj, { strictUnions: true, pipeStrategy: 'input' });
    assert.ok(jsonSchema.properties.topic, 'should have topic property');
    assert.deepStrictEqual(jsonSchema.properties.topic.enum, ['email', 'calendar', 'teams', 'users', 'patterns']);
  });
});
