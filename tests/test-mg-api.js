#!/usr/bin/env node
'use strict';

const { describe, it } = require('node:test');
const assert = require('node:assert');
const { spawnSync } = require('node:child_process');
const { existsSync, readFileSync } = require('node:fs');
const { join } = require('node:path');
const { capabilities } = require('../src/registry');
const {
  emitSchema,
  renderRootHelp,
  renderCapabilityHelp,
  renderVerbHelp,
  renderSkillRouter,
} = require('../src/renderers');
const {
  buildGraphRequest,
  buildBody,
  gitPullMadeNoChanges,
  selfUpdate,
  coerceValue,
} = require('../src/mg-api-core');

const repoRoot = join(__dirname, '..');
const cliPath = join(repoRoot, 'bin', 'mg-api.js');
const skillPath = join(repoRoot, '.claude', 'skills', 'mg-api', 'SKILL.md');

function runCli(args) {
  return spawnSync(process.execPath, [cliPath, ...args], {
    cwd: repoRoot,
    encoding: 'utf8',
    stdio: ['ignore', 'pipe', 'pipe'],
  });
}

function parseJson(stdout) {
  return JSON.parse(stdout);
}

describe('mg-api package wiring', () => {
  it('exposes a bin script with a shebang', () => {
    const pkg = JSON.parse(readFileSync(join(repoRoot, 'package.json'), 'utf8'));
    assert.strictEqual(pkg.name, 'mg-api');
    assert.match(pkg.version, /^\d+\.\d+\.\d+(?:-[\w.]+)?$/);
    assert.strictEqual(pkg.bin['mg-api'], 'bin/mg-api.js');
    assert.strictEqual(pkg.scripts.build, 'node scripts/build.js');
    assert.strictEqual(pkg.scripts.prepare, 'npm run build');
    assert.strictEqual(pkg.scripts['link:local'], 'npm run build && npm link');
    assert.strictEqual(pkg.scripts.auth, 'mg-api auth status');
    assert.strictEqual(pkg.scripts.login, 'mg-api auth login');
    assert.strictEqual(pkg.scripts.logout, 'mg-api auth logout');
    assert.match(readFileSync(cliPath, 'utf8').split(/\r?\n/)[0], /node/);
  });
});

describe('registry invariants', () => {
  it('has unique capability and verb ids with examples and output contracts', () => {
    const ids = new Set();
    for (const [capabilityName, capability] of Object.entries(capabilities)) {
      assert.strictEqual(capabilityName, capability.id);
      assert.ok(capability.summary);
      for (const [verbName, verb] of Object.entries(capability.verbs)) {
        assert.strictEqual(verb.id, `${capabilityName}.${verbName}`);
        assert.ok(!ids.has(verb.id), `Duplicate id ${verb.id}`);
        ids.add(verb.id);
        assert.ok(verb.summary, `${verb.id} missing summary`);
        assert.ok(verb.output?.envelope, `${verb.id} missing envelope`);
        assert.ok(verb.examples?.length, `${verb.id} missing examples`);
        const params = new Set();
        for (const param of verb.params) {
          assert.match(param.name, /^[a-z][a-z0-9-]*$/);
          assert.ok(!params.has(param.name), `${verb.id} duplicate param ${param.name}`);
          params.add(param.name);
          assert.ok(param.doc, `${verb.id}.${param.name} missing docs`);
        }
      }
    }
  });

  it('declares token and base for every graph-rest verb', () => {
    for (const capability of Object.values(capabilities)) {
      for (const [verbName, verb] of Object.entries(capability.verbs)) {
        if (verb.handler !== 'graph-rest') continue;
        assert.ok(['graph', 'outlook', 'chat'].includes(verb.token), `${verb.id} missing token`);
        assert.ok(['graph', 'outlook'].includes(verb.base), `${verb.id} missing base`);
        assert.ok(verb.path, `${verb.id} missing path`);
        assert.ok(['GET', 'POST', 'PATCH', 'DELETE'].includes(verb.method), `${verb.id} invalid method`);
        assert.ok(verb.id.endsWith(`.${verbName}`), `${verb.id} id mismatch`);
      }
    }
  });

  it('routes email and chat verbs to the expected token and base', () => {
    assert.strictEqual(capabilities.email.verbs.list.token, 'outlook');
    assert.strictEqual(capabilities.email.verbs.list.base, 'outlook');
    assert.strictEqual(capabilities.email.verbs.send.token, 'outlook');
    assert.strictEqual(capabilities.email.verbs.send.base, 'outlook');
    assert.strictEqual(capabilities.email.verbs.reply.token, 'outlook');
    assert.strictEqual(capabilities.teams.verbs['send-channel-message'].token, 'chat');
    assert.strictEqual(capabilities.teams.verbs['list-channels'].token, 'chat');
    assert.strictEqual(capabilities.chats.verbs.list.token, 'outlook');
    assert.strictEqual(capabilities.chats.verbs.list.base, 'outlook');
    assert.strictEqual(capabilities.chats.verbs.messages.token, 'outlook');
    assert.strictEqual(capabilities.chats.verbs.send.token, 'chat');
    assert.strictEqual(capabilities.chats.verbs.send.base, 'graph');
  });

  it('does not expose raw HTTP verbs as top-level capabilities', () => {
    assert.ok(!capabilities.get);
    assert.ok(!capabilities.post);
    assert.ok(!capabilities.request);
  });

  it('exposes all expected Microsoft Graph capability groups', () => {
    assert.ok(capabilities.email.verbs.list);
    assert.ok(capabilities.email.verbs.send);
    assert.ok(capabilities.email.verbs.reply);
    assert.ok(capabilities.email.verbs.attachments);
    assert.ok(capabilities.calendar.verbs.view);
    assert.ok(capabilities.calendar.verbs['find-times']);
    assert.ok(capabilities.calendar.verbs.accept);
    assert.ok(capabilities.teams.verbs['list-joined']);
    assert.ok(capabilities.teams.verbs['send-channel-message']);
    assert.ok(capabilities.chats.verbs.messages);
    assert.ok(capabilities.chats.verbs.send);
    assert.ok(capabilities.users.verbs.me);
    assert.ok(capabilities.users.verbs.search);
  });

  it('keeps shipped verbs out of the planned capability lists', () => {
    const schema = emitSchema();
    assert.ok(schema.plannedCapabilities.email.includes('folders'));
    assert.ok(schema.plannedCapabilities.email.includes('forward'));
    assert.ok(schema.plannedCapabilities.calendar.includes('instances'));
    assert.ok(schema.plannedCapabilities.teams.includes('members'));
    assert.ok(schema.plannedCapabilities.users.includes('list'));
    assert.ok(schema.plannedCapabilities.files.includes('upload'));
    for (const verb of schema.plannedCapabilities.email) {
      assert.ok(!capabilities.email.verbs[verb], `${verb} should not be implemented yet`);
    }
  });

  it('keeps Playwright isolated to the auth module', () => {
    const nonAuthFiles = [
      join(repoRoot, 'bin', 'mg-api.js'),
      join(repoRoot, 'src', 'registry.js'),
      join(repoRoot, 'src', 'renderers.js'),
      join(repoRoot, 'src', 'mg-api-core.js'),
      join(repoRoot, 'src', 'graph-rest.js'),
      join(repoRoot, 'src', 'graph-fetch.js'),
    ];
    for (const file of nonAuthFiles) {
      assert.doesNotMatch(readFileSync(file, 'utf8'), /require\(['"]playwright['"]\)/, file);
    }
    assert.match(readFileSync(join(repoRoot, 'src', 'graph-auth.js'), 'utf8'), /require\(['"]playwright['"]\)/);
  });

  it('does not read environment variables anywhere in src or bin', () => {
    const files = [
      join(repoRoot, 'bin', 'mg-api.js'),
      join(repoRoot, 'src', 'registry.js'),
      join(repoRoot, 'src', 'renderers.js'),
      join(repoRoot, 'src', 'mg-api-core.js'),
      join(repoRoot, 'src', 'graph-auth.js'),
      join(repoRoot, 'src', 'graph-rest.js'),
      join(repoRoot, 'src', 'graph-fetch.js'),
    ];
    for (const file of files) {
      const text = readFileSync(file, 'utf8');
      assert.doesNotMatch(text, /process\.env\b/, file);
    }
  });
});

describe('schema output', () => {
  it('emits a full schema envelope', () => {
    const r = runCli(['schema']);
    assert.strictEqual(r.status, 0);
    const json = parseJson(r.stdout);
    assert.strictEqual(json.ok, true);
    assert.strictEqual(json.command, 'schema');
    assert.deepStrictEqual(json.data.capabilities.email.verbs.list, capabilities.email.verbs.list);
  });

  it('emits a focused capability and verb schema', () => {
    const cap = parseJson(runCli(['schema', 'email']).stdout);
    assert.strictEqual(cap.data.id, 'email');
    const verb = parseJson(runCli(['schema', 'email', 'send']).stdout);
    assert.strictEqual(verb.data.id, 'email.send');
    assert.strictEqual(verb.data.method, 'POST');
    assert.strictEqual(verb.data.token, 'outlook');
    assert.strictEqual(verb.data.base, 'outlook');
  });

  it('emits focused schemas for new capability groups', () => {
    const view = parseJson(runCli(['schema', 'calendar', 'view']).stdout);
    assert.strictEqual(view.data.id, 'calendar.view');
    const findTimes = parseJson(runCli(['schema', 'calendar', 'find-times']).stdout);
    assert.strictEqual(findTimes.data.method, 'POST');
    const sendChannel = parseJson(runCli(['schema', 'teams', 'send-channel-message']).stdout);
    assert.strictEqual(sendChannel.data.token, 'chat');
  });

  it('fails unknown schema targets with a JSON envelope', () => {
    const r = runCli(['schema', 'missing']);
    assert.strictEqual(r.status, 2);
    const json = parseJson(r.stdout);
    assert.strictEqual(json.ok, false);
    assert.strictEqual(json.error.code, 'UNKNOWN_SCHEMA');
  });
});

describe('generated help', () => {
  it('generates root help from the registry', () => {
    const r = runCli(['--help']);
    assert.strictEqual(r.status, 0);
    assert.strictEqual(r.stdout, renderRootHelp());
    assert.match(r.stdout, /email\s+Read and send Outlook mail/);
  });

  it('generates capability help from the registry', () => {
    const r = runCli(['email', '--help']);
    assert.strictEqual(r.status, 0);
    assert.strictEqual(r.stdout, renderCapabilityHelp(capabilities.email));
    assert.match(r.stdout, /attachments/);
  });

  it('generates help for new capability groups from the registry', () => {
    const calendar = runCli(['calendar', '--help']);
    assert.strictEqual(calendar.status, 0);
    assert.strictEqual(calendar.stdout, renderCapabilityHelp(capabilities.calendar));
    assert.match(calendar.stdout, /find-times/);

    const send = runCli(['email', 'send', '--help']);
    assert.strictEqual(send.status, 0);
    assert.strictEqual(send.stdout, renderVerbHelp(capabilities.email, 'send', capabilities.email.verbs.send));
    assert.match(send.stdout, /--to/);
    assert.match(send.stdout, /Token: outlook/);
    assert.match(send.stdout, /Base: outlook\.office\.com/);
  });

  it('generates verb help from the registry', () => {
    const r = runCli(['email', 'list', '--help']);
    assert.strictEqual(r.status, 0);
    assert.strictEqual(r.stdout, renderVerbHelp(capabilities.email, 'list', capabilities.email.verbs.list));
    assert.match(r.stdout, /--top/);
    assert.match(r.stdout, /Endpoint:/);
  });
});

describe('JSON envelope behavior', () => {
  it('returns validation failures as JSON on stdout', () => {
    const r = runCli(['email', 'send']);
    assert.strictEqual(r.status, 2);
    assert.strictEqual(r.stderr, '');
    const json = parseJson(r.stdout);
    assert.strictEqual(json.ok, false);
    assert.strictEqual(json.command, 'email.send');
    assert.strictEqual(json.error.code, 'VALIDATION_FAILED');
    assert.match(json.error.message, /--to/);
  });

  it('rejects unknown capabilities and verbs without raw fallback', () => {
    const cap = parseJson(runCli(['request', '--method', 'GET']).stdout);
    assert.strictEqual(cap.ok, false);
    assert.strictEqual(cap.error.code, 'UNKNOWN_CAPABILITY');
    const verb = parseJson(runCli(['email', 'request']).stdout);
    assert.strictEqual(verb.ok, false);
    assert.strictEqual(verb.error.code, 'UNKNOWN_VERB');
  });

  it('reports auth status without loading Playwright', () => {
    const r = runCli(['auth', 'status']);
    assert.strictEqual(r.status, 0);
    const json = parseJson(r.stdout);
    assert.strictEqual(json.ok, true);
    assert.ok(Object.hasOwn(json.data, 'exists'));
    assert.ok(json.data.authFile.endsWith(join('.mg-api', 'auth.json')));
  });

  it('keeps the doctor checks payload on both success and failure', () => {
    const r = runCli(['doctor']);
    const json = parseJson(r.stdout);
    assert.strictEqual(json.command, 'doctor');
    assert.ok(Array.isArray(json.data.checks), 'doctor.data.checks must always be present');
    assert.ok(json.data.checks.length >= 2, 'doctor should report at least node + auth-file checks');
    const knownNames = new Set(['node', 'auth-file']);
    for (const check of json.data.checks) {
      assert.ok(knownNames.has(check.name), `unexpected check name: ${check.name}`);
      assert.strictEqual(typeof check.ok, 'boolean');
      if (check.ok) {
        assert.strictEqual(check.hint, undefined, `passing checks must not carry a hint (${check.name})`);
      } else {
        assert.strictEqual(typeof check.hint, 'string');
      }
    }
    if (json.ok) {
      assert.strictEqual(r.status, 0);
      assert.strictEqual(json.error, null);
    } else {
      assert.strictEqual(r.status, 1);
      assert.strictEqual(json.error.code, 'DOCTOR_FAILED');
      assert.ok(Array.isArray(json.error.failed), 'error.failed must list the failed check names');
      assert.ok(json.error.failed.length >= 1);
      for (const name of json.error.failed) {
        const check = json.data.checks.find(c => c.name === name);
        assert.ok(check, `failed check ${name} must appear in data.checks`);
        assert.strictEqual(check.ok, false);
      }
      assert.match(json.error.message, /Doctor check failed/);
    }
  });
});

describe('Graph request construction', () => {
  it('builds encoded search query URLs without exposing raw HTTP args', () => {
    const request = buildGraphRequest(capabilities.email.verbs.search, {
      query: 'quarterly review',
      top: 10,
    });
    assert.strictEqual(request.endpoint, '/me/messages?%24search=%22quarterly%20review%22&%24top=10');
    assert.strictEqual(request.body, '');
  });

  it('routes mailbox message verbs through the Outlook REST token', () => {
    for (const verbName of ['list', 'get', 'search', 'move', 'delete', 'attachments']) {
      assert.strictEqual(capabilities.email.verbs[verbName].token, 'outlook', `email ${verbName} token`);
      assert.strictEqual(capabilities.email.verbs[verbName].base, 'outlook', `email ${verbName} base`);
    }
  });

  it('builds email send bodies with PascalCase recipients and content type', () => {
    const request = buildGraphRequest(capabilities.email.verbs.send, {
      to: [{ emailAddress: { address: 'alice@example.com' } }],
      cc: [{ emailAddress: { address: 'bob@example.com' } }],
      subject: 'Hello',
      body: 'Hi there',
      'body-type': 'Text',
      'save-to-sent': true,
    });
    assert.strictEqual(request.endpoint, '/me/sendmail');
    const parsed = JSON.parse(request.body);
    assert.deepStrictEqual(parsed.Message.ToRecipients, [{ emailAddress: { address: 'alice@example.com' } }]);
    assert.deepStrictEqual(parsed.Message.CcRecipients, [{ emailAddress: { address: 'bob@example.com' } }]);
    assert.strictEqual(parsed.Message.Subject, 'Hello');
    assert.strictEqual(parsed.Message.Body.ContentType, 'Text');
    assert.strictEqual(parsed.Message.Body.Content, 'Hi there');
    assert.strictEqual(parsed.SaveToSentItems, true);
    assert.ok(!Object.hasOwn(parsed.Message, 'BccRecipients'), 'omitted optional Bcc');
  });

  it('builds calendar create bodies with attendees and start/end blocks', () => {
    const request = buildGraphRequest(capabilities.calendar.verbs.create, {
      subject: 'Standup',
      body: 'Daily sync',
      'body-type': 'Text',
      start: '2026-01-15T09:00:00',
      end: '2026-01-15T09:30:00',
      'time-zone': 'America/New_York',
      attendees: [
        { emailAddress: { address: 'alice@example.com' }, type: 'required' },
        { emailAddress: { address: 'bob@example.com' }, type: 'required' },
      ],
    });
    const parsed = JSON.parse(request.body);
    assert.strictEqual(parsed.subject, 'Standup');
    assert.deepStrictEqual(parsed.start, { dateTime: '2026-01-15T09:00:00', timeZone: 'America/New_York' });
    assert.deepStrictEqual(parsed.end, { dateTime: '2026-01-15T09:30:00', timeZone: 'America/New_York' });
    assert.strictEqual(parsed.attendees.length, 2);
    assert.ok(!Object.hasOwn(parsed, 'location'), 'omitted optional location');
    assert.ok(!Object.hasOwn(parsed, 'isOnlineMeeting'), 'omitted optional online flag');
  });

  it('builds calendar view, accept, and chat send endpoints/bodies', () => {
    const view = buildGraphRequest(capabilities.calendar.verbs.view, {
      start: '2026-01-15T00:00:00Z',
      end: '2026-01-16T00:00:00Z',
      top: 50,
    });
    assert.ok(view.endpoint.startsWith('/me/calendarView?'));
    assert.ok(view.endpoint.includes('startDateTime=2026-01-15T00%3A00%3A00Z'));
    assert.ok(view.endpoint.includes('endDateTime=2026-01-16T00%3A00%3A00Z'));
    assert.ok(view.endpoint.includes('%24top=50'));

    const accept = buildGraphRequest(capabilities.calendar.verbs.accept, {
      'event-id': 'AAMkAGI=',
      comment: 'See you there',
      'send-response': true,
    });
    assert.strictEqual(accept.endpoint, '/me/events/AAMkAGI%3D/accept');
    assert.deepStrictEqual(JSON.parse(accept.body), { comment: 'See you there', sendResponse: true });

    const send = buildGraphRequest(capabilities.chats.verbs.send, {
      'chat-id': '19:abc',
      content: 'Hello',
      'content-type': 'text',
    });
    assert.strictEqual(send.endpoint, '/chats/19%3Aabc/messages');
    assert.deepStrictEqual(JSON.parse(send.body), { body: { contentType: 'text', content: 'Hello' } });
  });

  it('passes explicit JSON bodies through for update verbs', () => {
    const update = buildGraphRequest(capabilities.calendar.verbs.update, {
      'event-id': 'AAMkAGI=',
      body: { subject: 'Renamed' },
    });
    assert.strictEqual(update.endpoint, '/me/events/AAMkAGI%3D');
    assert.deepStrictEqual(JSON.parse(update.body), { subject: 'Renamed' });
  });

  it('coerces csv attendee/recipient flags into Graph-shaped objects', () => {
    const param = capabilities.email.verbs.send.params.find(p => p.name === 'to');
    const value = coerceValue(param, 'alice@example.com,bob@example.com');
    assert.deepStrictEqual(value, [
      { emailAddress: { address: 'alice@example.com' } },
      { emailAddress: { address: 'bob@example.com' } },
    ]);
    const attendeeParam = capabilities.calendar.verbs.create.params.find(p => p.name === 'attendees');
    const attendees = coerceValue(attendeeParam, 'alice@example.com');
    assert.deepStrictEqual(attendees, [{ emailAddress: { address: 'alice@example.com' }, type: 'required' }]);
  });

  it('drops empty nested template branches from the body', () => {
    const body = buildBody(capabilities.email.verbs.send, {
      to: [{ emailAddress: { address: 'alice@example.com' } }],
      subject: 'Hi',
      body: 'Body',
    });
    const parsed = JSON.parse(body);
    assert.ok(!Object.hasOwn(parsed.Message, 'CcRecipients'));
    assert.ok(!Object.hasOwn(parsed.Message, 'BccRecipients'));
  });
});

describe('SKILL.md router generation', () => {
  it('matches the registry-rendered router', () => {
    assert.strictEqual(readFileSync(skillPath, 'utf8').replace(/\r\n/g, '\n'), renderSkillRouter());
  });

  it('routes every implemented capability', () => {
    const content = readFileSync(skillPath, 'utf8');
    for (const capability of Object.keys(capabilities)) {
      assert.match(content, new RegExp(`mg-api ${capability}`));
    }
  });
});

describe('update command', () => {
  it('appears in root help and schema', () => {
    const help = runCli(['--help']).stdout;
    assert.match(help, /update\s+Self-update this mg-api checkout/);
    const schema = parseJson(runCli(['schema', 'update', 'run']).stdout);
    assert.strictEqual(schema.data.id, 'update.run');
  });

  it('generates update help from the registry', () => {
    const r = runCli(['update', '--help']);
    assert.strictEqual(r.status, 0);
    assert.strictEqual(r.stdout, renderVerbHelp(capabilities.update, 'run', capabilities.update.verbs.run));
  });

  it('recognizes current and older no-change git pull output', () => {
    assert.strictEqual(gitPullMadeNoChanges('Already up to date.'), true);
    assert.strictEqual(gitPullMadeNoChanges('Already up-to-date.'), true);
    assert.strictEqual(gitPullMadeNoChanges('Fast-forward\n package.json | 2 +-'), false);
  });

  it('skips install and build when git pull made no changes', () => {
    const commands = [];
    const result = selfUpdate({
      repoRoot,
      isGitRepo: () => true,
      runCommand: (command, args) => {
        commands.push(`${command} ${args.join(' ')}`);
        return { status: 0, stdout: 'Already up to date.\n', stderr: '' };
      },
    });
    assert.strictEqual(result.ok, true);
    assert.strictEqual(result.data.updated, false);
    assert.deepStrictEqual(commands, ['git pull --ff-only']);
  });

  it('runs install and build when git pull returns changes', () => {
    const commands = [];
    const result = selfUpdate({
      repoRoot,
      isGitRepo: () => true,
      runCommand: (command, args) => {
        commands.push(`${command} ${args.join(' ')}`);
        return { status: 0, stdout: command === 'git' ? 'Fast-forward\n' : 'ok\n', stderr: '' };
      },
    });
    assert.strictEqual(result.ok, true);
    assert.strictEqual(result.data.updated, true);
    assert.deepStrictEqual(commands, [
      'git pull --ff-only',
      'npm install --no-audit --no-fund',
      'npm run build',
    ]);
  });
});

describe('agentic contract documentation', () => {
  it('documents schema-generated help and no raw passthrough', () => {
    const doc = readFileSync(join(repoRoot, 'docs', 'AGENTIC_CONTRACT.md'), 'utf8');
    assert.match(doc, /generated from that registry/);
    assert.match(doc, /There is no raw HTTP passthrough/);
    assert.match(doc, /token routing/i);
  });

  it('keeps required repo files present', () => {
    assert.ok(existsSync(join(repoRoot, 'src', 'registry.js')));
    assert.ok(existsSync(join(repoRoot, 'src', 'renderers.js')));
    assert.ok(existsSync(join(repoRoot, 'src', 'mg-api-core.js')));
    assert.ok(existsSync(join(repoRoot, 'src', 'graph-rest.js')));
    assert.ok(existsSync(join(repoRoot, 'src', 'graph-auth.js')));
    assert.ok(existsSync(join(repoRoot, 'src', 'graph-fetch.js')));
  });
});
