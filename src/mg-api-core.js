'use strict';

const fs = require('fs');
const path = require('path');
const { spawnSync } = require('child_process');
const { capabilities } = require('./registry');
const { emitSchema, renderRootHelp, renderCapabilityHelp, renderVerbHelp } = require('./renderers');
const { AUTH_FILE, authenticate, authStatus, logout } = require('./graph-auth');
const { executeGraphRequest } = require('./graph-rest');

const repoRoot = path.join(__dirname, '..');

function envelope(ok, command, data, error, meta = {}) {
  return {
    ok,
    command,
    data: ok ? data : null,
    error: ok ? null : error,
    meta: { ...meta, schemaVersion: '0.1.0' },
  };
}

function writeJson(stdout, value) {
  stdout.write(JSON.stringify(value, null, 2) + '\n');
}

function parseArgs(args) {
  const positional = [];
  const flags = {};
  for (let i = 0; i < args.length; i++) {
    const arg = args[i];
    if (!arg.startsWith('--')) {
      positional.push(arg);
      continue;
    }
    const name = arg.slice(2);
    const next = args[i + 1];
    if (!next || next.startsWith('--')) {
      flags[name] = true;
      continue;
    }
    flags[name] = next;
    i++;
  }
  return { positional, flags };
}

function splitCsv(value) {
  return String(value)
    .split(',')
    .map(item => item.trim())
    .filter(item => item.length > 0);
}

function shapeValue(shape, value) {
  if (shape === 'recipient') return { emailAddress: { address: value } };
  if (shape === 'attendee') return { emailAddress: { address: value }, type: 'required' };
  return value;
}

function coerceValue(param, value) {
  if (value === undefined && Object.hasOwn(param, 'default')) return param.default;
  if (value === undefined) return undefined;
  if (param.type === 'number') {
    const numberValue = Number(value);
    if (!Number.isFinite(numberValue)) throw new Error(`--${param.name} must be a number`);
    return numberValue;
  }
  if (param.type === 'boolean') {
    if (value === true) return true;
    if (value === 'true') return true;
    if (value === 'false') return false;
    throw new Error(`--${param.name} must be true or false`);
  }
  if (param.type === 'json') {
    try {
      return JSON.parse(value);
    } catch {
      throw new Error(`--${param.name} must be valid JSON`);
    }
  }
  if (param.type === 'csv') {
    if (value === true) throw new Error(`--${param.name} requires a value`);
    const parts = splitCsv(value);
    return parts.map(item => shapeValue(param.valueShape, item));
  }
  return String(value);
}

function collectParams(spec, flags) {
  const values = {};
  for (const param of spec.params) {
    const value = coerceValue(param, flags[param.name]);
    if (value === undefined && param.required) {
      throw new Error(`Missing required option --${param.name}`);
    }
    if (value !== undefined) values[param.name] = value;
  }
  return values;
}

function formatScalar(value) {
  if (typeof value === 'boolean') return value ? 'true' : 'false';
  return String(value);
}

function replacePlaceholders(template, values) {
  return template.replace(/\{([^}]+)\}/g, (_, name) => {
    if (!Object.hasOwn(values, name)) return '';
    return formatScalar(values[name]);
  });
}

function replacePathPlaceholders(template, values) {
  return template.replace(/\{([^}]+)\}/g, (_, name) => {
    if (!Object.hasOwn(values, name)) return '';
    return encodeURIComponent(formatScalar(values[name]));
  });
}

function addQuery(endpoint, query, values) {
  if (!query) return endpoint;
  const parts = [];
  for (const [name, template] of Object.entries(query)) {
    const raw = replacePlaceholders(template, values);
    if (raw === '') continue;
    parts.push(`${encodeURIComponent(name)}=${encodeURIComponent(raw)}`);
  }
  if (!parts.length) return endpoint;
  return `${endpoint}${endpoint.includes('?') ? '&' : '?'}${parts.join('&')}`;
}

function buildBody(spec, values) {
  if (Object.hasOwn(values, 'body') && spec.params.some(p => p.name === 'body' && p.type === 'json')) {
    return JSON.stringify(values.body);
  }
  if (!spec.bodyTemplate) return '';
  function visit(value) {
    if (typeof value === 'string') {
      const match = value.match(/^\{([^}]+)\}$/);
      if (match) {
        const v = values[match[1]];
        return v;
      }
      const expanded = replacePlaceholders(value, values);
      return expanded === '' ? undefined : expanded;
    }
    if (Array.isArray(value)) {
      const mapped = value.map(visit).filter(item => item !== undefined);
      return mapped.length ? mapped : undefined;
    }
    if (value && typeof value === 'object') {
      const result = {};
      let hasAny = false;
      for (const [key, child] of Object.entries(value)) {
        const r = visit(child);
        if (r !== undefined) {
          result[key] = r;
          hasAny = true;
        }
      }
      return hasAny ? result : undefined;
    }
    return value;
  }
  const body = visit(spec.bodyTemplate);
  return body === undefined ? '' : JSON.stringify(body);
}

function buildGraphRequest(spec, values) {
  return {
    endpoint: addQuery(replacePathPlaceholders(spec.path, values), spec.query, values),
    body: buildBody(spec, values),
  };
}

async function runGraph(spec, values, deps = {}) {
  const { endpoint, body } = buildGraphRequest(spec, values);
  return executeGraphRequest(spec, endpoint, body, deps);
}

function doctor() {
  const nodeMajor = Number.parseInt(process.versions.node.split('.')[0], 10);
  const raw = [
    { name: 'node', ok: nodeMajor >= 24, detail: process.version, hint: 'Install Node.js 24 or later' },
    { name: 'auth-file', ok: fs.existsSync(AUTH_FILE), detail: AUTH_FILE, hint: 'Run "mg-api auth login" to create it' },
  ];
  const checks = raw.map(c => (c.ok ? { name: c.name, ok: true, detail: c.detail } : c));
  return { checks };
}

function gitPullMadeNoChanges(output) {
  return /already up[- ]to[- ]date\.?/i.test(output);
}

function runCommand(command, args, cwd) {
  return spawnSync(command, args, {
    cwd,
    encoding: 'utf8',
    stdio: ['ignore', 'pipe', 'pipe'],
    shell: process.platform === 'win32',
  });
}

function isGitRepo(cwd) {
  const result = runCommand('git', ['rev-parse', '--is-inside-work-tree'], cwd);
  return result.status === 0 && result.stdout.trim() === 'true';
}

function selfUpdate(deps = {}) {
  const root = deps.repoRoot || repoRoot;
  const checkGitRepo = deps.isGitRepo || isGitRepo;
  const run = deps.runCommand || runCommand;
  const steps = [];

  if (!checkGitRepo(root)) {
    return {
      ok: false,
      data: { repoRoot: root, steps },
      error: { code: 'NOT_GIT_REPO', message: 'mg-api update requires a git clone install' },
    };
  }

  const pull = run('git', ['pull', '--ff-only'], root);
  const pullOutput = `${pull.stdout || ''}${pull.stderr || ''}`.trim();
  steps.push({ name: 'git pull --ff-only', ok: pull.status === 0, output: pullOutput });
  if (pull.status !== 0) {
    return {
      ok: false,
      data: { repoRoot: root, steps },
      error: { code: 'GIT_PULL_FAILED', message: pullOutput || 'git pull --ff-only failed' },
    };
  }
  if (gitPullMadeNoChanges(pullOutput)) {
    return { ok: true, data: { repoRoot: root, updated: false, steps } };
  }

  const install = run('npm', ['install', '--no-audit', '--no-fund'], root);
  const installOutput = `${install.stdout || ''}${install.stderr || ''}`.trim();
  steps.push({ name: 'npm install --no-audit --no-fund', ok: install.status === 0, output: installOutput });
  if (install.status !== 0) {
    return {
      ok: false,
      data: { repoRoot: root, steps },
      error: { code: 'NPM_INSTALL_FAILED', message: installOutput || 'npm install failed' },
    };
  }

  const build = run('npm', ['run', 'build'], root);
  const buildOutput = `${build.stdout || ''}${build.stderr || ''}`.trim();
  steps.push({ name: 'npm run build', ok: build.status === 0, output: buildOutput });
  if (build.status !== 0) {
    return {
      ok: false,
      data: { repoRoot: root, steps },
      error: { code: 'BUILD_FAILED', message: buildOutput || 'npm run build failed' },
    };
  }

  return { ok: true, data: { repoRoot: root, updated: true, steps } };
}

async function runAuth(verbName, flags) {
  if (verbName === 'status') return { ok: true, data: authStatus() };
  if (verbName === 'login') {
    const data = await authenticate({ forceLogin: !!flags.force, verbose: !!flags.verbose });
    return {
      ok: true,
      data: {
        authenticated: true,
        hasGraphToken: !!data.GRAPH_TOKEN,
        hasOutlookToken: !!data.OUTLOOK_TOKEN,
        hasChatToken: !!data.GRAPH_CHAT_TOKEN,
        hasChannelMessageToken: !!data.GRAPH_CHAT_TOKEN && !!data.CHANNEL_MESSAGE_SCOPE_OBSERVED,
        channelMessageScopeObserved: !!data.CHANNEL_MESSAGE_SCOPE_OBSERVED,
        teamsChannelProbe: data.TEAMS_CHANNEL_PROBE,
        graphScopes: (data.GRAPH_SCOPES || []).length,
        outlookScopes: (data.OUTLOOK_SCOPES || []).length,
        chatScopes: (data.GRAPH_CHAT_SCOPES || []).length,
      },
    };
  }
  const data = logout();
  return { ok: true, data: { loggedOut: data.cleared, authFile: data.authFile, profileDir: data.profileDir } };
}

function fail(stdout, code, command, message, details) {
  writeJson(stdout, envelope(false, command, null, { code, message, details }));
}

async function main(args, io) {
  const { stdout, exit } = io;
  const parsed = parseArgs(args);
  const [capabilityName, verbName] = parsed.positional;
  if (!capabilityName || capabilityName === 'help') {
    stdout.write(renderRootHelp());
    exit(0);
    return;
  }
  if (capabilityName === 'schema') {
    const schema = emitSchema(verbName, parsed.positional[2]);
    if (!schema) {
      fail(stdout, 'UNKNOWN_SCHEMA', 'schema', `Unknown schema target: ${parsed.positional.slice(1).join(' ')}`);
      exit(2);
      return;
    }
    writeJson(stdout, envelope(true, 'schema', schema, null));
    exit(0);
    return;
  }
  if (capabilityName === 'doctor') {
    if (verbName === '--help' || parsed.flags.help) {
      stdout.write(renderVerbHelp(capabilities.doctor, 'run', capabilities.doctor.verbs.run));
      exit(0);
      return;
    }
    const data = doctor();
    const failed = data.checks.filter(check => !check.ok).map(check => check.name);
    const ok = failed.length === 0;
    const result = {
      ok,
      command: 'doctor',
      data,
      error: ok
        ? null
        : {
            code: 'DOCTOR_FAILED',
            message: `Doctor check failed: ${failed.join(', ')}. Run "mg-api auth login" if auth-file is the only failure.`,
            failed,
          },
      meta: { schemaVersion: '0.1.0' },
    };
    writeJson(stdout, result);
    exit(ok ? 0 : 1);
    return;
  }
  if (capabilityName === 'update') {
    if (verbName === '--help' || parsed.flags.help) {
      stdout.write(renderVerbHelp(capabilities.update, 'run', capabilities.update.verbs.run));
      exit(0);
      return;
    }
    const result = selfUpdate();
    writeJson(stdout, envelope(result.ok, 'update.run', result.data, result.error));
    exit(result.ok ? 0 : 1);
    return;
  }
  const capability = capabilities[capabilityName];
  if (!capability) {
    fail(stdout, 'UNKNOWN_CAPABILITY', capabilityName, `Unknown capability: ${capabilityName}`);
    exit(2);
    return;
  }
  if (!verbName) {
    stdout.write(renderCapabilityHelp(capability));
    exit(0);
    return;
  }
  const spec = capability.verbs[verbName];
  if (!spec) {
    fail(stdout, 'UNKNOWN_VERB', `${capabilityName}.${verbName}`, `Unknown verb: ${capabilityName} ${verbName}`);
    exit(2);
    return;
  }
  if (parsed.flags.help) {
    stdout.write(renderVerbHelp(capability, verbName, spec));
    exit(0);
    return;
  }
  let values;
  try {
    values = collectParams(spec, parsed.flags);
  } catch (err) {
    fail(stdout, 'VALIDATION_FAILED', spec.id, err.message);
    exit(2);
    return;
  }
  if (capabilityName === 'auth') {
    try {
      const result = await runAuth(verbName, parsed.flags);
      writeJson(stdout, envelope(result.ok, spec.id, result.data, null));
      exit(0);
    } catch (err) {
      writeJson(stdout, envelope(false, spec.id, null, { code: 'AUTH_FAILED', message: err.message }));
      exit(1);
    }
    return;
  }
  let execution;
  try {
    execution = await runGraph(spec, values);
  } catch (err) {
    fail(stdout, 'GRAPH_REQUEST_FAILED', spec.id, err.message);
    exit(1);
    return;
  }
  writeJson(stdout, envelope(true, spec.id, execution.data, null, {
    endpoint: execution.endpoint,
    method: execution.method,
    base: execution.base,
    token: execution.token,
  }));
  exit(0);
}

module.exports = {
  main,
  parseArgs,
  collectParams,
  coerceValue,
  emitSchema,
  renderRootHelp,
  renderCapabilityHelp,
  renderVerbHelp,
  envelope,
  gitPullMadeNoChanges,
  buildGraphRequest,
  buildBody,
  addQuery,
  replacePlaceholders,
  runGraph,
  selfUpdate,
  doctor,
};
