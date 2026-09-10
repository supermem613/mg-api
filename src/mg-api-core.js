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

// Recipient and attendee payload shapes are NOT portable across the two backends.
// Graph v1.0 (base: 'graph') types are camelCase; Outlook REST v2.0 (base: 'outlook')
// types are PascalCase and reject camelCase outright with
// "The property 'emailAddress' does not exist on type 'Microsoft.OutlookServices.Recipient'".
// email send/reply are the only outlook-based verbs that shape values, so this is the
// one place the distinction has to be made.
function shapeValue(shape, value, base) {
  const isOutlook = base === 'outlook';
  if (shape === 'recipient') {
    return isOutlook
      ? { EmailAddress: { Address: value } }
      : { emailAddress: { address: value } };
  }
  if (shape === 'attendee') {
    return isOutlook
      ? { EmailAddress: { Address: value }, Type: 'Required' }
      : { emailAddress: { address: value }, type: 'required' };
  }
  return value;
}

function coerceValue(param, value, base) {
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
    return parts.map(item => shapeValue(param.valueShape, item, base));
  }
  // A 'file' param carries its content, not its path. Windows caps a process
  // command line at 32,767 characters, so any body that can outgrow that must
  // reach the CLI through the filesystem instead of through argv.
  if (param.type === 'file') {
    if (value === true) throw new Error(`--${param.name} requires a file path`);
    const filePath = path.resolve(String(value));
    let content;
    try {
      content = fs.readFileSync(filePath, 'utf8');
    } catch (err) {
      const why = err.code === 'ENOENT' ? 'file not found'
        : err.code === 'EISDIR' ? 'path is a directory'
        : err.code === 'EACCES' ? 'permission denied'
        : err.message;
      throw new Error(`--${param.name} cannot read ${filePath}: ${why}`);
    }
    // A UTF-8 BOM would ship as a stray glyph at the top of an HTML mail body.
    return content.charCodeAt(0) === 0xfeff ? content.slice(1) : content;
  }
  return String(value);
}

function collectParams(spec, flags) {
  // Conflicts are settled on the raw flags, before coercion. A 'file' param reads
  // from disk during coercion, so checking afterwards would report "file not found"
  // for a user whose actual mistake was passing both --body and --body-file.
  for (const param of spec.params) {
    if (!param.fills) continue;
    if (flags[param.name] !== undefined && flags[param.fills] !== undefined) {
      throw new Error(`Pass either --${param.fills} or --${param.name}, not both`);
    }
  }
  const values = {};
  for (const param of spec.params) {
    const value = coerceValue(param, flags[param.name], spec.base);
    if (value !== undefined) values[param.name] = value;
  }
  // A param declaring `fills` hands its value to another param and disappears,
  // so --body-file and --body are one slot in the body template, never two.
  for (const param of spec.params) {
    if (!param.fills || values[param.name] === undefined) continue;
    values[param.fills] = values[param.name];
    delete values[param.name];
  }
  for (const param of spec.params) {
    if (param.required && values[param.name] === undefined) {
      throw new Error(`Missing required option --${param.name}`);
    }
  }
  for (const group of spec.requireOneOf ?? []) {
    if (group.some(name => values[name] !== undefined)) continue;
    throw new Error(`Provide one of ${group.map(name => `--${name}`).join(' or ')}`);
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

function resolveCommandInvocation(command, args, platform = process.platform) {
  if (platform === 'win32') {
    return { command: 'cmd.exe', args: ['/d', '/s', '/c', command, ...args] };
  }
  return { command, args };
}

function runCommand(command, args, cwd) {
  const invocation = resolveCommandInvocation(command, args);
  return spawnSync(invocation.command, invocation.args, {
    cwd,
    encoding: 'utf8',
    stdio: ['ignore', 'pipe', 'pipe'],
  });
}

function isGitRepo(cwd) {
  const result = runCommand('git', ['rev-parse', '--is-inside-work-tree'], cwd);
  return result.status === 0 && result.stdout.trim() === 'true';
}

function hasSodaWorkspaceMarkers(dir) {
  const workspaceDir = path.join(dir, '.sd');
  const metaPath = path.join(workspaceDir, 'meta.json');
  const repoIdPath = path.join(workspaceDir, 'repo-id');
  if (!fs.existsSync(metaPath) || !fs.existsSync(repoIdPath)) {
    return false;
  }
  try {
    return fs.readFileSync(repoIdPath, 'utf8').trim().length > 0;
  } catch {
    return false;
  }
}

function isSodaGitInterlockError(message) {
  return /sd-powered repo/i.test(message) || /raw git .* blocked/i.test(message);
}

function sodaWorktreeChanged(outcomes) {
  return outcomes.some(outcome => outcome?.worktreeUpdated === true || outcome?.worktree === true);
}

function parseJsonEnvelope(stdout) {
  try {
    return JSON.parse(stdout || '');
  } catch {
    return null;
  }
}

// soda's git-interlock hooks block raw git writes in an sd-powered repo, and
// soda tracks stream state a raw `git pull` bypasses, so a soda-managed
// checkout must self-update through `sd`. Detection uses sd status when
// available, plus local .sd workspace markers so a missing sd binary cannot be
// mistaken for a plain checkout.
function isSodaManagedRepo(run, cwd, hasSodaWorkspace) {
  try {
    const result = run('sd', ['status'], cwd);
    const envelope = parseJsonEnvelope(result.stdout);
    if (envelope?.ok === true && envelope.data?.summary?.initialized === true) {
      return true;
    }
  } catch {
    // Fall through to workspace markers. A missing sd binary is not a plain git checkout.
  }
  return hasSodaWorkspace(cwd);
}

function sodaEnvelopeError(envelope) {
  if (!envelope || envelope.error == null) return null;
  if (typeof envelope.error === 'string') return envelope.error;
  if (typeof envelope.error.message === 'string') return envelope.error.message;
  return JSON.stringify(envelope.error);
}

function sodaPullUpdate(run, root, steps) {
  const pull = run('sd', ['pull'], root);
  const pullOutput = `${pull.stdout || ''}${pull.stderr || ''}`.trim();
  steps.push({ name: 'sd pull', ok: pull.status === 0, output: pullOutput });
  const envelope = parseJsonEnvelope(pull.stdout);
  if (pull.status !== 0 || envelope?.ok !== true) {
    return {
      ok: false,
      data: { repoRoot: root, steps },
      error: { code: 'SD_PULL_FAILED', message: sodaEnvelopeError(envelope) || pullOutput || 'sd pull failed' },
    };
  }
  const outcomes = Array.isArray(envelope.data) ? envelope.data : [];
  if (!sodaWorktreeChanged(outcomes)) {
    return { ok: true, data: { repoRoot: root, updated: false, steps } };
  }
  return null;
}

function selfUpdate(deps = {}) {
  const root = deps.repoRoot || repoRoot;
  const checkGitRepo = deps.isGitRepo || isGitRepo;
  const run = deps.runCommand || runCommand;
  const hasSodaWorkspace = deps.hasSodaWorkspace || hasSodaWorkspaceMarkers;
  const steps = [];

  if (!checkGitRepo(root)) {
    return {
      ok: false,
      data: { repoRoot: root, steps },
      error: { code: 'NOT_GIT_REPO', message: 'mg-api update requires a git clone install' },
    };
  }

  if (isSodaManagedRepo(run, root, hasSodaWorkspace)) {
    const sodaResult = sodaPullUpdate(run, root, steps);
    if (sodaResult) return sodaResult;
  } else {
    const pull = run('git', ['pull', '--ff-only'], root);
    const pullOutput = `${pull.stdout || ''}${pull.stderr || ''}`.trim();
    steps.push({ name: 'git pull --ff-only', ok: pull.status === 0, output: pullOutput });
    if (pull.status !== 0) {
      if (isSodaGitInterlockError(pullOutput)) {
        const sodaResult = sodaPullUpdate(run, root, steps);
        if (sodaResult) return sodaResult;
      } else {
        return {
          ok: false,
          data: { repoRoot: root, steps },
          error: { code: 'GIT_PULL_FAILED', message: pullOutput || 'git pull --ff-only failed' },
        };
      }
    } else if (gitPullMadeNoChanges(pullOutput)) {
      return { ok: true, data: { repoRoot: root, updated: false, steps } };
    }
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
  hasSodaWorkspaceMarkers,
  isSodaGitInterlockError,
  buildGraphRequest,
  buildBody,
  addQuery,
  replacePlaceholders,
  runGraph,
  selfUpdate,
  doctor,
};
