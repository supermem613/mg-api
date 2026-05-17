'use strict';

const { capabilities, plannedCapabilities } = require('./registry');

function cloneJson(value) {
  return JSON.parse(JSON.stringify(value));
}

function emitSchema(targetCapability, targetVerb) {
  if (!targetCapability) {
    return {
      version: '0.1.0',
      capabilities: cloneJson(capabilities),
      plannedCapabilities: cloneJson(plannedCapabilities),
    };
  }
  const capability = capabilities[targetCapability];
  if (!capability) return null;
  if (!targetVerb) return cloneJson(capability);
  const verb = capability.verbs[targetVerb];
  if (!verb) return null;
  return cloneJson(verb);
}

function renderRootHelp() {
  const lines = [
    'mg-api - agentic Microsoft Graph capability CLI',
    '',
    'Usage:',
    '  mg-api <capability> <verb> [options]',
    '  mg-api schema [capability] [verb]',
    '  mg-api doctor',
    '  mg-api update',
    '',
    'Capabilities:',
  ];
  for (const capability of Object.values(capabilities)) {
    lines.push(`  ${capability.id.padEnd(12)} ${capability.summary}`);
  }
  lines.push('');
  lines.push('Planned capability groups:');
  for (const [group, verbs] of Object.entries(plannedCapabilities)) {
    lines.push(`  ${group.padEnd(12)} ${verbs.join(', ')}`);
  }
  lines.push('');
  lines.push('Run "mg-api <capability> --help" or "mg-api schema <capability>" for generated details.');
  return lines.join('\n') + '\n';
}

function renderCapabilityHelp(capability) {
  const lines = [
    `mg-api ${capability.id} - ${capability.summary}`,
    '',
    capability.description || capability.summary,
    '',
    'Usage:',
    `  mg-api ${capability.id} <verb> [options]`,
    '',
    'Verbs:',
  ];
  for (const [verbName, verb] of Object.entries(capability.verbs)) {
    lines.push(`  ${verbName.padEnd(20)} ${verb.summary}`);
  }
  lines.push('');
  lines.push(`Run "mg-api ${capability.id} <verb> --help" for generated option details.`);
  return lines.join('\n') + '\n';
}

function renderVerbHelp(capability, verbName, verb) {
  const topLevelRun = verbName === 'run' && ['doctor', 'update'].includes(capability.id);
  const command = topLevelRun ? `mg-api ${capability.id}` : `mg-api ${capability.id} ${verbName}`;
  const lines = [
    `${command} - ${verb.summary}`,
    '',
    'Usage:',
    `  ${command}${verb.params.length ? ' [options]' : ''}`,
    '',
    `Auth: ${verb.auth}`,
    `Method: ${verb.method}`,
  ];
  if (verb.path) lines.push(`Endpoint: ${verb.path}`);
  if (verb.base) lines.push(`Base: ${verb.base === 'outlook' ? 'outlook.office.com/api/v2.0' : 'graph.microsoft.com/v1.0'}`);
  if (verb.token) lines.push(`Token: ${verb.token}`);
  if (verb.params.length) {
    lines.push('');
    lines.push('Options:');
    for (const param of verb.params) {
      const required = param.required ? 'required' : `optional${Object.hasOwn(param, 'default') ? `, default ${param.default}` : ''}`;
      lines.push(`  --${param.name.padEnd(22)} ${param.type.padEnd(8)} ${required}. ${param.doc}`);
    }
  }
  if (verb.examples && verb.examples.length) {
    lines.push('');
    lines.push('Examples:');
    for (const example of verb.examples) {
      lines.push(`  ${example}`);
    }
  }
  return lines.join('\n') + '\n';
}

function renderSkillRouter() {
  const commandRows = [];
  for (const capability of Object.values(capabilities)) {
    commandRows.push(`| \`${capability.id}\` | ${capability.summary} | \`mg-api ${capability.id} --help\` |`);
  }
  return `---
name: mg-api
description: "Use when you need to interact with Microsoft Graph through the agentic mg-api CLI for mail, calendar, Teams chats and channels, users, auth, schema inspection, and other Graph capabilities."
metadata:
  author: "Marcus Markiewicz"
  version: "1.0"
  license: "MIT"
  repo: "https://github.com/supermem613/mg-api"
---

# mg-api

This bundled skill is a thin router for the \`mg-api\` CLI. Use the CLI for Microsoft Graph work. The CLI is agentic-only: stdout is JSON, progress and remediation go to stderr, and help/schema are generated from the same capability registry.

## Execution sequence

1. Run \`mg-api doctor\` if setup or auth is uncertain.
2. Use \`mg-api schema\` to inspect the full machine-readable contract, or \`mg-api schema <capability> <verb>\` for one command.
3. Run semantic commands such as \`mg-api email list --top 10\`, \`mg-api calendar create --subject Standup --start 2026-01-15T09:00:00 --end 2026-01-15T09:30:00\`, or \`mg-api teams send-channel-message --team-id ... --channel-id ... --content "Build passed"\`.
4. If a capability is not listed in \`schema\`, do not fall back to raw HTTP. Report the missing capability so a verb can be added.

## Capabilities

| Capability | Purpose | Details |
|------------|---------|---------|
${commandRows.join('\n')}

## References

Load these only when you need deeper Microsoft Graph REST details behind a capability:

| File | Covers |
|------|--------|
| [\`references/email/\`](references/email/README.md) | Mailbox messages, search, send, reply, move, attachments |
| [\`references/calendar/\`](references/calendar/README.md) | Events, calendarView, RSVP, find meeting times |
| [\`references/teams/\`](references/teams/README.md) | Teams, channels, channel messages, plus 1:1 and group chats |
| [\`references/users/\`](references/users/README.md) | Me, people search, user profiles |
| [\`references/api-patterns/\`](references/api-patterns/README.md) | Graph vs Outlook REST, OData, token routing, pagination, errors |
`;
}

module.exports = {
  emitSchema,
  renderRootHelp,
  renderCapabilityHelp,
  renderVerbHelp,
  renderSkillRouter,
};
