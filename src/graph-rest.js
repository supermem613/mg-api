'use strict';

const { graphFetch } = require('./graph-fetch');
const { AUTH_FILE, readAuthFile } = require('./graph-auth');

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';
const OUTLOOK_BASE = 'https://outlook.office.com/api/v2.0';

function normalizeEndpoint(endpoint) {
  return endpoint.startsWith('/') ? endpoint : `/${endpoint}`;
}

function parseResponseBody(text) {
  const trimmed = text.trim();
  if (!trimmed) return null;
  try {
    return JSON.parse(trimmed);
  } catch {
    return trimmed;
  }
}

function loadAuth(authFile = AUTH_FILE) {
  const auth = readAuthFile(authFile);
  if (!auth) {
    throw new Error('No Microsoft Graph credentials are cached. Run mg-api auth login.');
  }
  if (!auth.GRAPH_TOKEN && !auth.OUTLOOK_TOKEN) {
    throw new Error('No Microsoft Graph or Outlook token is cached. Run mg-api auth login.');
  }
  return auth;
}

function resolveBase(spec) {
  return (spec.base || 'graph') === 'outlook' ? OUTLOOK_BASE : GRAPH_BASE;
}

function resolveToken(spec, auth) {
  const wanted = spec.token || 'graph';
  if (wanted === 'outlook') {
    if (!auth.OUTLOOK_TOKEN) {
      throw new Error(`Verb ${spec.id} requires an Outlook token but none is cached. Re-run mg-api auth login.`);
    }
    return auth.OUTLOOK_TOKEN;
  }
  if (wanted === 'chat') {
    const token = auth.GRAPH_CHAT_TOKEN || auth.GRAPH_TOKEN;
    if (!token) {
      throw new Error(`Verb ${spec.id} requires a Graph chat token but none is cached. Re-run mg-api auth login.`);
    }
    return token;
  }
  if (!auth.GRAPH_TOKEN) {
    throw new Error(`Verb ${spec.id} requires a Graph token but none is cached. Re-run mg-api auth login.`);
  }
  return auth.GRAPH_TOKEN;
}

async function executeGraphRequest(spec, endpoint, body, deps = {}) {
  const auth = deps.auth || loadAuth(deps.authFile);
  const base = resolveBase(spec);
  const token = resolveToken(spec, auth);
  const url = `${base}${normalizeEndpoint(endpoint)}`;
  const method = spec.method || 'GET';
  const headers = {
    Authorization: `Bearer ${token}`,
    Accept: spec.accept || 'application/json',
  };
  const fetchOptions = { method, headers };
  if (method !== 'GET' && method !== 'DELETE' && body) {
    fetchOptions.headers['Content-Type'] = spec.contentType || 'application/json';
    fetchOptions.body = body;
  }
  const res = await graphFetch(url, fetchOptions, deps);
  const text = await res.text();
  if (!res.ok) {
    throw new Error(`HTTP ${res.status} on ${method} ${url}: ${text}`);
  }
  return {
    data: parseResponseBody(text),
    endpoint: normalizeEndpoint(endpoint),
    method,
    base: spec.base || 'graph',
    token: spec.token || 'graph',
  };
}

module.exports = {
  GRAPH_BASE,
  OUTLOOK_BASE,
  executeGraphRequest,
  loadAuth,
  normalizeEndpoint,
  parseResponseBody,
  resolveBase,
  resolveToken,
};
