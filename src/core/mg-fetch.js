#!/usr/bin/env node
'use strict';

const RETRYABLE = new Set([
  'ETIMEDOUT', 'ECONNRESET', 'ECONNREFUSED',
  'UND_ERR_CONNECT_TIMEOUT', 'UND_ERR_SOCKET',
]);
const MAX_RETRIES = 2;

/** Walk cause/errors chain to find the root network error code. */
function extractCode(err) {
  let cur = err;
  while (cur) {
    if (cur.code) return cur.code;
    if (cur.cause) { cur = cur.cause; continue; }
    if (cur.errors && cur.errors.length) { cur = cur.errors[0]; continue; }
    break;
  }
  return null;
}

/** Build a multi-line diagnostic from a fetch or Graph API error. */
function formatError(err) {
  const code = extractCode(err);
  const lines = [`ERROR: fetch failed — ${err.message}`];
  if (code) lines.push(`  Code: ${code}`);

  let cause = err.cause;
  let depth = 0;
  while (cause && depth < 3) {
    lines.push(`  Cause: ${cause.message}${cause.code ? ` (${cause.code})` : ''}`);
    cause = cause.cause || (cause.errors && cause.errors[0]);
    depth++;
  }

  if (code === 'ETIMEDOUT' || code === 'UND_ERR_CONNECT_TIMEOUT') {
    lines.push('  Hint: Connection timed out. Check network connectivity and DNS.');
  } else if (code === 'ECONNREFUSED') {
    lines.push('  Hint: Connection refused. Verify the Graph API endpoint is reachable.');
  } else if (code === 'ENOTFOUND') {
    lines.push('  Hint: DNS lookup failed. Check hostname and network connectivity.');
  }

  return lines.join('\n');
}

/** Parse Graph API error body: {"error":{"code":"...", "message":"..."}} */
async function parseGraphError(res) {
  try {
    const body = await res.json();
    if (body && body.error) return body.error;
  } catch { /* non-JSON body */ }
  return null;
}

/** Build diagnostic lines for an HTTP-level Graph API error. */
function formatHttpError(res, graphError) {
  const lines = [`ERROR: Graph API ${res.status} ${res.statusText}`];
  lines.push(`  URL: ${res.url}`);
  if (graphError) {
    if (graphError.code) lines.push(`  Code: ${graphError.code}`);
    if (graphError.message) lines.push(`  Message: ${graphError.message}`);
  }

  if (res.status === 401) {
    lines.push('  Hint: Token expired or invalid. Re-run auth to get a fresh token.');
  } else if (res.status === 403) {
    lines.push('  Hint: Insufficient permissions. Check required Graph API scopes.');
  } else if (res.status === 404) {
    lines.push('  Hint: Resource not found. Verify the endpoint path and resource ID.');
  } else if (res.status === 429) {
    const retryAfter = res.headers.get('Retry-After');
    lines.push(`  Hint: Throttled by Graph API.${retryAfter ? ` Retry after ${retryAfter}s.` : ''}`);
  }

  return lines.join('\n');
}

/** fetch() with retry on transient/throttle errors and enriched diagnostics. */
async function graphFetch(url, options) {
  let lastErr;
  for (let attempt = 0; attempt <= MAX_RETRIES; attempt++) {
    try {
      const res = await fetch(url, options);

      if (res.status === 429 && attempt < MAX_RETRIES) {
        const retryAfter = parseInt(res.headers.get('Retry-After'), 10);
        const delay = (retryAfter > 0 ? retryAfter : (attempt + 1)) * 1000;
        await new Promise(r => setTimeout(r, delay));
        continue;
      }

      return res;
    } catch (err) {
      lastErr = err;
      const code = extractCode(err);
      if (!code || !RETRYABLE.has(code) || attempt === MAX_RETRIES) break;
      await new Promise(r => setTimeout(r, 1000 * (attempt + 1)));
    }
  }
  const enriched = new Error(formatError(lastErr));
  enriched.originalError = lastErr;
  throw enriched;
}

module.exports = { graphFetch, formatError, formatHttpError, parseGraphError, extractCode };
