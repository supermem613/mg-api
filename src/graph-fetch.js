'use strict';

const RETRYABLE = new Set([
  'ETIMEDOUT',
  'ECONNRESET',
  'ECONNREFUSED',
  'UND_ERR_CONNECT_TIMEOUT',
  'UND_ERR_SOCKET',
]);
const MAX_RETRIES = 2;

function extractCode(err) {
  let cur = err;
  while (cur) {
    if (cur.code) return cur.code;
    if (cur.cause) {
      cur = cur.cause;
      continue;
    }
    if (cur.errors?.length) {
      cur = cur.errors[0];
      continue;
    }
    break;
  }
  return null;
}

function formatError(err) {
  const code = extractCode(err);
  const lines = [`ERROR: fetch failed - ${err.message}`];
  if (code) lines.push(`  Code: ${code}`);

  let cause = err.cause;
  let depth = 0;
  while (cause && depth < 3) {
    lines.push(`  Cause: ${cause.message}${cause.code ? ` (${cause.code})` : ''}`);
    cause = cause.cause || cause.errors?.[0];
    depth++;
  }

  if (code === 'ETIMEDOUT' || code === 'UND_ERR_CONNECT_TIMEOUT') {
    lines.push('  Hint: Connection timed out. Check network connectivity and DNS.');
  } else if (code === 'ECONNREFUSED') {
    lines.push('  Hint: Connection refused. Verify the Graph endpoint is reachable.');
  } else if (code === 'ENOTFOUND') {
    lines.push('  Hint: DNS lookup failed. Check hostname and network connectivity.');
  }

  return lines.join('\n');
}

async function graphFetch(url, options, deps = {}) {
  const fetchImpl = deps.fetch || fetch;
  let lastErr;
  for (let attempt = 0; attempt <= MAX_RETRIES; attempt++) {
    try {
      const res = await fetchImpl(url, options);
      if (res.status === 429 && attempt < MAX_RETRIES) {
        const retryAfter = parseInt(res.headers?.get?.('Retry-After') ?? '', 10);
        const delayMs = (retryAfter > 0 ? retryAfter : (attempt + 1)) * 1000;
        await new Promise(r => setTimeout(r, delayMs));
        continue;
      }
      return res;
    } catch (err) {
      lastErr = err;
      const code = extractCode(err);
      if (!code || !RETRYABLE.has(code) || attempt === MAX_RETRIES) break;
    }
  }
  const enriched = new Error(formatError(lastErr));
  enriched.originalError = lastErr;
  throw enriched;
}

module.exports = { graphFetch, formatError, extractCode };
