import crypto from 'crypto';
import { createToolError } from '../../utils/mcpErrorResponse.js';

// Short-lived, in-process dedup for non-idempotent sends. Guards against the
// same message being delivered twice when a send is re-issued after an
// ambiguous failure (e.g. MCP transport timeout then a model retry). Keyed by
// a client-supplied idempotency key when given, else a hash of the content.
// In-memory by design: the duplicate re-issue happens within seconds in the
// same process; NOT persisted across restarts.

const DEFAULT_TTL_MS = 5 * 60 * 1000;
const entries = new Map();

function pruneExpired(now) {
  for (const [key, entry] of entries) {
    if (entry.expiresAt <= now) entries.delete(key);
  }
}

export function buildSendKey(parts) {
  return crypto.createHash('sha256').update(JSON.stringify(parts)).digest('hex');
}

export function resetSendDedupe() {
  entries.clear();
}

export async function withSendDedupe(key, sendFn, { ttlMs = DEFAULT_TTL_MS, now = Date.now } = {}) {
  const startedAt = now();
  pruneExpired(startedAt);

  const existing = entries.get(key);
  if (existing) {
    if (existing.state === 'pending') {
      return createToolError(
        'An identical message is already being sent; not sending a duplicate. '
        + 'Wait for the in-progress send to finish, then check Sent Items.',
        { retryable: false, duplicateSuppressed: true }
      );
    }
    return existing.result;
  }

  entries.set(key, { state: 'pending', result: null, expiresAt: startedAt + ttlMs });

  let result;
  try {
    result = await sendFn();
  } catch (error) {
    entries.delete(key);
    throw error;
  }

  const finishedAt = now();
  if (result?.isError) {
    if (result._errorDetails?.ambiguousOutcome) {
      entries.set(key, { state: 'ambiguous', result, expiresAt: finishedAt + ttlMs });
    } else {
      entries.delete(key);
    }
  } else {
    entries.set(key, { state: 'succeeded', result, expiresAt: finishedAt + ttlMs });
  }
  return result;
}
