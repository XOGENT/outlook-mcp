import crypto from 'crypto';
import fs from 'fs';
import path from 'path';
import { createToolError } from '../../utils/mcpErrorResponse.js';
import { getDataDir, ensureDir } from '../../auth/dataPaths.js';

const DEFAULT_TTL_MS = 5 * 60 * 1000;
const entries = new Map();

let loaded = false;
let storePathOverride = null;

function storePath() {
  return storePathOverride || path.join(getDataDir(), 'send-dedupe.json');
}

// Test seam: point the journal at a scratch file. Also resets load state.
export function _setSendDedupeStorePath(p) {
  storePathOverride = p;
  loaded = false;
  entries.clear();
}

function ambiguousCrashResult() {
  // A 'pending' entry from a previous process means it started a send and died
  // before recording the outcome; the send may already have gone through.
  return createToolError(
    'A previous send with identical content did not record a result (the server '
    + 'restarted mid-send). The email may already have been sent, so it was NOT '
    + 'sent again to avoid a duplicate. Check Sent Items before retrying.',
    { retryable: false, ambiguousOutcome: true, duplicateSuppressed: true }
  );
}

function loadFromDisk(now) {
  if (loaded) return;
  loaded = true;
  let raw;
  try {
    raw = fs.readFileSync(storePath(), 'utf8');
  } catch (error) {
    if (error.code !== 'ENOENT') {
      console.error('send-dedupe: could not read journal, starting empty:', error.message);
    }
    return;
  }
  let parsed;
  try {
    parsed = JSON.parse(raw);
  } catch (error) {
    console.error('send-dedupe: corrupt journal, ignoring it:', error.message);
    return;
  }
  for (const [key, entry] of Object.entries(parsed || {})) {
    if (!entry || typeof entry.expiresAt !== 'number' || entry.expiresAt <= now) continue;
    if (entry.state === 'pending') {
      entries.set(key, { state: 'ambiguous', result: ambiguousCrashResult(), expiresAt: entry.expiresAt });
    } else if (entry.state === 'succeeded' || entry.state === 'ambiguous') {
      entries.set(key, { state: entry.state, result: entry.result, expiresAt: entry.expiresAt });
    }
  }
}

function persist() {
  const target = storePath();
  const serializable = {};
  for (const [key, entry] of entries) {
    serializable[key] = entry;
  }
  try {
    ensureDir(path.dirname(target));
    const tmp = target + '.' + process.pid + '.tmp';
    fs.writeFileSync(tmp, JSON.stringify(serializable), 'utf8');
    fs.renameSync(tmp, target);
  } catch (error) {
    console.error('send-dedupe: could not persist journal:', error.message);
  }
}

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
  loaded = false;
  try {
    fs.rmSync(storePath(), { force: true });
  } catch {
    // ignore
  }
}

export async function withSendDedupe(key, sendFn, { ttlMs = DEFAULT_TTL_MS, now = Date.now } = {}) {
  const startedAt = now();
  loadFromDisk(startedAt);
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

  // Durably record the in-flight send BEFORE issuing it, so a crash mid-send is
  // recoverable as ambiguous rather than silently re-sendable next process.
  entries.set(key, { state: 'pending', result: null, expiresAt: startedAt + ttlMs });
  persist();

  let result;
  try {
    result = await sendFn();
  } catch (error) {
    entries.delete(key);
    persist();
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
  persist();
  return result;
}
