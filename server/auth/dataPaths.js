import path from 'path';
import { fileURLToPath } from 'url';
import fs from 'fs';
import os from 'os';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const PROJECT_ROOT = path.join(__dirname, '../..');

export function getDataDir() {
  if (process.env.MCP_OUTLOOK_DATA_DIR) {
    return path.resolve(process.env.MCP_OUTLOOK_DATA_DIR);
  }
  if (isHeadlessMode()) {
    return '/data';
  }
  // Persist under a stable per-user location, NOT the install directory. When
  // packaged as a .dxt, the server runs from a Claude Desktop-managed extension
  // folder that is re-extracted (wiped) on every update/reinstall — storing
  // tokens there loses every connection across restarts.
  return path.join(os.homedir(), '.outlook-mcp');
}

// Historical install-relative data dir, retained only so migrateInstallDataDir()
// can move an existing store to the stable per-user location.
export function getInstallDataDir() {
  return path.join(PROJECT_ROOT, '.tokens');
}

// One-time move of a pre-existing install-relative store (<install>/.tokens) to
// the stable per-user data dir. No-ops when an env override or headless volume
// is in use, when the store was already migrated, or when there is nothing to
// migrate. Safe to call on every startup.
export function migrateInstallDataDir(from = getInstallDataDir(), to = getDataDir()) {
  if (process.env.MCP_OUTLOOK_DATA_DIR || isHeadlessMode()) return;
  const target = to;
  const legacy = from;
  if (path.resolve(target) === path.resolve(legacy)) return;
  if (fs.existsSync(target)) return;
  if (!fs.existsSync(legacy)) return;
  ensureDir(path.dirname(target));
  try {
    fs.renameSync(legacy, target);
  } catch (error) {
    if (error.code !== 'EXDEV') throw error;
    // Old store lives on a different filesystem — copy then remove.
    fs.cpSync(legacy, target, { recursive: true });
    fs.rmSync(legacy, { recursive: true, force: true });
  }
}

export function getAccountsDir() {
  return path.join(getDataDir(), 'accounts');
}

export function getAccountDir(accountId) {
  return path.join(getAccountsDir(), sanitizeAccountId(accountId));
}

export function getRegistryPath() {
  return path.join(getAccountsDir(), 'registry.json');
}

export function getLegacyTokensDir() {
  if (isHeadlessMode() || process.env.MCP_OUTLOOK_DATA_DIR) {
    return path.join(getDataDir(), '.legacy-tokens');
  }
  // The pre-multi-account single-account token store was rooted at the data dir
  // itself; keep resolving it there (now the stable per-user dir).
  return getDataDir();
}

export function getWorkDir() {
  if (process.env.MCP_OUTLOOK_WORK_DIR) {
    return path.resolve(process.env.MCP_OUTLOOK_WORK_DIR);
  }
  if (process.env.MCP_OUTLOOK_DATA_DIR) {
    return path.join(getDataDir(), 'downloads');
  }
  return path.join(PROJECT_ROOT, '.downloads');
}

export function sanitizeAccountId(accountId) {
  return accountId.replace(/[^a-zA-Z0-9._@-]/g, '_');
}

export function ensureDir(dirPath) {
  if (!fs.existsSync(dirPath)) {
    fs.mkdirSync(dirPath, { recursive: true });
  }
}

export function isHeadlessMode() {
  if (process.env.MCP_OUTLOOK_HEADLESS === 'true') return true;
  if (process.env.CI === 'true') return true;
  return false;
}
