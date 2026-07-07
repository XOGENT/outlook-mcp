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
// is in use, or when there is nothing to migrate. Safe to call on every startup.
//
// The guard keys off the *registry file*, not the target directory: the target
// may already exist as an empty/partial dir (e.g. a registry read created
// `accounts/` before the old store was moved), and that must NOT block the
// migration — otherwise the accounts stay stranded in the old location.
export function migrateInstallDataDir(from = getInstallDataDir(), to = getDataDir()) {
  if (process.env.MCP_OUTLOOK_DATA_DIR || isHeadlessMode()) return;
  if (path.resolve(from) === path.resolve(to)) return;
  if (!fs.existsSync(from)) return;
  // Never clobber a destination that already holds a registry — that store wins.
  if (fs.existsSync(path.join(to, 'accounts', 'registry.json'))) return;
  // Copy contents in (merging into any pre-existing empty dir) rather than
  // renaming, so a partially-created target and cross-filesystem moves both work.
  fs.cpSync(from, to, { recursive: true, force: false, errorOnExist: false });
  fs.rmSync(from, { recursive: true, force: true });
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
