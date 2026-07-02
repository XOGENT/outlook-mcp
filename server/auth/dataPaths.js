import path from 'path';
import { fileURLToPath } from 'url';
import fs from 'fs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const PROJECT_ROOT = path.join(__dirname, '../..');

export function getDataDir() {
  if (process.env.MCP_OUTLOOK_DATA_DIR) {
    return path.resolve(process.env.MCP_OUTLOOK_DATA_DIR);
  }
  if (isHeadlessMode()) {
    return '/data';
  }
  return path.join(PROJECT_ROOT, '.tokens');
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
  return path.join(PROJECT_ROOT, '.tokens');
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
