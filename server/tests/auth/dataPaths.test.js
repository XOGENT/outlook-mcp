import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import fs from 'fs';
import path from 'path';
import os from 'os';
import { getDataDir, getAccountsDir, getAccountDir, getRegistryPath, getWorkDir, getLegacyTokensDir, getInstallDataDir, migrateInstallDataDir } from '../../auth/dataPaths.js';

describe('dataPaths', () => {
  let originalDataDir;

  beforeEach(() => {
    originalDataDir = process.env.MCP_OUTLOOK_DATA_DIR;
    process.env.MCP_OUTLOOK_DATA_DIR = path.join(os.tmpdir(), `outlook-mcp-test-${Date.now()}`);
  });

  afterEach(() => {
    if (process.env.MCP_OUTLOOK_DATA_DIR) {
      fs.rmSync(process.env.MCP_OUTLOOK_DATA_DIR, { recursive: true, force: true });
    }
    if (originalDataDir) {
      process.env.MCP_OUTLOOK_DATA_DIR = originalDataDir;
    } else {
      delete process.env.MCP_OUTLOOK_DATA_DIR;
    }
  });

  it('resolves data directory from env', () => {
    expect(getDataDir()).toBe(process.env.MCP_OUTLOOK_DATA_DIR);
    expect(getAccountsDir()).toBe(path.join(process.env.MCP_OUTLOOK_DATA_DIR, 'accounts'));
    expect(getRegistryPath()).toBe(path.join(process.env.MCP_OUTLOOK_DATA_DIR, 'accounts', 'registry.json'));
  });

  it('resolves account and work directories', () => {
    expect(getAccountDir('tenant:user')).toContain('tenant_user');
    expect(getWorkDir()).toBe(path.join(process.env.MCP_OUTLOOK_DATA_DIR, 'downloads'));
  });

  it('uses data volume for legacy tokens in headless mode', () => {
    delete process.env.MCP_OUTLOOK_DATA_DIR;
    process.env.MCP_OUTLOOK_HEADLESS = 'true';
    expect(getDataDir()).toBe('/data');
    expect(getLegacyTokensDir()).toBe('/data/.legacy-tokens');
    delete process.env.MCP_OUTLOOK_HEADLESS;
  });

  it('defaults to a stable per-user dir, not the install directory', () => {
    delete process.env.MCP_OUTLOOK_DATA_DIR;
    delete process.env.MCP_OUTLOOK_HEADLESS;
    const dataDir = getDataDir();
    expect(dataDir).toBe(path.join(os.homedir(), '.outlook-mcp'));
    expect(dataDir).not.toBe(getInstallDataDir());
    // Legacy single-account store resolves at the (stable) data dir root.
    expect(getLegacyTokensDir()).toBe(dataDir);
  });
});

describe('migrateInstallDataDir', () => {
  let root;
  let savedEnv;

  beforeEach(() => {
    savedEnv = { data: process.env.MCP_OUTLOOK_DATA_DIR, ci: process.env.CI, headless: process.env.MCP_OUTLOOK_HEADLESS };
    delete process.env.MCP_OUTLOOK_DATA_DIR;
    delete process.env.CI;
    delete process.env.MCP_OUTLOOK_HEADLESS;
    root = path.join(os.tmpdir(), `outlook-mcp-migrate-${Date.now()}-${Math.floor(process.hrtime()[1])}`);
    fs.mkdirSync(root, { recursive: true });
  });

  afterEach(() => {
    fs.rmSync(root, { recursive: true, force: true });
    if (savedEnv.data) process.env.MCP_OUTLOOK_DATA_DIR = savedEnv.data;
    if (savedEnv.ci) process.env.CI = savedEnv.ci;
    if (savedEnv.headless) process.env.MCP_OUTLOOK_HEADLESS = savedEnv.headless;
  });

  it('moves an existing install store to the new location', () => {
    const from = path.join(root, 'install', '.tokens');
    const to = path.join(root, 'home', '.outlook-mcp');
    fs.mkdirSync(path.join(from, 'accounts'), { recursive: true });
    fs.writeFileSync(path.join(from, 'accounts', 'registry.json'), '{"accounts":[{"email":"a@b.com"}]}');

    migrateInstallDataDir(from, to);

    expect(fs.existsSync(from)).toBe(false);
    expect(fs.readFileSync(path.join(to, 'accounts', 'registry.json'), 'utf8')).toContain('a@b.com');
  });

  it('does not overwrite an existing new-location store', () => {
    const from = path.join(root, 'install', '.tokens');
    const to = path.join(root, 'home', '.outlook-mcp');
    fs.mkdirSync(from, { recursive: true });
    fs.writeFileSync(path.join(from, 'marker'), 'old');
    fs.mkdirSync(to, { recursive: true });
    fs.writeFileSync(path.join(to, 'marker'), 'new');

    migrateInstallDataDir(from, to);

    expect(fs.readFileSync(path.join(to, 'marker'), 'utf8')).toBe('new');
    expect(fs.existsSync(from)).toBe(true);
  });

  it('no-ops when an explicit data dir is configured', () => {
    process.env.MCP_OUTLOOK_DATA_DIR = path.join(root, 'env');
    const from = path.join(root, 'install', '.tokens');
    const to = path.join(root, 'home', '.outlook-mcp');
    fs.mkdirSync(from, { recursive: true });

    migrateInstallDataDir(from, to);

    expect(fs.existsSync(to)).toBe(false);
    expect(fs.existsSync(from)).toBe(true);
  });
});
