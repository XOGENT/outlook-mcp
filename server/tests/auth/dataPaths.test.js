import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import fs from 'fs';
import path from 'path';
import os from 'os';
import { getDataDir, getAccountsDir, getAccountDir, getRegistryPath, getWorkDir, getLegacyTokensDir } from '../../auth/dataPaths.js';

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
});
