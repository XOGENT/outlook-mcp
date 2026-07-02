import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import fs from 'fs';
import path from 'path';
import os from 'os';
import { AccountRegistry } from '../../auth/accountRegistry.js';

describe('AccountRegistry', () => {
  let registry;
  let tmpDir;

  beforeEach(() => {
    tmpDir = path.join(os.tmpdir(), `outlook-registry-${Date.now()}`);
    process.env.MCP_OUTLOOK_DATA_DIR = tmpDir;
    registry = new AccountRegistry();
    registry.invalidateCache();
  });

  afterEach(() => {
    fs.rmSync(tmpDir, { recursive: true, force: true });
    delete process.env.MCP_OUTLOOK_DATA_DIR;
  });

  it('adds and lists accounts', async () => {
    await registry.addAccount({
      accountId: 't1:u1',
      tenantId: 't1',
      clientId: 'c1',
      userId: 'u1',
      email: 'a@test.com',
      displayName: 'A',
    });
    const accounts = await registry.listAccounts();
    expect(accounts).toHaveLength(1);
    expect(accounts[0].isDefault).toBe(true);
    expect(accounts[0].email).toBe('a@test.com');
  });

  it('sets default account', async () => {
    await registry.addAccount({ accountId: 't1:u1', tenantId: 't1', clientId: 'c1', userId: 'u1', email: 'a@test.com', displayName: 'A' });
    await registry.addAccount({ accountId: 't2:u2', tenantId: 't2', clientId: 'c1', userId: 'u2', email: 'b@test.com', displayName: 'B' });
    await registry.setDefaultAccount('t2:u2');
    const def = await registry.getDefaultAccount();
    expect(def.accountId).toBe('t2:u2');
  });

  it('removes account', async () => {
    await registry.addAccount({ accountId: 't1:u1', tenantId: 't1', clientId: 'c1', userId: 'u1', email: 'a@test.com', displayName: 'A' });
    await registry.removeAccount('t1:u1');
    expect(await registry.hasAccounts()).toBe(false);
  });
});
