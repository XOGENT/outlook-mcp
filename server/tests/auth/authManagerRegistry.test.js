import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import fs from 'fs';
import path from 'path';
import os from 'os';
import { AuthManagerRegistry } from '../../auth/authManagerRegistry.js';
import { accountRegistry } from '../../auth/accountRegistry.js';
import { resetTokenStorageCache } from '../../auth/tokenManager.js';

const sampleAccount = (id, email) => ({
  accountId: id,
  tenantId: `tenant-${id}`,
  clientId: 'test-client',
  authMode: 'hosted',
  userId: `user-${id}`,
  email,
  displayName: email.split('@')[0],
});

describe('AuthManagerRegistry', () => {
  let registry;
  let tmpDir;

  beforeEach(() => {
    tmpDir = path.join(os.tmpdir(), `outlook-auth-registry-${Date.now()}`);
    process.env.MCP_OUTLOOK_DATA_DIR = tmpDir;
    process.env.MCP_OUTLOOK_HEADLESS = 'true';
    accountRegistry.invalidateCache();
    resetTokenStorageCache();
    registry = new AuthManagerRegistry();
    registry.migrated = true;
  });

  afterEach(() => {
    resetTokenStorageCache();
    fs.rmSync(tmpDir, { recursive: true, force: true });
    delete process.env.MCP_OUTLOOK_DATA_DIR;
    delete process.env.MCP_OUTLOOK_HEADLESS;
  });

  it('resolve throws when no accounts are connected', async () => {
    await expect(registry.resolve()).rejects.toMatchObject({ isError: true });
    try {
      await registry.resolve();
    } catch (error) {
      expect(error.content[0].text).toContain('outlook_connect_account');
    }
  });

  it('resolve returns the default account when accountId is omitted', async () => {
    await accountRegistry.addAccount(sampleAccount('t1:u1', 'a@test.com'));
    await accountRegistry.addAccount(sampleAccount('t2:u2', 'b@test.com'));
    await accountRegistry.setDefaultAccount('t2:u2');

    const { account, manager } = await registry.resolve();
    expect(account.accountId).toBe('t2:u2');
    expect(manager.accountId).toBe('t2:u2');
  });

  it('resolve returns a specific account by id', async () => {
    await accountRegistry.addAccount(sampleAccount('t1:u1', 'a@test.com'));
    await accountRegistry.addAccount(sampleAccount('t2:u2', 'b@test.com'));

    const { account } = await registry.resolve('t1:u1');
    expect(account.email).toBe('a@test.com');
  });

  it('resolve throws for an unknown account id', async () => {
    await accountRegistry.addAccount(sampleAccount('t1:u1', 'a@test.com'));
    await expect(registry.resolve('missing:account')).rejects.toMatchObject({ isError: true });
  });

  it('resolveAll returns every connected account', async () => {
    await accountRegistry.addAccount(sampleAccount('t1:u1', 'a@test.com'));
    await accountRegistry.addAccount(sampleAccount('t2:u2', 'b@test.com'));

    const accounts = await registry.resolveAll();
    expect(accounts).toHaveLength(2);
    expect(accounts.map(a => a.accountId).sort()).toEqual(['t1:u1', 't2:u2']);
  });

  it('getOrCreateManager caches manager instances per account', async () => {
    await accountRegistry.addAccount(sampleAccount('t1:u1', 'a@test.com'));
    const account = await accountRegistry.getAccount('t1:u1');

    const manager1 = await registry.getOrCreateManager(account);
    const manager2 = await registry.getOrCreateManager(account);
    expect(manager1).toBe(manager2);
  });

  it('setDefaultAccount updates the default account', async () => {
    await accountRegistry.addAccount(sampleAccount('t1:u1', 'a@test.com'));
    await accountRegistry.addAccount(sampleAccount('t2:u2', 'b@test.com'));

    const result = await registry.setDefaultAccount('t2:u2');
    expect(result.success).toBe(true);
    expect(result.account.accountId).toBe('t2:u2');

    const { account } = await registry.resolve();
    expect(account.accountId).toBe('t2:u2');
  });

  it('removeAccount deletes the account from the registry', async () => {
    await accountRegistry.addAccount(sampleAccount('t1:u1', 'a@test.com'));
    const result = await registry.removeAccount('t1:u1');
    expect(result.success).toBe(true);
    expect(await registry.hasAccounts()).toBe(false);
  });

  it('evictManager drops a cached manager instance', async () => {
    await accountRegistry.addAccount(sampleAccount('t1:u1', 'a@test.com'));
    const account = await accountRegistry.getAccount('t1:u1');

    const manager1 = await registry.getOrCreateManager(account);
    registry.evictManager('t1:u1');
    const manager2 = await registry.getOrCreateManager(account);
    expect(manager1).not.toBe(manager2);
  });
});
