import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import fs from 'fs';
import path from 'path';
import os from 'os';
import { AuthManagerRegistry } from '../../auth/authManagerRegistry.js';
import { accountRegistry } from '../../auth/accountRegistry.js';
import { OutlookAuthManager } from '../../auth/auth.js';
import { LegacyTokenManager, TokenManager, resetTokenStorageCache } from '../../auth/tokenManager.js';
import * as jwtUtils from '../../auth/jwtUtils.js';

describe('legacy token migration', () => {
  let tmpDir;

  beforeEach(() => {
    tmpDir = path.join(os.tmpdir(), `outlook-migration-${Date.now()}`);
    process.env.MCP_OUTLOOK_DATA_DIR = tmpDir;
    process.env.MCP_OUTLOOK_HEADLESS = 'true';
    accountRegistry.invalidateCache();
    resetTokenStorageCache();
    vi.restoreAllMocks();
  });

  afterEach(() => {
    resetTokenStorageCache();
    fs.rmSync(tmpDir, { recursive: true, force: true });
    delete process.env.MCP_OUTLOOK_DATA_DIR;
    delete process.env.MCP_OUTLOOK_HEADLESS;
  });

  it('migrates legacy tokens into a namespaced account', async () => {
    vi.spyOn(LegacyTokenManager.prototype, 'hasLegacyTokens').mockResolvedValue(true);
    vi.spyOn(LegacyTokenManager.prototype, 'getAccessToken').mockResolvedValue('legacy-access-token');
    vi.spyOn(LegacyTokenManager.prototype, 'getRefreshToken').mockResolvedValue('legacy-refresh-token');
    vi.spyOn(LegacyTokenManager.prototype, 'getTokenMetadata').mockResolvedValue({
      accessTokenExpiry: Date.now() + 3600 * 1000,
    });
    const clearLegacySpy = vi.spyOn(LegacyTokenManager.prototype, 'clearTokens').mockResolvedValue();

    vi.spyOn(jwtUtils, 'extractAccountClaims').mockReturnValue({
      accountId: 'tenant-1:oid-1',
      tenantId: 'tenant-1',
      userId: 'oid-1',
      email: 'user@contoso.com',
      displayName: 'User',
    });

    vi.spyOn(OutlookAuthManager.prototype, 'initializeGraphClient').mockResolvedValue();
    vi.spyOn(OutlookAuthManager.prototype, 'validateAuthentication').mockResolvedValue({
      success: true,
      user: { id: 'oid-1', mail: 'user@contoso.com', displayName: 'User' },
    });

    const storeSpy = vi.spyOn(TokenManager.prototype, 'storeTokens').mockResolvedValue();

    const registry = new AuthManagerRegistry();
    await registry.initialize();

    expect(await accountRegistry.hasAccounts()).toBe(true);
    const accounts = await accountRegistry.listAccounts();
    expect(accounts).toHaveLength(1);
    expect(accounts[0].accountId).toBe('tenant-1:oid-1');
    expect(accounts[0].email).toBe('user@contoso.com');
    expect(storeSpy).toHaveBeenCalledWith('legacy-access-token', 'legacy-refresh-token', expect.any(Number));
    expect(clearLegacySpy).toHaveBeenCalled();
  });

  it('skips migration when accounts already exist', async () => {
    await accountRegistry.addAccount({
      accountId: 'existing:acct',
      tenantId: 'existing',
      clientId: 'client',
      userId: 'acct',
      email: 'existing@test.com',
      displayName: 'Existing',
    });

    const hasLegacySpy = vi.spyOn(LegacyTokenManager.prototype, 'hasLegacyTokens');

    const registry = new AuthManagerRegistry();
    await registry.initialize();

    expect(hasLegacySpy).not.toHaveBeenCalled();
    const accounts = await accountRegistry.listAccounts();
    expect(accounts).toHaveLength(1);
    expect(accounts[0].accountId).toBe('existing:acct');
  });

  it('skips migration when no legacy tokens are present', async () => {
    vi.spyOn(LegacyTokenManager.prototype, 'hasLegacyTokens').mockResolvedValue(false);
    const storeSpy = vi.spyOn(TokenManager.prototype, 'storeTokens');

    const registry = new AuthManagerRegistry();
    await registry.initialize();

    expect(await accountRegistry.hasAccounts()).toBe(false);
    expect(storeSpy).not.toHaveBeenCalled();
  });
});
