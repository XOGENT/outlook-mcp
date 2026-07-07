import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import fs from 'fs';
import path from 'path';
import os from 'os';
import { OutlookAuthManager } from '../../auth/auth.js';
import { TokenManager, resetTokenStorageCache } from '../../auth/tokenManager.js';

describe('cold-start token refresh', () => {
  let manager;

  beforeEach(() => {
    manager = new OutlookAuthManager({ accountId: 't1:u1', clientId: 'client', tenantId: 't1' });
  });

  afterEach(() => {
    vi.restoreAllMocks();
  });

  it('refreshes when the access token is expired but a refresh token exists', async () => {
    vi.spyOn(manager.tokenManager, 'isAuthenticated').mockResolvedValue(false);
    vi.spyOn(manager.tokenManager, 'hasRefreshToken').mockResolvedValue(true);
    const refreshSpy = vi.spyOn(manager, 'refreshAccessToken').mockResolvedValue(true);
    vi.spyOn(manager, 'initializeGraphClient').mockResolvedValue();
    vi.spyOn(manager, 'validateAuthentication').mockResolvedValue({ success: true, user: { mail: 'a@b.com' } });

    const result = await manager.authenticate();

    expect(refreshSpy).toHaveBeenCalledOnce();
    expect(result.success).toBe(true);
  });

  it('reports no account only when there is no usable refresh token', async () => {
    vi.spyOn(manager.tokenManager, 'isAuthenticated').mockResolvedValue(false);
    vi.spyOn(manager.tokenManager, 'hasRefreshToken').mockResolvedValue(false);
    const refreshSpy = vi.spyOn(manager, 'refreshAccessToken');

    const result = await manager.authenticate();

    expect(refreshSpy).not.toHaveBeenCalled();
    expect(result.success).toBe(false);
    expect(result.error.content[0].text).toContain('No account connected');
  });
});

describe('refreshAccessToken credential preservation', () => {
  let manager;
  let originalFetch;

  beforeEach(() => {
    originalFetch = global.fetch;
    manager = new OutlookAuthManager({ accountId: 't1:u1', clientId: 'client', tenantId: 't1' });
    vi.spyOn(manager.tokenManager, 'getRefreshToken').mockResolvedValue('stored-refresh-token');
  });

  afterEach(() => {
    global.fetch = originalFetch;
    vi.restoreAllMocks();
  });

  it('keeps stored credentials on a transient (5xx) failure', async () => {
    const clearSpy = vi.spyOn(manager.tokenManager, 'clearTokens').mockResolvedValue();
    global.fetch = vi.fn().mockResolvedValue({ ok: false, status: 503, text: async () => 'temporarily unavailable' });

    await expect(manager.refreshAccessToken()).rejects.toBeDefined();
    expect(clearSpy).not.toHaveBeenCalled();
  });

  it('keeps stored credentials on a network error', async () => {
    const clearSpy = vi.spyOn(manager.tokenManager, 'clearTokens').mockResolvedValue();
    global.fetch = vi.fn().mockRejectedValue(new Error('ECONNRESET'));

    await expect(manager.refreshAccessToken()).rejects.toBeDefined();
    expect(clearSpy).not.toHaveBeenCalled();
  });

  it('clears credentials only when the refresh token is definitively rejected', async () => {
    const clearSpy = vi.spyOn(manager.tokenManager, 'clearTokens').mockResolvedValue();
    global.fetch = vi.fn().mockResolvedValue({ ok: false, status: 400, text: async () => '{"error":"invalid_grant"}' });

    await expect(manager.refreshAccessToken()).rejects.toBeDefined();
    expect(clearSpy).toHaveBeenCalledOnce();
  });
});

describe('TokenManager.hasRefreshToken', () => {
  let tmpDir;

  beforeEach(() => {
    tmpDir = path.join(os.tmpdir(), `outlook-hasrt-${Date.now()}-${process.pid}`);
    process.env.MCP_OUTLOOK_DATA_DIR = tmpDir;
    process.env.MCP_OUTLOOK_HEADLESS = 'true'; // disable keytar -> fallback file storage
    resetTokenStorageCache();
  });

  afterEach(() => {
    resetTokenStorageCache();
    fs.rmSync(tmpDir, { recursive: true, force: true });
    delete process.env.MCP_OUTLOOK_DATA_DIR;
    delete process.env.MCP_OUTLOOK_HEADLESS;
  });

  it('is true after storing tokens and false when nothing is stored', async () => {
    const stored = new TokenManager('client', 't1:with-token');
    await stored.storeTokens('access', 'refresh', 3600);
    expect(await stored.hasRefreshToken()).toBe(true);

    const empty = new TokenManager('client', 't1:no-token');
    expect(await empty.hasRefreshToken()).toBe(false);
  });
});
