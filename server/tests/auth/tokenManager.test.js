import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import fs from 'fs';
import path from 'path';
import os from 'os';
import { TokenManager, resetTokenStorageCache } from '../../auth/tokenManager.js';

describe('TokenManager', () => {
  let tmpDir;

  beforeEach(() => {
    tmpDir = path.join(os.tmpdir(), `outlook-tokens-${Date.now()}`);
    process.env.MCP_OUTLOOK_DATA_DIR = tmpDir;
    process.env.MCP_OUTLOOK_HEADLESS = 'true';
  });

  afterEach(() => {
    resetTokenStorageCache();
    fs.rmSync(tmpDir, { recursive: true, force: true });
    delete process.env.MCP_OUTLOOK_DATA_DIR;
    delete process.env.MCP_OUTLOOK_HEADLESS;
  });

  it('stores and retrieves tokens per account namespace', async () => {
    const tm1 = new TokenManager('client-a', 'tenant:oid1-ns');
    const tm2 = new TokenManager('client-a', 'tenant:oid2-ns');

    await tm1.storeTokens('access-1', 'refresh-1', 3600);
    await tm2.storeTokens('access-2', 'refresh-2', 3600);

    expect(await tm1.getAccessToken()).toBe('access-1');
    expect(await tm2.getAccessToken()).toBe('access-2');
  });

  it('round-trips encrypted tokens', async () => {
    const tm = new TokenManager('client-a', 'tenant:oid-roundtrip');
    await tm.storeTokens('secret-token-value', 'refresh-value', 3600);
    const access = await tm.getAccessToken();
    const refresh = await tm.getRefreshToken();
    expect(access).toBe('secret-token-value');
    expect(refresh).toBe('refresh-value');
  });

  // Regression: the storage dir can contain subdirectories (e.g. the legacy
  // store is rooted at the data dir, which also holds `accounts/`). node-persist's
  // background expired-keys sweep would readFile() those subdirs and crash with
  // an unhandled "EISDIR: illegal operation on a directory, read". The interval
  // must stay disabled so the sweep never runs.
  it('disables the expired-keys sweep so subdirectories never trigger EISDIR', async () => {
    const tm = new TokenManager('client-a', 'tenant:oid-nosweep');
    await tm.storeTokens('access', 'refresh', 3600);

    expect(tm.storage.options.expiredInterval).toBe(false);

    // A directory inside the storage dir would blow up node-persist's dir scan;
    // prove the scan is exactly what fails, and that we never invoke it.
    const storageDir = tm.storage.options.dir;
    fs.mkdirSync(path.join(storageDir, 'accounts'), { recursive: true });
    await expect(tm.storage.keys()).rejects.toMatchObject({ code: 'EISDIR' });

    // Normal per-key access is unaffected by the sibling subdirectory.
    expect(await tm.getAccessToken()).toBe('access');
  });
});
