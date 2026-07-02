import { describe, it, expect } from 'vitest';
import { getStartupConfig, resolveAuthAuthority, resolveAuthMode } from '../../auth/defaultApp.js';

describe('hosted OAuth defaults', () => {
  it('uses organizations authority when tenant unset', () => {
    const originalTenant = process.env.AZURE_TENANT_ID;
    delete process.env.AZURE_TENANT_ID;
    expect(resolveAuthAuthority()).toBe('organizations');
    if (originalTenant) process.env.AZURE_TENANT_ID = originalTenant;
  });

  it('returns startup config', () => {
    const config = getStartupConfig();
    expect(config.clientId).toBeTruthy();
    expect(config.authority).toBeTruthy();
    expect(['hosted', 'byo']).toContain(config.authMode);
  });

  it('detects byo mode with env client id', () => {
    const origClient = process.env.AZURE_CLIENT_ID;
    const origTenant = process.env.AZURE_TENANT_ID;
    process.env.AZURE_CLIENT_ID = 'custom-client';
    process.env.AZURE_TENANT_ID = 'custom-tenant';
    expect(resolveAuthMode('custom-client', 'custom-tenant')).toBe('byo');
    if (origClient) process.env.AZURE_CLIENT_ID = origClient; else delete process.env.AZURE_CLIENT_ID;
    if (origTenant) process.env.AZURE_TENANT_ID = origTenant; else delete process.env.AZURE_TENANT_ID;
  });
});
