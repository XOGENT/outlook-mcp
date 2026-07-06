import { describe, it, expect } from 'vitest';
import { isClientIdConfigured, PLACEHOLDER_CLIENT_ID } from '../../auth/defaultApp.js';
import { authManagerRegistry } from '../../auth/authManagerRegistry.js';

describe('client ID configuration guard', () => {
  it('treats the placeholder client ID as unconfigured', () => {
    expect(isClientIdConfigured(PLACEHOLDER_CLIENT_ID)).toBe(false);
    expect(isClientIdConfigured('')).toBe(false);
    expect(isClientIdConfigured(undefined)).toBe(false);
  });

  it('treats a real client ID as configured', () => {
    expect(isClientIdConfigured('11111111-2222-3333-4444-555555555555')).toBe(true);
  });

  it('connectAccount refuses to start OAuth without a real client ID', async () => {
    const result = await authManagerRegistry.connectAccount({ clientId: PLACEHOLDER_CLIENT_ID });
    expect(result.success).toBe(false);
    expect(result.error?.isError).toBe(true);
    const text = result.error.content?.[0]?.text || '';
    expect(text).toContain('No Azure application is configured');
  });
});

describe('requestAdminConsent', () => {
  const CLIENT = '11111111-2222-3333-4444-555555555555';

  it('throws without a real client ID', () => {
    expect(() => authManagerRegistry.requestAdminConsent({
      tenantId: 'contoso.com',
      clientId: PLACEHOLDER_CLIENT_ID,
    })).toThrow();
  });

  it('requires a specific tenant', () => {
    expect(() => authManagerRegistry.requestAdminConsent({ clientId: CLIENT })).toThrow();
    expect(() => authManagerRegistry.requestAdminConsent({ clientId: CLIENT, tenantId: 'organizations' })).toThrow();
  });

  it('builds a tenant-scoped admin-consent URL for the configured app', () => {
    const result = authManagerRegistry.requestAdminConsent({ tenantId: 'contoso.com', clientId: CLIENT });
    const url = new URL(result.adminConsentUrl);
    expect(url.origin + url.pathname).toBe('https://login.microsoftonline.com/contoso.com/adminconsent');
    expect(url.searchParams.get('client_id')).toBe(CLIENT);
    expect(url.searchParams.get('state')).toBeTruthy();
    expect(result.tenantId).toBe('contoso.com');
  });

  it('always includes a redirect_uri so sign-in does not fail with AADSTS900971', () => {
    const result = authManagerRegistry.requestAdminConsent({ tenantId: 'contoso.com', clientId: CLIENT });
    const url = new URL(result.adminConsentUrl);
    const redirectUri = url.searchParams.get('redirect_uri');
    expect(redirectUri).toBeTruthy();
    expect(result.redirectUri).toBe(redirectUri);
  });
});
