/**
 * Hosted multi-tenant Azure AD app defaults.
 * Set OUTLOOK_MCP_CLIENT_ID env or AZURE_CLIENT_ID for BYO registration.
 */
export const HOSTED_CLIENT_ID = process.env.OUTLOOK_MCP_CLIENT_ID
  || process.env.AZURE_CLIENT_ID
  || '00000000-0000-0000-0000-000000000000';

export const ORGANIZATIONS_AUTHORITY = 'organizations';

export function resolveClientId(override) {
  return override || HOSTED_CLIENT_ID;
}

export function resolveAuthAuthority(tenantId) {
  if (tenantId && tenantId !== ORGANIZATIONS_AUTHORITY) {
    return tenantId;
  }
  if (process.env.AZURE_TENANT_ID) {
    return process.env.AZURE_TENANT_ID;
  }
  return ORGANIZATIONS_AUTHORITY;
}

export function resolveAuthMode(clientId, tenantId) {
  const usingHostedClient = !process.env.AZURE_CLIENT_ID
    && !process.env.OUTLOOK_MCP_CLIENT_ID
    && clientId === HOSTED_CLIENT_ID;
  const usingHostedAuthority = !tenantId || tenantId === ORGANIZATIONS_AUTHORITY;
  if (usingHostedClient && usingHostedAuthority) return 'hosted';
  return 'byo';
}

export function getStartupConfig() {
  const clientId = resolveClientId();
  const authority = resolveAuthAuthority();
  const authMode = resolveAuthMode(clientId, authority);
  return { clientId, authority, authMode };
}
