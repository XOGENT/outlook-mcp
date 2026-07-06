/**
 * Hosted multi-tenant Azure AD app defaults.
 * Set OUTLOOK_MCP_CLIENT_ID env or AZURE_CLIENT_ID for BYO registration.
 */
export const PLACEHOLDER_CLIENT_ID = '00000000-0000-0000-0000-000000000000';

export const HOSTED_CLIENT_ID = process.env.OUTLOOK_MCP_CLIENT_ID
  || process.env.AZURE_CLIENT_ID
  || PLACEHOLDER_CLIENT_ID;

export const ORGANIZATIONS_AUTHORITY = 'organizations';

/**
 * True when a real Azure application client ID is available. The shipped
 * default is a placeholder, so OAuth cannot work until a real (multi-tenant)
 * app client ID is supplied via AZURE_CLIENT_ID / OUTLOOK_MCP_CLIENT_ID or the
 * connect tool's clientId argument.
 */
export function isClientIdConfigured(clientId) {
  const id = clientId || HOSTED_CLIENT_ID;
  return Boolean(id) && id !== PLACEHOLDER_CLIENT_ID;
}

/**
 * Human-readable guidance shown when no Azure app is configured.
 */
export const NO_APP_CONFIGURED_MESSAGE =
  'No Azure application is configured, so Microsoft sign-in cannot start. '
  + 'Register ONE multi-tenant Azure AD app (once), then set its Application (client) ID '
  + 'via the AZURE_CLIENT_ID environment variable (or pass clientId to outlook_connect_account). '
  + 'Additional tenants are then onboarded with outlook_request_admin_consent — no per-tenant registration needed.';

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
