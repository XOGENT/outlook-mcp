import { createSafeResponse } from '../../utils/jsonUtils.js';
import { convertErrorToToolError } from '../../utils/mcpErrorResponse.js';

/**
 * Generate a tenant-wide admin-consent URL for the configured multi-tenant app.
 * Hand the URL to an administrator of the target tenant; once they approve it,
 * users in that tenant can connect with outlook_connect_account. No per-tenant
 * app registration is needed.
 */
export async function requestAdminConsentTool(registry, args = {}) {
  try {
    const { adminConsentUrl, tenantId, clientId, redirectUri, scopes } = registry.requestAdminConsent(args);
    return createSafeResponse({
      success: true,
      tenantId,
      clientId,
      redirectUri,
      scopes,
      adminConsentUrl,
      message:
        `Send this admin-consent URL to an administrator of tenant "${tenantId}". `
        + 'After they approve it, users in that tenant can run outlook_connect_account. '
        + 'This grants the app org-wide consent — no per-tenant app registration is required. '
        + `The reply URL "${redirectUri}" must be registered as a redirect URI on the app `
        + '(otherwise sign-in fails with AADSTS900971/AADSTS50011); override it with the '
        + 'MCP_OUTLOOK_ADMIN_CONSENT_REDIRECT environment variable.',
    });
  } catch (error) {
    if (error?.isError) return error;
    return convertErrorToToolError(error, 'Failed to build admin-consent URL');
  }
}
