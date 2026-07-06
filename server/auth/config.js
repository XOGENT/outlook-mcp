/**
 * Delegated Graph scopes that individual users can consent to themselves
 * (Microsoft Graph "Admin consent required: No"). These are the default so a
 * tenant that allows user consent needs no admin involvement at all.
 */
export const CORE_DELEGATED_SCOPES = [
  'Mail.Read',
  'Mail.ReadWrite',
  'Mail.Send',
  'Mail.Read.Shared',
  'Mail.ReadWrite.Shared',
  'Calendars.Read',
  'Calendars.ReadWrite',
  'Calendars.Read.Shared',
  'Calendars.ReadWrite.Shared',
  'Contacts.Read',
  'Contacts.ReadWrite',
  'Tasks.Read',
  'Tasks.ReadWrite',
  'User.Read',
  'MailboxSettings.Read',
  'Files.Read.All',         // delegated: files the signed-in user can access (user-consentable)
  'Files.ReadWrite.All',
  'offline_access',         // required for refresh tokens
];

/**
 * SharePoint site scopes. These require TENANT ADMIN consent, so they are
 * opt-in via OUTLOOK_MCP_ENABLE_SHAREPOINT=true. When enabled, onboard each
 * tenant with outlook_request_admin_consent before users connect.
 */
export const SHAREPOINT_SCOPES = [
  'Sites.Read.All',
  'Sites.ReadWrite.All',
];

/**
 * Resolve the effective delegated scope string. Defaults to the
 * admin-consent-free core set; SharePoint scopes and arbitrary extras are opt-in.
 */
export function resolveScopes() {
  const scopes = [...CORE_DELEGATED_SCOPES];
  if (process.env.OUTLOOK_MCP_ENABLE_SHAREPOINT === 'true') {
    scopes.push(...SHAREPOINT_SCOPES);
  }
  if (process.env.OUTLOOK_MCP_EXTRA_SCOPES) {
    scopes.push(...process.env.OUTLOOK_MCP_EXTRA_SCOPES.split(/[\s,]+/).filter(Boolean));
  }
  // De-duplicate while preserving order.
  return [...new Set(scopes)].join(' ');
}

export const authConfig = {
  oauth: {
    authorizeUrl: (tenantId) =>
      `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize`,
    tokenUrl: (tenantId) =>
      `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
    adminConsentUrl: (tenantId) =>
      `https://login.microsoftonline.com/${tenantId}/adminconsent`,
    // Reply URL the admin-consent endpoint redirects to after approval. Must be
    // registered on the app registration's redirect URIs. Microsoft's hosted
    // native-client page is the safe default (shows the consent result).
    adminConsentRedirectUri:
      process.env.MCP_OUTLOOK_ADMIN_CONSENT_REDIRECT
      || 'https://login.microsoftonline.com/common/oauth2/nativeclient',
    scope: resolveScopes(),
    redirectUri: process.env.MCP_OUTLOOK_REDIRECT_URI || 'http://localhost:0/callback',
  },

  token: {
    accessTokenTTL: 60 * 60 * 1000, // 60 minutes in milliseconds
    refreshThreshold: 55 * 60 * 1000, // Refresh at 55 minutes
    refreshTokenTTL: 90 * 24 * 60 * 60 * 1000, // 90 days
  },

  retry: {
    maxAttempts: 3,
    initialDelay: 1000, // 1 second
    maxDelay: 30000, // 30 seconds
    backoffMultiplier: 2,
  },

  security: {
    usePKCE: true,        // PKCE ensures secure authentication without client secrets
    encryptTokens: true,  // Tokens are encrypted in storage
    auditLogging: true,   // All authentication events are logged
  },
};