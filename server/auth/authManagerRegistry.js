import { OutlookAuthManager } from './auth.js';
import { accountRegistry } from './accountRegistry.js';
import { LegacyTokenManager } from './tokenManager.js';
import {
  getStartupConfig,
  resolveClientId,
  resolveAuthAuthority,
  resolveAuthMode,
  isClientIdConfigured,
  NO_APP_CONFIGURED_MESSAGE,
  ORGANIZATIONS_AUTHORITY,
} from './defaultApp.js';
import { authConfig } from './config.js';
import { extractAccountClaims } from './jwtUtils.js';
import { createAuthError } from '../utils/mcpErrorResponse.js';
import { clearStylingCache, clearSignatureCache } from '../tools/common/sharedUtils.js';
import crypto from 'crypto';

const pendingAuthSessions = new Map();
const managerCache = new Map();

export class AuthManagerRegistry {
  constructor() {
    this.startupConfig = getStartupConfig();
    this.migrated = false;
  }

  async initialize() {
    if (!this.migrated) {
      await this.migrateLegacyTokens();
      this.migrated = true;
    }
  }

  async migrateLegacyTokens() {
    const hasAccounts = await accountRegistry.hasAccounts();
    if (hasAccounts) return;

    const clientId = this.startupConfig.clientId;
    const legacyManager = new LegacyTokenManager(clientId);
    if (!(await legacyManager.hasLegacyTokens())) return;

    try {
      const tenantId = process.env.AZURE_TENANT_ID || 'organizations';
      const authManager = new OutlookAuthManager({
        accountId: null,
        clientId,
        tenantId,
        tokenManager: legacyManager,
      });
      await authManager.initializeGraphClient();
      const result = await authManager.validateAuthentication();
      if (!result.success) return;

      const accessToken = await legacyManager.getAccessToken();
      const claims = extractAccountClaims(accessToken);
      const accountId = claims?.accountId || `legacy:${result.user.id}`;

      const newManager = new OutlookAuthManager({
        accountId,
        clientId,
        tenantId: claims?.tenantId || tenantId,
        authAuthority: tenantId,
      });
      const refreshToken = await legacyManager.getRefreshToken().catch(() => null);
      const metadata = await legacyManager.getTokenMetadata();
      const expiresIn = metadata
        ? Math.max(60, Math.floor((metadata.accessTokenExpiry - Date.now()) / 1000))
        : 3600;
      await newManager.tokenManager.storeTokens(accessToken, refreshToken, expiresIn);

      await accountRegistry.addAccount({
        accountId,
        tenantId: claims?.tenantId || tenantId,
        clientId,
        authMode: this.startupConfig.authMode,
        userId: result.user.id,
        email: result.user.mail,
        displayName: result.user.displayName,
      });

      await legacyManager.clearTokens();
      console.error(`Migrated legacy tokens to account ${accountId}`);
    } catch (error) {
      console.error('Legacy token migration failed:', error.message);
    }
  }

  async listAccounts() {
    await this.initialize();
    return accountRegistry.listAccounts();
  }

  async hasAccounts() {
    await this.initialize();
    return accountRegistry.hasAccounts();
  }

  async resolve(accountId) {
    await this.initialize();
    const account = accountId
      ? await accountRegistry.getAccount(accountId)
      : await accountRegistry.getDefaultAccount();
    if (!account) {
      throw createAuthError('No account connected. Call outlook_connect_account to sign in.', true);
    }
    const manager = await this.getOrCreateManager(account);
    return { manager, account };
  }

  async resolveAll() {
    await this.initialize();
    const accounts = await accountRegistry.listAccounts();
    if (accounts.length === 0) {
      throw createAuthError('No account connected. Call outlook_connect_account to sign in.', true);
    }
    return accounts;
  }

  async getOrCreateManager(account) {
    if (managerCache.has(account.accountId)) {
      return managerCache.get(account.accountId);
    }
    const manager = new OutlookAuthManager({
      accountId: account.accountId,
      clientId: account.clientId,
      tenantId: account.tenantId,
      authAuthority: account.tenantId,
    });
    managerCache.set(account.accountId, manager);
    return manager;
  }

  async connectAccount(opts = {}) {
    await this.initialize();
    const clientId = resolveClientId(opts.clientId);
    if (!isClientIdConfigured(clientId)) {
      return { success: false, error: createAuthError(NO_APP_CONFIGURED_MESSAGE, true) };
    }
    const authority = resolveAuthAuthority(opts.tenantId);
    const authMode = resolveAuthMode(clientId, authority);

    const pendingAccountId = `pending-${Date.now()}`;
    const manager = new OutlookAuthManager({
      accountId: pendingAccountId,
      clientId,
      tenantId: authority,
      authAuthority: authority,
      isConnectFlow: true,
    });

    const result = await manager.authenticateConnect();
    if (!result.success) {
      return result;
    }

    if (result.pending) {
      result.authCompletion
        .then((authResult) => this.finalizeConnectedAccount({
          manager,
          authResult,
          clientId,
          authority,
          authMode,
        }))
        .catch((error) => {
          console.error('Background account connection failed:', error);
        });

      return {
        success: true,
        pending: true,
        message: result.message,
        authUrl: result.authUrl,
      };
    }

    return this.finalizeConnectedAccount({
      manager,
      authResult: result,
      clientId,
      authority,
      authMode,
    });
  }

  /**
   * Build a tenant-wide admin-consent URL for the configured (multi-tenant) app.
   * An administrator of the target tenant opens this URL once to grant the app
   * org-wide consent; afterwards users in that tenant can connect normally.
   * No per-tenant app registration is required.
   */
  requestAdminConsent(opts = {}) {
    const clientId = resolveClientId(opts.clientId);
    if (!isClientIdConfigured(clientId)) {
      throw createAuthError(NO_APP_CONFIGURED_MESSAGE, true);
    }

    const tenant = (opts.tenantId || '').trim();
    if (!tenant || tenant === ORGANIZATIONS_AUTHORITY) {
      throw createAuthError(
        'A specific tenant is required for admin consent. Pass tenantId as the tenant\'s '
        + 'domain (e.g. contoso.com) or directory (tenant) GUID.',
        true
      );
    }

    const url = new URL(authConfig.oauth.adminConsentUrl(tenant));
    url.searchParams.set('client_id', clientId);
    url.searchParams.set('state', crypto.randomBytes(16).toString('hex'));
    const redirectUri = process.env.MCP_OUTLOOK_ADMIN_CONSENT_REDIRECT;
    if (redirectUri) {
      url.searchParams.set('redirect_uri', redirectUri);
    }

    return {
      adminConsentUrl: url.toString(),
      tenantId: tenant,
      clientId,
      scopes: authConfig.oauth.scope,
    };
  }

  async finalizeConnectedAccount({ manager, authResult, clientId, authority, authMode }) {
    if (!authResult.success) {
      return authResult;
    }

    const accessToken = await manager.tokenManager.getAccessToken();
    const claims = extractAccountClaims(accessToken);
    const accountId = claims?.accountId || `${authority}:${authResult.user.id}`;

    const finalManager = new OutlookAuthManager({
      accountId,
      clientId,
      tenantId: claims?.tenantId || authority,
      authAuthority: claims?.tenantId || authority,
    });

    const refreshToken = await manager.tokenManager.getRefreshToken().catch(() => null);
    const metadata = await manager.tokenManager.getTokenMetadata();
    const expiresIn = metadata
      ? Math.max(60, Math.floor((metadata.accessTokenExpiry - Date.now()) / 1000))
      : 3600;

    await finalManager.tokenManager.storeTokens(accessToken, refreshToken, expiresIn);
    await manager.tokenManager.clearTokens();

    const account = await accountRegistry.addAccount({
      accountId,
      tenantId: claims?.tenantId || authority,
      clientId,
      authMode,
      userId: authResult.user.id,
      email: authResult.user.mail,
      displayName: authResult.user.displayName,
    });

    managerCache.set(accountId, finalManager);
    finalManager.authenticationRecord = account;
    finalManager.isAuthenticated = true;
    await finalManager.initializeGraphClient();

    return {
      success: true,
      account: {
        accountId,
        email: authResult.user.mail,
        displayName: authResult.user.displayName,
        tenantId: claims?.tenantId || authority,
        authMode,
        isDefault: (await accountRegistry.listAccounts()).find(a => a.accountId === accountId)?.isDefault,
      },
      deviceCodeInfo: authResult.deviceCodeInfo || null,
    };
  }

  async removeAccount(accountId) {
    await this.initialize();
    const manager = managerCache.get(accountId);
    if (manager) {
      await manager.logout();
      managerCache.delete(accountId);
    } else {
      const account = await accountRegistry.getAccount(accountId);
      if (account) {
        const m = new OutlookAuthManager({
          accountId,
          clientId: account.clientId,
          tenantId: account.tenantId,
          authAuthority: account.tenantId,
        });
        await m.logout();
      }
    }
    clearStylingCache(accountId);
    clearSignatureCache(accountId);
    await accountRegistry.removeAccount(accountId);
    return { success: true, accountId };
  }

  async setDefaultAccount(accountId) {
    await this.initialize();
    const account = await accountRegistry.setDefaultAccount(accountId);
    return { success: true, account };
  }

  registerPendingAuth(state, session) {
    pendingAuthSessions.set(state, session);
  }

  consumePendingAuth(state) {
    const session = pendingAuthSessions.get(state);
    pendingAuthSessions.delete(state);
    return session;
  }

  evictManager(accountId) {
    managerCache.delete(accountId);
  }
}

export const authManagerRegistry = new AuthManagerRegistry();
