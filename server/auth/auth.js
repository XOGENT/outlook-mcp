import { Client } from '@microsoft/microsoft-graph-client';
import { TokenManager } from './tokenManager.js';
import { authConfig } from './config.js';
import { GraphApiClient } from '../graph/graphClient.js';
import { createAuthError, convertErrorToToolError } from '../utils/mcpErrorResponse.js';
import { authenticateWithDeviceCode, shouldUseDeviceCodeFlow } from './deviceCodeFlow.js';
import { extractAccountClaims } from './jwtUtils.js';
import http from 'http';
import url from 'url';
import crypto from 'crypto';
import { exec } from 'child_process';

function successHtml(title, message) {
  return `<html><head><title>${title}</title></head><body style="font-family:Segoe UI,Arial,sans-serif;text-align:center;padding:50px"><h1>${title}</h1><p>${message}</p><p>You can close this window.</p></body></html>`;
}

export class OutlookAuthManager {
  constructor({ accountId, clientId, tenantId, authAuthority, tokenManager, isConnectFlow = false }) {
    this.accountId = accountId;
    this.clientId = clientId;
    this.tenantId = tenantId;
    this.authAuthority = authAuthority || tenantId;
    this.isConnectFlow = isConnectFlow;
    this.tokenManager = tokenManager || new TokenManager(clientId, accountId);
    this.graphClient = null;
    this.graphApiClient = null;
    this.isAuthenticated = false;
    this.authenticationRecord = null;
    this.lastUsedPort = null;
    this.sessionId = crypto.randomBytes(8).toString('hex');
  }

  getCurrentUser() {
    return this.authenticationRecord;
  }

  openBrowser(targetUrl) {
    const platform = process.platform;
    let command;
    switch (platform) {
      case 'darwin':
        command = `open "${targetUrl}"`;
        break;
      case 'win32':
        command = `start "" "${targetUrl}"`;
        break;
      default:
        command = `xdg-open "${targetUrl}"`;
        break;
    }
    exec(command, () => {});
  }

  async authenticateConnect() {
    try {
      if (shouldUseDeviceCodeFlow()) {
        return await this.authenticateDeviceCode();
      }
      return await this.authenticateInteractiveNonBlocking();
    } catch (error) {
      console.error('Connect authentication error:', error);
      this.isAuthenticated = false;
      if (error.isError) {
        return { success: false, error };
      }
      return { success: false, error: createAuthError(error.message, true) };
    }
  }

  async authenticateDeviceCode() {
    const { tokenResponse, deviceCodeInfo } = await authenticateWithDeviceCode(
      this.clientId,
      this.authAuthority
    );
    await this.tokenManager.storeTokens(
      tokenResponse.access_token,
      tokenResponse.refresh_token,
      tokenResponse.expires_in
    );
    await this.initializeGraphClient();
    const validation = await this.validateAuthentication();
    return { ...validation, deviceCodeInfo };
  }

  async authenticate() {
    try {
      const isTokenValid = await this.tokenManager.isAuthenticated();
      if (isTokenValid) {
        await this.initializeGraphClient();
        return await this.validateAuthentication();
      }
      if (this.isConnectFlow) {
        return await this.authenticateConnect();
      }
      return {
        success: false,
        error: createAuthError('No account connected. Call outlook_connect_account to sign in.', true),
      };
    } catch (error) {
      console.error('Authentication error:', error);
      this.isAuthenticated = false;
      if (error.isError) {
        return { success: false, error };
      }
      return { success: false, error: createAuthError(error.message, true) };
    }
  }

  async authenticateInteractive() {
    const codeVerifier = this.tokenManager.generateCodeVerifier();
    const codeChallenge = this.tokenManager.generateCodeChallenge(codeVerifier);
    await this.tokenManager.storePKCEVerifier(codeVerifier, this.sessionId);

    const authorizationCode = await this.getAuthorizationCode(codeChallenge);
    if (!authorizationCode) {
      return { success: false, error: createAuthError('Failed to get authorization code', true) };
    }

    return await this.completeInteractiveAuth(authorizationCode);
  }

  async authenticateInteractiveNonBlocking() {
    const codeVerifier = this.tokenManager.generateCodeVerifier();
    const codeChallenge = this.tokenManager.generateCodeChallenge(codeVerifier);
    await this.tokenManager.storePKCEVerifier(codeVerifier, this.sessionId);

    let listenInfo = null;
    const codePromise = this.getAuthorizationCode(codeChallenge, {
      onListening: (info) => {
        listenInfo = info;
      },
    });

    const deadline = Date.now() + 5000;
    while (!listenInfo && Date.now() < deadline) {
      await new Promise((resolve) => setTimeout(resolve, 25));
    }

    if (!listenInfo) {
      return { success: false, error: createAuthError('Failed to start authentication server', true) };
    }

    const authCompletion = codePromise
      .then((authorizationCode) => this.completeInteractiveAuth(authorizationCode))
      .catch((error) => {
        console.error('Background interactive authentication failed:', error);
        if (error?.isError) {
          return { success: false, error };
        }
        return { success: false, error: createAuthError(error.message, true) };
      });

    return {
      success: true,
      pending: true,
      message: 'Browser opened for Microsoft sign-in. Complete sign-in, then call outlook_list_accounts.',
      authUrl: listenInfo.url,
      authCompletion,
    };
  }

  async completeInteractiveAuth(authorizationCode) {
    const tokenResponse = await this.exchangeCodeForToken(authorizationCode);
    await this.tokenManager.storeTokens(
      tokenResponse.access_token,
      tokenResponse.refresh_token,
      tokenResponse.expires_in
    );

    const claims = extractAccountClaims(tokenResponse.access_token);
    if (claims?.tenantId && this.authAuthority === 'organizations') {
      this.tenantId = claims.tenantId;
    }

    await this.initializeGraphClient();
    return await this.validateAuthentication();
  }

  async getAuthorizationCode(codeChallenge, { onListening } = {}) {
    return new Promise((resolve, reject) => {
      const state = crypto.randomBytes(16).toString('hex');
      const authUrl = new URL(authConfig.oauth.authorizeUrl(this.authAuthority));

      authUrl.searchParams.append('client_id', this.clientId);
      authUrl.searchParams.append('response_type', 'code');
      authUrl.searchParams.append('scope', authConfig.oauth.scope);
      authUrl.searchParams.append('state', state);
      authUrl.searchParams.append('code_challenge', codeChallenge);
      authUrl.searchParams.append('code_challenge_method', 'S256');
      authUrl.searchParams.append('prompt', 'select_account');

      const server = http.createServer(async (req, res) => {
        const parsedUrl = url.parse(req.url, true);
        if (parsedUrl.pathname !== '/callback') return;

        const code = parsedUrl.query.code;
        const returnedState = parsedUrl.query.state;

        if (returnedState !== state) {
          res.writeHead(400, { 'Content-Type': 'text/html' });
          res.end(successHtml('Security Error', 'State mismatch. Please try again.'));
          server.close();
          reject(createAuthError('State mismatch - possible CSRF attack', false));
          return;
        }

        if (code) {
          res.writeHead(200, { 'Content-Type': 'text/html' });
          res.end(successHtml('Authentication Successful', 'Outlook MCP account connected.'));
          server.close();
          resolve(code);
        } else {
          res.writeHead(400, { 'Content-Type': 'text/html' });
          res.end(successHtml('Authentication Failed', 'No authorization code received.'));
          server.close();
          reject(createAuthError('No authorization code received', true));
        }
      });

      server.listen(0, () => {
        const port = server.address().port;
        this.lastUsedPort = port;
        authUrl.searchParams.set('redirect_uri', `http://localhost:${port}/callback`);
        console.error(`Listening for authentication callback on port ${port}...`);
        console.error(`Opening browser for Microsoft account selection...`);
        console.error(authUrl.toString());
        this.openBrowser(authUrl.toString());
        onListening?.({ url: authUrl.toString(), port });
      });

      setTimeout(() => {
        server.close();
        reject(createAuthError('Authentication timeout - please try again', true));
      }, 5 * 60 * 1000);
    });
  }

  async exchangeCodeForToken(code) {
    const codeVerifier = await this.tokenManager.getPKCEVerifier(this.sessionId);
    const tokenUrl = authConfig.oauth.tokenUrl(this.authAuthority);
    const params = new URLSearchParams({
      client_id: this.clientId,
      scope: authConfig.oauth.scope,
      code,
      redirect_uri: `http://localhost:${this.lastUsedPort}/callback`,
      grant_type: 'authorization_code',
      code_verifier: codeVerifier,
    });

    const response = await fetch(tokenUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: params.toString(),
    });

    if (!response.ok) {
      const error = await response.text();
      throw createAuthError(`Token exchange failed: ${error}`, true);
    }
    return await response.json();
  }

  getTokenTenantId() {
    return this.tenantId === 'organizations' ? this.authAuthority : this.tenantId;
  }

  async refreshAccessToken() {
    try {
      const refreshToken = await this.tokenManager.getRefreshToken();
      const tokenTenant = this.getTokenTenantId();
      const tokenUrl = authConfig.oauth.tokenUrl(tokenTenant);

      const params = new URLSearchParams({
        client_id: this.clientId,
        scope: authConfig.oauth.scope,
        refresh_token: refreshToken,
        grant_type: 'refresh_token',
      });

      const response = await fetch(tokenUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: params.toString(),
      });

      if (!response.ok) {
        const error = await response.text();
        throw createAuthError(`Token refresh failed: ${error}`, true);
      }

      const tokenResponse = await response.json();
      const claims = extractAccountClaims(tokenResponse.access_token);
      if (claims?.tenantId) {
        this.tenantId = claims.tenantId;
      }

      await this.tokenManager.storeTokens(
        tokenResponse.access_token,
        tokenResponse.refresh_token || refreshToken,
        tokenResponse.expires_in
      );

      await this.initializeGraphClient();
      return true;
    } catch (error) {
      console.error('Token refresh failed:', error);
      await this.tokenManager.clearTokens();
      if (error.isError) throw error;
      throw convertErrorToToolError(error, 'Token refresh failed');
    }
  }

  async initializeGraphClient() {
    const authProvider = {
      getAccessToken: async () => {
        try {
          return await this.tokenManager.getAccessToken();
        } catch (error) {
          if (error.message?.includes('needs refresh') || error._errorDetails?.needsRefresh) {
            await this.refreshAccessToken();
            return await this.tokenManager.getAccessToken();
          }
          throw error;
        }
      },
    };

    this.graphClient = Client.init({
      authProvider: (done) => {
        authProvider.getAccessToken()
          .then(token => done(null, token))
          .catch(error => done(error, null));
      },
      defaultVersion: 'v1.0',
    });

    this.graphApiClient = new GraphApiClient(this);
    await this.graphApiClient.initialize();
  }

  async validateAuthentication() {
    try {
      const user = await this.graphClient.api('/me').get();
      this.isAuthenticated = true;
      this.authenticationRecord = {
        id: user.id,
        displayName: user.displayName,
        mail: user.mail || user.userPrincipalName,
        accountId: this.accountId,
      };
      return {
        success: true,
        user: this.authenticationRecord,
      };
    } catch (error) {
      this.isAuthenticated = false;
      if (error.isError) throw error;
      throw convertErrorToToolError(error, 'User validation failed');
    }
  }

  async ensureAuthenticated() {
    if (!this.isAuthenticated || !this.graphClient) {
      const result = await this.authenticate();
      if (!result.success) {
        if (result.error?.isError) throw result.error;
        throw createAuthError(`Authentication failed: ${result.error}`, true);
      }
    }

    try {
      await this.tokenManager.getAccessToken();
    } catch (error) {
      if (error.isError) {
        if (error._errorDetails?.needsRefresh) {
          await this.refreshAccessToken();
        } else {
          throw error;
        }
      } else if (error.message?.includes('needs refresh')) {
        await this.refreshAccessToken();
      } else {
        throw convertErrorToToolError(error, 'Token validation failed');
      }
    }

    return this.graphClient;
  }

  getGraphClient() {
    if (!this.graphClient) {
      throw createAuthError('Not authenticated. Call outlook_connect_account first.', true);
    }
    return this.graphClient;
  }

  getGraphApiClient() {
    if (!this.graphApiClient) {
      throw createAuthError('Not authenticated. Call outlook_connect_account first.', true);
    }
    return this.graphApiClient;
  }

  async logout() {
    await this.tokenManager.clearTokens();
    this.graphClient = null;
    this.graphApiClient = null;
    this.isAuthenticated = false;
    this.authenticationRecord = null;
  }
}
