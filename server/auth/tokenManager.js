let keytar;
let keytarImportPromise;

const getKeytar = async () => {
  if (keytar === undefined) {
    if (!keytarImportPromise) {
      keytarImportPromise = (async () => {
        try {
          if (process.env.MCP_OUTLOOK_HEADLESS === 'true') return null;
          const keytarModule = await import('keytar');
          return keytarModule.default;
        } catch (error) {
          console.error('Keytar not available - using fallback token storage:', error.message);
          return null;
        }
      })();
    }
    keytar = await keytarImportPromise;
  }
  return keytar;
};

import storage from 'node-persist';
import crypto from 'crypto';
import path from 'path';
import fs from 'fs';
import { createAuthError, convertErrorToToolError } from '../utils/mcpErrorResponse.js';
import { getAccountDir, getLegacyTokensDir, ensureDir } from './dataPaths.js';

const SERVICE_NAME = 'outlook-mcp';
const ENCRYPTION_KEY_ACCOUNT = 'encryption-key';
const ACCESS_TOKEN_ACCOUNT = 'access-token';
const REFRESH_TOKEN_ACCOUNT = 'refresh-token';
const TOKEN_METADATA_KEY = 'token-metadata';

const storageInstances = new Map();

export function resetTokenStorageCache() {
  storageInstances.clear();
}

async function getStorage(accountId) {
  const key = accountId || '__legacy__';
  if (storageInstances.has(key)) return storageInstances.get(key);

  const dir = accountId
    ? getAccountDir(accountId)
    : getLegacyTokensDir();

  ensureDir(dir);

  const instance = storage.create({
    dir,
    logging: false,
  });
  await instance.init();
  storageInstances.set(key, instance);
  return instance;
}

export class TokenManager {
  constructor(clientId, accountId = null) {
    this.clientId = clientId;
    this.accountId = accountId;
    this.storageInitialized = false;
    this.encryptionKey = null;
    this.storage = null;
  }

  getKeytarServiceName() {
    return this.accountId ? `${SERVICE_NAME}/${this.accountId}` : SERVICE_NAME;
  }

  async initialize() {
    if (this.storageInitialized) return;
    this.storage = await getStorage(this.accountId);
    this.encryptionKey = await this.getOrCreateEncryptionKey();
    this.storageInitialized = true;
  }

  async getOrCreateEncryptionKey() {
    const serviceName = this.getKeytarServiceName();
    const keyAccount = this.accountId
      ? `${ENCRYPTION_KEY_ACCOUNT}-${this.accountId}`
      : ENCRYPTION_KEY_ACCOUNT;

    try {
      const keytarInstance = await getKeytar();
      if (keytarInstance) {
        const existingKey = await keytarInstance.getPassword(serviceName, keyAccount);
        if (existingKey) {
          return Buffer.from(existingKey, 'base64');
        }
        const newKey = crypto.randomBytes(32);
        await keytarInstance.setPassword(serviceName, keyAccount, newKey.toString('base64'));
        return newKey;
      }
    } catch {
      // Fall through to file-based key
    }

    const keyPath = path.join(
      this.accountId ? getAccountDir(this.accountId) : getLegacyTokensDir(),
      '.encryption-key'
    );
    try {
      if (fs.existsSync(keyPath)) {
        return Buffer.from(await fs.promises.readFile(keyPath));
      }
    } catch {
      // continue
    }

    const fallbackKey = crypto.createHash('sha256')
      .update(`${this.clientId}:${this.accountId || 'legacy'}`)
      .digest();
    try {
      await fs.promises.writeFile(keyPath, fallbackKey);
    } catch {
      // non-fatal
    }
    return fallbackKey;
  }

  encrypt(text) {
    const iv = crypto.randomBytes(16);
    const cipher = crypto.createCipheriv('aes-256-cbc', this.encryptionKey, iv);
    let encrypted = cipher.update(text, 'utf8', 'hex');
    encrypted += cipher.final('hex');
    return iv.toString('hex') + ':' + encrypted;
  }

  decrypt(encryptedText) {
    const parts = encryptedText.split(':');
    const iv = Buffer.from(parts.shift(), 'hex');
    const encrypted = parts.join(':');
    const decipher = crypto.createDecipheriv('aes-256-cbc', this.encryptionKey, iv);
    let decrypted = decipher.update(encrypted, 'hex', 'utf8');
    decrypted += decipher.final('utf8');
    return decrypted;
  }

  async storeTokens(accessToken, refreshToken, expiresIn = 3600) {
    await this.initialize();
    const serviceName = this.getKeytarServiceName();

    let usingFallback = false;
    try {
      const keytarInstance = await getKeytar();
      if (keytarInstance) {
        await keytarInstance.setPassword(serviceName, ACCESS_TOKEN_ACCOUNT, this.encrypt(accessToken));
        if (refreshToken) {
          await keytarInstance.setPassword(serviceName, REFRESH_TOKEN_ACCOUNT, this.encrypt(refreshToken));
        }
      } else {
        usingFallback = true;
        await this.storage.setItem('fallback_access_token', this.encrypt(accessToken));
        if (refreshToken) {
          await this.storage.setItem('fallback_refresh_token', this.encrypt(refreshToken));
        }
      }
    } catch {
      usingFallback = true;
      await this.storage.setItem('fallback_access_token', this.encrypt(accessToken));
      if (refreshToken) {
        await this.storage.setItem('fallback_refresh_token', this.encrypt(refreshToken));
      }
    }

    const metadata = {
      accessTokenExpiry: Date.now() + (expiresIn * 1000),
      refreshTokenExpiry: Date.now() + (90 * 24 * 60 * 60 * 1000),
      lastRefresh: Date.now(),
    };
    await this.storage.setItem(TOKEN_METADATA_KEY, metadata);

    if (usingFallback) {
      console.error(`Tokens stored securely for account ${this.accountId || 'legacy'}`);
    }
  }

  async getAccessToken() {
    try {
      await this.initialize();
      const metadata = await this.storage.getItem(TOKEN_METADATA_KEY);
      if (!metadata) {
        throw createAuthError('No token metadata found', true);
      }

      const refreshThreshold = 55 * 60 * 1000;
      const shouldRefresh = Date.now() > (metadata.accessTokenExpiry - refreshThreshold);
      if (shouldRefresh) {
        const error = createAuthError('Access token needs refresh', true);
        error._errorDetails = { ...error._errorDetails, needsRefresh: true };
        throw error;
      }

      const serviceName = this.getKeytarServiceName();
      const keytarInstance = await getKeytar();
      if (keytarInstance) {
        try {
          const encryptedToken = await keytarInstance.getPassword(serviceName, ACCESS_TOKEN_ACCOUNT);
          if (encryptedToken) return this.decrypt(encryptedToken);
        } catch {
          // fall through
        }
      }

      const fallbackToken = await this.storage.getItem('fallback_access_token');
      if (fallbackToken) return this.decrypt(fallbackToken);
      throw createAuthError('No access token found', true);
    } catch (error) {
      if (error.isError) throw error;
      throw convertErrorToToolError(error, 'Failed to retrieve access token');
    }
  }

  async getRefreshToken() {
    try {
      await this.initialize();
      const metadata = await this.storage.getItem(TOKEN_METADATA_KEY);
      if (!metadata) throw createAuthError('No token metadata found', true);
      if (Date.now() > metadata.refreshTokenExpiry) {
        throw createAuthError('Refresh token has expired', true);
      }

      const serviceName = this.getKeytarServiceName();
      const keytarInstance = await getKeytar();
      if (keytarInstance) {
        try {
          const encryptedToken = await keytarInstance.getPassword(serviceName, REFRESH_TOKEN_ACCOUNT);
          if (encryptedToken) return this.decrypt(encryptedToken);
        } catch {
          // fall through
        }
      }

      const fallbackToken = await this.storage.getItem('fallback_refresh_token');
      if (fallbackToken) return this.decrypt(fallbackToken);
      throw createAuthError('No refresh token found', true);
    } catch (error) {
      if (error.isError) throw error;
      throw convertErrorToToolError(error, 'Failed to retrieve refresh token');
    }
  }

  async clearTokens() {
    await this.initialize();
    const serviceName = this.getKeytarServiceName();
    const keytarInstance = await getKeytar();
    if (keytarInstance) {
      try {
        await keytarInstance.deletePassword(serviceName, ACCESS_TOKEN_ACCOUNT);
        await keytarInstance.deletePassword(serviceName, REFRESH_TOKEN_ACCOUNT);
      } catch {
        // continue
      }
    }
    await this.storage.removeItem('fallback_access_token');
    await this.storage.removeItem('fallback_refresh_token');
    await this.storage.removeItem(TOKEN_METADATA_KEY);
  }

  generateCodeVerifier() {
    return crypto.randomBytes(32).toString('base64url');
  }

  generateCodeChallenge(verifier) {
    return crypto.createHash('sha256').update(verifier).digest('base64url');
  }

  async storePKCEVerifier(verifier, sessionId = 'default') {
    await this.initialize();
    await this.storage.setItem(`pkce_verifier_${sessionId}`, verifier);
  }

  async getPKCEVerifier(sessionId = 'default') {
    try {
      await this.initialize();
      const key = `pkce_verifier_${sessionId}`;
      const verifier = await this.storage.getItem(key);
      await this.storage.removeItem(key);
      if (!verifier) throw createAuthError('PKCE verifier not found or expired', true);
      return verifier;
    } catch (error) {
      if (error.isError) throw error;
      throw convertErrorToToolError(error, 'Failed to retrieve PKCE verifier');
    }
  }

  async isAuthenticated() {
    try {
      await this.getAccessToken();
      return true;
    } catch {
      return false;
    }
  }

  async getTokenMetadata() {
    await this.initialize();
    return await this.storage.getItem(TOKEN_METADATA_KEY);
  }
}

export class LegacyTokenManager extends TokenManager {
  constructor(clientId) {
    super(clientId, null);
  }

  async hasLegacyTokens() {
    await this.initialize();
    const metadata = await this.storage.getItem(TOKEN_METADATA_KEY);
    return !!metadata;
  }
}
