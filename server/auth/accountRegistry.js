import fs from 'fs/promises';
import { getRegistryPath, getAccountsDir, ensureDir } from './dataPaths.js';

let registryCache = null;

async function loadRegistry() {
  if (registryCache) return registryCache;
  ensureDir(getAccountsDir());
  const registryPath = getRegistryPath();
  try {
    const data = await fs.readFile(registryPath, 'utf8');
    registryCache = JSON.parse(data);
  } catch {
    registryCache = { accounts: [], defaultAccountId: null };
  }
  return registryCache;
}

async function saveRegistry(registry) {
  ensureDir(getAccountsDir());
  registryCache = registry;
  await fs.writeFile(getRegistryPath(), JSON.stringify(registry, null, 2), 'utf8');
}

export class AccountRegistry {
  async listAccounts() {
    const registry = await loadRegistry();
    return registry.accounts.map(a => ({
      accountId: a.accountId,
      tenantId: a.tenantId,
      clientId: a.clientId,
      authMode: a.authMode,
      userId: a.userId,
      email: a.email,
      displayName: a.displayName,
      addedAt: a.addedAt,
      isDefault: a.accountId === registry.defaultAccountId,
    }));
  }

  async getAccount(accountId) {
    const registry = await loadRegistry();
    const account = registry.accounts.find(a => a.accountId === accountId);
    if (!account) return null;
    return { ...account, isDefault: account.accountId === registry.defaultAccountId };
  }

  async hasAccounts() {
    const registry = await loadRegistry();
    return registry.accounts.length > 0;
  }

  async getDefaultAccount() {
    const registry = await loadRegistry();
    if (!registry.defaultAccountId) {
      return registry.accounts[0] || null;
    }
    return registry.accounts.find(a => a.accountId === registry.defaultAccountId) || registry.accounts[0] || null;
  }

  async addAccount(profile) {
    const registry = await loadRegistry();
    const existing = registry.accounts.findIndex(a => a.accountId === profile.accountId);
    const record = {
      accountId: profile.accountId,
      tenantId: profile.tenantId,
      clientId: profile.clientId,
      authMode: profile.authMode || 'hosted',
      userId: profile.userId,
      email: profile.email,
      displayName: profile.displayName,
      addedAt: profile.addedAt || new Date().toISOString(),
    };
    if (existing >= 0) {
      registry.accounts[existing] = { ...registry.accounts[existing], ...record };
    } else {
      registry.accounts.push(record);
    }
    if (!registry.defaultAccountId || registry.accounts.length === 1) {
      registry.defaultAccountId = profile.accountId;
    }
    await saveRegistry(registry);
    return record;
  }

  async removeAccount(accountId) {
    const registry = await loadRegistry();
    registry.accounts = registry.accounts.filter(a => a.accountId !== accountId);
    if (registry.defaultAccountId === accountId) {
      registry.defaultAccountId = registry.accounts[0]?.accountId || null;
    }
    await saveRegistry(registry);
  }

  async setDefaultAccount(accountId) {
    const registry = await loadRegistry();
    const account = registry.accounts.find(a => a.accountId === accountId);
    if (!account) {
      throw new Error(`Account not found: ${accountId}`);
    }
    registry.defaultAccountId = accountId;
    await saveRegistry(registry);
    return account;
  }

  invalidateCache() {
    registryCache = null;
  }
}

export const accountRegistry = new AccountRegistry();
