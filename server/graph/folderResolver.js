// Folder resolution utilities for Microsoft Graph API
import { buildMailboxBase } from './mailboxPath.js';

export class FolderResolver {
  constructor(graphApiClient) {
    this.graphApiClient = graphApiClient;
    this.cacheByMailbox = new Map();
    this.cacheExpiry = 5 * 60 * 1000;
  }

  getCache(mailboxBase = '/me') {
    if (!this.cacheByMailbox.has(mailboxBase)) {
      this.cacheByMailbox.set(mailboxBase, {
        foldersByName: new Map(),
        foldersById: new Map(),
        foldersList: [],
        lastCacheUpdate: 0,
      });
    }
    return this.cacheByMailbox.get(mailboxBase);
  }

  async refreshFolderCache(mailboxBase = '/me') {
    const cache = this.getCache(mailboxBase);
    const result = await this.graphApiClient.makeRequest(`${mailboxBase}/mailFolders`, {
      select: 'id,displayName,parentFolderId',
      top: 1000,
    });

    cache.foldersByName.clear();
    cache.foldersById.clear();
    cache.foldersList = [];

    if (result.value) {
      result.value.forEach(folder => {
        const folderInfo = {
          id: folder.id,
          displayName: folder.displayName,
          parentFolderId: folder.parentFolderId,
        };
        cache.foldersByName.set(folder.displayName.toLowerCase(), folderInfo);
        cache.foldersById.set(folder.id.toLowerCase(), folderInfo);
        cache.foldersList.push(folderInfo);
      });
    }

    cache.lastCacheUpdate = Date.now();
    console.error(`Folder cache refreshed with ${cache.foldersList.length} folders for ${mailboxBase}`);
  }

  shouldRefreshCache(mailboxBase = '/me') {
    const cache = this.getCache(mailboxBase);
    return (Date.now() - cache.lastCacheUpdate) > this.cacheExpiry || cache.foldersList.length === 0;
  }

  async resolveFolderToId(folderNameOrId, mailbox) {
    const mailboxBase = buildMailboxBase(mailbox);
    if (!folderNameOrId) throw new Error('Folder name or ID is required');
    if (folderNameOrId.toLowerCase() === 'inbox') return 'inbox';

    const folderIdRegex = /^[A-Za-z0-9+/]+=*$/;
    if (folderIdRegex.test(folderNameOrId) && folderNameOrId.length > 20) {
      return folderNameOrId;
    }

    const cache = this.getCache(mailboxBase);
    if (this.shouldRefreshCache(mailboxBase)) {
      await this.refreshFolderCache(mailboxBase);
    }

    const folderInfo = cache.foldersByName.get(folderNameOrId.toLowerCase());
    if (folderInfo) return folderInfo.id;

    await this.refreshFolderCache(mailboxBase);
    const refreshed = cache.foldersByName.get(folderNameOrId.toLowerCase());
    if (refreshed) return refreshed.id;

    throw new Error(`Folder '${folderNameOrId}' not found. Available folders: ${cache.foldersList.map(f => f.displayName).join(', ')}`);
  }

  async resolveFoldersToIds(folderNamesOrIds, mailbox) {
    if (!Array.isArray(folderNamesOrIds) || folderNamesOrIds.length === 0) return [];
    const resolvedIds = [];
    for (const folderNameOrId of folderNamesOrIds) {
      resolvedIds.push(await this.resolveFolderToId(folderNameOrId, mailbox));
    }
    return resolvedIds;
  }

  async getFolderInfo(folderNameOrId, mailbox) {
    const mailboxBase = buildMailboxBase(mailbox);
    const folderId = await this.resolveFolderToId(folderNameOrId, mailbox);
    const cache = this.getCache(mailboxBase);
    const cachedInfo = cache.foldersById.get(folderId.toLowerCase());
    if (cachedInfo) return cachedInfo;

    const folderData = await this.graphApiClient.makeRequest(`${mailboxBase}/mailFolders/${folderId}`, {
      select: 'id,displayName,parentFolderId,totalItemCount,unreadItemCount',
    });

    return {
      id: folderData.id,
      displayName: folderData.displayName,
      parentFolderId: folderData.parentFolderId,
      totalItemCount: folderData.totalItemCount,
      unreadItemCount: folderData.unreadItemCount,
    };
  }

  async listAllFolders(mailbox) {
    const mailboxBase = buildMailboxBase(mailbox);
    if (this.shouldRefreshCache(mailboxBase)) {
      await this.refreshFolderCache(mailboxBase);
    }
    return [...this.getCache(mailboxBase).foldersList].sort((a, b) => a.displayName.localeCompare(b.displayName));
  }

  clearCache(mailbox) {
    if (mailbox) {
      this.cacheByMailbox.delete(buildMailboxBase(mailbox));
    } else {
      this.cacheByMailbox.clear();
    }
  }
}
