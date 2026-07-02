import { convertErrorToToolError, createValidationError } from '../../utils/mcpErrorResponse.js';
import { createSafeResponse } from '../../utils/jsonUtils.js';
import { resolveReadAccount } from '../common/crossAccountFanOut.js';
import { buildMailboxBase } from '../../graph/mailboxPath.js';

// List mail folders
export async function listFoldersTool(registry, args) {
  const { includeHidden = false, includeChildFolders = true, top = 100 } = args;

  try {
    const { manager } = await resolveReadAccount(registry, args);
    await manager.ensureAuthenticated();
    const graphApiClient = manager.getGraphApiClient();
    const mailboxBase = buildMailboxBase(args.mailbox);

    const options = {
      select: 'id,displayName,parentFolderId,childFolderCount,unreadItemCount,totalItemCount,isHidden',
      top: Math.min(top, 1000)
    };

    if (!includeHidden) {
      options.filter = 'isHidden eq false';
    }

    let endpoint = `${mailboxBase}/mailFolders`;
    if (includeChildFolders) {
      endpoint = `${mailboxBase}/mailFolders?includeNestedFolders=true`;
    }

    const result = await graphApiClient.makeRequest(endpoint, options);

    const folders = result.value?.map(folder => ({
      id: folder.id,
      name: folder.displayName,
      parentFolderId: folder.parentFolderId,
      childFolderCount: folder.childFolderCount || 0,
      unreadItemCount: folder.unreadItemCount || 0,
      totalItemCount: folder.totalItemCount || 0,
      isHidden: folder.isHidden || false
    })) || [];

    return createSafeResponse({ folders, count: folders.length });
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to list folders');
  }
}