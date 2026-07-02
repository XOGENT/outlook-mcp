import { convertErrorToToolError, createValidationError } from '../../utils/mcpErrorResponse.js';
import { resolveWriteAccount } from '../common/crossAccountFanOut.js';
import { buildMailboxBase } from '../../graph/mailboxPath.js';

// Create mail folder
export async function createFolderTool(registry, args) {
  const { displayName, parentFolderId } = args;

  if (!displayName) {
    return createValidationError('displayName', 'Parameter is required');
  }

  try {
    const resolvedWrite = await resolveWriteAccount(registry, args);
    if (resolvedWrite?.isError) return resolvedWrite;
    const { manager } = resolvedWrite;
    await manager.ensureAuthenticated();
    const graphApiClient = manager.getGraphApiClient();
    const mailboxBase = buildMailboxBase(args.mailbox);

    const folderData = {
      displayName: displayName
    };

    let endpoint = `${mailboxBase}/mailFolders`;
    if (parentFolderId) {
      endpoint = `${mailboxBase}/mailFolders/${parentFolderId}/childFolders`;
    }

    const result = await graphApiClient.postWithRetry(endpoint, folderData);

    return {
      content: [
        {
          type: 'text',
          text: `Folder "${displayName}" created successfully. Folder ID: ${result.id}`,
        },
      ],
    };
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to create folder');
  }
}