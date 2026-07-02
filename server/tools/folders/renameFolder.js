import { convertErrorToToolError, createValidationError } from '../../utils/mcpErrorResponse.js';
import { resolveWriteAccount } from '../common/crossAccountFanOut.js';
import { buildMailboxBase } from '../../graph/mailboxPath.js';

// Rename mail folder
export async function renameFolderTool(registry, args) {
  const { folderId, newDisplayName } = args;

  if (!folderId) {
    return createValidationError('folderId', 'Parameter is required');
  }

  if (!newDisplayName) {
    return createValidationError('newDisplayName', 'Parameter is required');
  }

  try {
    const resolvedWrite = await resolveWriteAccount(registry, args);
    if (resolvedWrite?.isError) return resolvedWrite;
    const { manager } = resolvedWrite;
    await manager.ensureAuthenticated();
    const graphApiClient = manager.getGraphApiClient();
    const mailboxBase = buildMailboxBase(args.mailbox);

    await graphApiClient.makeRequest(`${mailboxBase}/mailFolders/${folderId}`, {
      body: { displayName: newDisplayName }
    }, 'PATCH');

    return {
      content: [
        {
          type: 'text',
          text: `Folder renamed to "${newDisplayName}" successfully. Folder ID: ${folderId}`,
        },
      ],
    };
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to rename folder');
  }
}