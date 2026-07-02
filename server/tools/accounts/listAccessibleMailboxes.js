import { createSafeResponse } from '../../utils/jsonUtils.js';
import { convertErrorToToolError } from '../../utils/mcpErrorResponse.js';
import { buildMailboxBase } from '../../graph/mailboxPath.js';

export async function listAccessibleMailboxesTool(registry, args = {}) {
  try {
    const { manager, account } = await registry.resolve(args.accountId);
    await manager.ensureAuthenticated();
    const graphApiClient = manager.getGraphApiClient();
    const base = buildMailboxBase(args.mailbox);

    const result = await graphApiClient.makeRequest(`${base}/mailFolders`, {
      select: 'id,displayName,parentFolderId',
      top: 100,
    });

    if (result.content && result.isError !== undefined) return result;

    return createSafeResponse({
      accountId: account.accountId,
      accountEmail: account.email,
      mailFolders: result.value || [],
      note: 'Shared mailboxes may also be accessed via the mailbox parameter using /users/{email} paths.',
    });
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to list accessible mailboxes');
  }
}
