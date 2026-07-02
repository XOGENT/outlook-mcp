import { createSafeResponse } from '../../utils/jsonUtils.js';
import { convertErrorToToolError } from '../../utils/mcpErrorResponse.js';

export async function disconnectAccountTool(registry, args) {
  try {
    const result = await registry.removeAccount(args.accountId);
    return createSafeResponse({
      success: true,
      accountId: result.accountId,
      message: `Disconnected account ${args.accountId}`,
    });
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to disconnect account');
  }
}
