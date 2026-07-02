import { createSafeResponse } from '../../utils/jsonUtils.js';
import { convertErrorToToolError } from '../../utils/mcpErrorResponse.js';

export async function setDefaultAccountTool(registry, args) {
  try {
    const result = await registry.setDefaultAccount(args.accountId);
    return createSafeResponse({
      success: true,
      account: result.account,
      message: `Default account set to ${args.accountId}`,
    });
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to set default account');
  }
}
