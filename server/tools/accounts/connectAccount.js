import { createSafeResponse } from '../../utils/jsonUtils.js';
import { convertErrorToToolError } from '../../utils/mcpErrorResponse.js';

export async function connectAccountTool(registry, args = {}) {
  try {
    const result = await registry.connectAccount(args);
    if (!result.success) {
      if (result.error?.isError) return result.error;
      return convertErrorToToolError(new Error(result.error?.message || 'Connect failed'), 'Failed to connect account');
    }
    return createSafeResponse({
      success: true,
      account: result.account,
      deviceCodeInfo: result.deviceCodeInfo,
      message: `Connected account ${result.account.email} (${result.account.accountId})`,
    });
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to connect account');
  }
}
