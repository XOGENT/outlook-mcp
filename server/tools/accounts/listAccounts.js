import { createSafeResponse } from '../../utils/jsonUtils.js';
import { convertErrorToToolError } from '../../utils/mcpErrorResponse.js';

export async function listAccountsTool(registry) {
  try {
    const accounts = await registry.listAccounts();
    return createSafeResponse({ accounts, count: accounts.length });
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to list accounts');
  }
}
