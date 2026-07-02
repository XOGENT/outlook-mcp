import { createAuthError, createValidationError } from '../../utils/mcpErrorResponse.js';

const MAX_PARALLEL_ACCOUNTS = 3;

async function runWithConcurrency(items, fn, limit) {
  const results = [];
  let index = 0;

  async function worker() {
    while (index < items.length) {
      const current = index++;
      results[current] = await fn(items[current], current);
    }
  }

  const workers = Array.from({ length: Math.min(limit, items.length) }, () => worker());
  await Promise.all(workers);
  return results;
}

function getNestedValue(obj, path) {
  return path.split('.').reduce((acc, key) => acc?.[key], obj);
}

export async function fanOutAcrossAccounts(registry, args, perAccountFn, options = {}) {
  const {
    sortKey,
    sortDesc = false,
    globalLimit = args.limit ?? 25,
    resultKey = 'items',
  } = options;

  const accounts = args.account
    ? [await registry.resolve(args.account).then(r => r.account)]
    : await registry.resolveAll();

  if (accounts.length === 0) {
    throw createAuthError('No account connected. Call outlook_connect_account to sign in.', true);
  }

  const perAccountLimit = Math.max(1, Math.ceil(globalLimit / accounts.length));

  const settled = await runWithConcurrency(
    accounts,
    async (account) => {
      try {
        const { manager } = await registry.resolve(account.accountId);
        const data = await perAccountFn(manager, account, { ...args, limit: perAccountLimit });
        return { account, data, error: null };
      } catch (error) {
        return {
          account,
          data: null,
          error: error.message || String(error),
        };
      }
    },
    MAX_PARALLEL_ACCOUNTS
  );

  const merged = [];
  const partialFailures = [];
  const accountsSearched = [];

  for (const result of settled) {
    accountsSearched.push(result.account.accountId);
    if (result.error) {
      partialFailures.push({ accountId: result.account.accountId, error: result.error });
      continue;
    }
    const items = result.data?.[resultKey] ?? result.data ?? [];
    for (const item of items) {
      merged.push({
        ...item,
        accountId: result.account.accountId,
        accountEmail: result.account.email,
      });
    }
  }

  if (sortKey) {
    merged.sort((a, b) => {
      const aVal = getNestedValue(a, sortKey) || '';
      const bVal = getNestedValue(b, sortKey) || '';
      const cmp = String(aVal).localeCompare(String(bVal));
      return sortDesc ? -cmp : cmp;
    });
  }

  const limited = merged.slice(0, globalLimit);

  return {
    [resultKey]: limited,
    fanOutSummary: {
      accountsSearched,
      accountsFailed: partialFailures,
      totalResults: limited.length,
      fanOutMode: args.account ? 'single-account' : 'all-accounts',
    },
  };
}

export async function resolveWriteAccount(registry, args) {
  const accounts = await registry.listAccounts();
  if (accounts.length === 0) {
    throw createAuthError('No account connected. Call outlook_connect_account to sign in.', true);
  }
  if (accounts.length > 1 && !args.account) {
    return createValidationError('account', 'Multiple accounts connected. Specify account to send from.');
  }
  const { manager, account } = await registry.resolve(args.account);
  return { manager, account };
}

export async function resolveReadAccount(registry, args) {
  const { manager, account } = await registry.resolve(args.account);
  return { manager, account, mailboxBase: args.mailbox };
}
