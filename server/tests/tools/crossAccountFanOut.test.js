import { describe, it, expect } from 'vitest';
import { fanOutAcrossAccounts } from '../../tools/common/crossAccountFanOut.js';

describe('crossAccountFanOut', () => {
  it('merges results from multiple accounts', async () => {
    const registry = {
      resolveAll: async () => [
        { accountId: 't1:u1', email: 'a@test.com' },
        { accountId: 't2:u2', email: 'b@test.com' },
      ],
      resolve: async (accountId) => ({
        manager: {},
        account: { accountId, email: accountId === 't1:u1' ? 'a@test.com' : 'b@test.com' },
      }),
    };

    const result = await fanOutAcrossAccounts(
      registry,
      { limit: 10 },
      async (_manager, account) => ({
        items: [{ subject: `mail-${account.accountId}`, receivedDateTime: account.accountId === 't1:u1' ? '2025-01-02' : '2025-01-01' }],
      }),
      { sortKey: 'receivedDateTime', sortDesc: true, resultKey: 'items', globalLimit: 10 }
    );

    expect(result.items).toHaveLength(2);
    expect(result.items[0].accountEmail).toBe('a@test.com');
    expect(result.fanOutSummary.fanOutMode).toBe('all-accounts');
  });
});
