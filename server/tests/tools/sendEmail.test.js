import { describe, it, expect, vi } from 'vitest';
import { sendEmailTool } from '../../tools/email/sendEmail.js';

function mockRegistry(accounts) {
  return {
    listAccounts: vi.fn().mockResolvedValue(accounts),
    resolve: vi.fn().mockImplementation(async (accountId) => {
      const account = accounts.find(a => a.accountId === accountId) || accounts[0];
      return {
        account,
        manager: {
          ensureAuthenticated: vi.fn(),
          getGraphApiClient: () => ({
            postWithRetry: vi.fn().mockResolvedValue({}),
            makeRequest: vi.fn().mockResolvedValue({ id: account.userId }),
          }),
          getCurrentUser: () => ({ id: account.userId, accountId: account.accountId }),
        },
      };
    }),
  };
}

describe('sendEmailTool', () => {
  it('requires account when multiple accounts connected', async () => {
    const registry = mockRegistry([
      { accountId: 't1:u1', email: 'a@test.com', userId: 'u1' },
      { accountId: 't2:u2', email: 'b@test.com', userId: 'u2' },
    ]);

    const result = await sendEmailTool(registry, {
      to: ['x@test.com'],
      subject: 'Hi',
      body: 'Hello',
    });

    expect(result.isError).toBe(true);
    expect(result.content[0].text).toContain('account');
  });

  it('sends from single account without account param', async () => {
    const registry = mockRegistry([
      { accountId: 't1:u1', email: 'a@test.com', userId: 'u1' },
    ]);

    const result = await sendEmailTool(registry, {
      to: ['x@test.com'],
      subject: 'Hi',
      body: 'Hello',
      preserveUserStyling: false,
    });

    expect(result.isError).toBeUndefined();
    const payload = JSON.parse(result.content[0].text);
    expect(payload.fromAccount.email).toBe('a@test.com');
  });
});
