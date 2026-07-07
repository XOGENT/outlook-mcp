import { describe, it, expect, vi } from 'vitest';
import { sendEmailTool } from '../../tools/email/sendEmail.js';
import { createToolError } from '../../utils/mcpErrorResponse.js';

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

  it('propagates a returned error instead of falsely reporting success', async () => {
    // makeRequest RETURNS handled errors (it does not throw); sendEmail must
    // surface them rather than always claiming the message was sent.
    const ambiguous = createToolError('outcome unknown — check Sent Items', { ambiguousOutcome: true });
    const registry = {
      listAccounts: vi.fn().mockResolvedValue([{ accountId: 't1:u1', email: 'a@test.com', userId: 'u1' }]),
      resolve: vi.fn().mockResolvedValue({
        account: { accountId: 't1:u1', email: 'a@test.com', userId: 'u1' },
        manager: {
          ensureAuthenticated: vi.fn(),
          getGraphApiClient: () => ({
            postWithRetry: vi.fn().mockResolvedValue(ambiguous),
            makeRequest: vi.fn().mockResolvedValue({ id: 'u1' }),
          }),
          getCurrentUser: () => ({ id: 'u1', accountId: 't1:u1' }),
        },
      }),
    };

    const result = await sendEmailTool(registry, {
      to: ['x@test.com'],
      subject: 'Hi',
      body: 'Hello',
      preserveUserStyling: false,
    });

    expect(result.isError).toBe(true);
    expect(result._errorDetails?.ambiguousOutcome).toBe(true);
    expect(result.content[0].text).not.toMatch(/sent successfully/i);
  });
});
