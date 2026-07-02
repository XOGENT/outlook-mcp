import { describe, it, expect } from 'vitest';
import { buildMailboxBase, buildMailboxPath, getCacheKey } from '../../graph/mailboxPath.js';

describe('mailboxPath', () => {
  it('returns /me for primary mailbox', () => {
    expect(buildMailboxBase()).toBe('/me');
    expect(buildMailboxBase('me')).toBe('/me');
  });

  it('returns /users path for shared mailbox', () => {
    expect(buildMailboxBase('billing@contoso.com')).toBe('/users/billing%40contoso.com');
  });

  it('builds nested paths', () => {
    expect(buildMailboxPath(null, 'messages', 'abc')).toBe('/me/messages/abc');
    expect(buildMailboxPath('user@x.com', 'mailFolders', 'inbox')).toBe('/users/user%40x.com/mailFolders/inbox');
  });

  it('builds cache keys', () => {
    expect(getCacheKey('acct1', null)).toBe('acct1:me');
    expect(getCacheKey('acct1', 'shared@x.com')).toBe('acct1:shared@x.com');
  });
});
