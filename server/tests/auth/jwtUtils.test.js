import { describe, it, expect } from 'vitest';
import { buildAccountId, decodeJwtPayload, extractAccountClaims } from '../../auth/jwtUtils.js';

describe('jwtUtils', () => {
  it('builds stable account id', () => {
    expect(buildAccountId('tenant', 'oid')).toBe('tenant:oid');
  });

  it('decodes jwt payload', () => {
    const header = Buffer.from(JSON.stringify({ alg: 'none' })).toString('base64url');
    const payload = Buffer.from(JSON.stringify({ tid: 't1', oid: 'o1', preferred_username: 'u@test.com' })).toString('base64url');
    const token = `${header}.${payload}.sig`;
    const claims = extractAccountClaims(token);
    expect(claims.accountId).toBe('t1:o1');
    expect(claims.email).toBe('u@test.com');
  });

  it('returns null for invalid token', () => {
    expect(decodeJwtPayload('bad')).toBeNull();
  });
});
