export function decodeJwtPayload(token) {
  try {
    const parts = token.split('.');
    if (parts.length < 2) return null;
    const payload = parts[1].replace(/-/g, '+').replace(/_/g, '/');
    const padded = payload + '='.repeat((4 - (payload.length % 4)) % 4);
    return JSON.parse(Buffer.from(padded, 'base64').toString('utf8'));
  } catch {
    return null;
  }
}

export function buildAccountId(tenantId, userId) {
  return `${tenantId}:${userId}`;
}

export function extractAccountClaims(accessToken) {
  const claims = decodeJwtPayload(accessToken);
  if (!claims) return null;
  const tenantId = claims.tid;
  const userId = claims.oid || claims.sub;
  const email = claims.preferred_username || claims.upn || claims.email;
  if (!tenantId || !userId) return null;
  return {
    accountId: buildAccountId(tenantId, userId),
    tenantId,
    userId,
    email: email || '',
    displayName: claims.name || email || '',
  };
}
