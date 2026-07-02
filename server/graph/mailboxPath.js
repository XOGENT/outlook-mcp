export function buildMailboxBase(mailbox) {
  if (!mailbox || mailbox === 'me') {
    return '/me';
  }
  return `/users/${encodeURIComponent(mailbox)}`;
}

export function buildMailboxPath(mailbox, ...segments) {
  const base = buildMailboxBase(mailbox);
  if (segments.length === 0) return base;
  const path = segments.join('/');
  return `${base}/${path}`;
}

export function getCacheKey(accountId, mailbox) {
  return `${accountId}:${mailbox || 'me'}`;
}
