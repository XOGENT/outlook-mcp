import { describe, it, expect } from 'vitest';
import { GraphApiClient } from '../../graph/graphClient.js';
import { authConfig } from '../../auth/config.js';

// Build a GraphApiClient whose underlying Graph request always rejects with the
// given error, counting how many times a verb was invoked (i.e. retries).
function stubClient(error) {
  const client = new GraphApiClient({ refreshAccessToken: async () => {} });
  client.initialize = async () => {};
  client.enforceRateLimit = async () => {};
  client.sleep = async () => {}; // skip real backoff delays
  let calls = 0;
  const req = {
    header() { return this; },
    select() { return this; },
    top() { return this; },
    filter() { return this; },
    orderby() { return this; },
    expand() { return this; },
    search() { return this; },
    query() { return this; },
    post: async () => { calls++; throw error; },
    patch: async () => { calls++; throw error; },
    get: async () => { calls++; throw error; },
  };
  client.client = { api: () => req };
  client.getCalls = () => calls;
  return client;
}

describe('makeRequest retry idempotency', () => {
  it('does NOT retry a POST (sendMail) on a 5xx and reports an ambiguous outcome', async () => {
    const client = stubClient(Object.assign(new Error('gateway timeout'), { status: 504 }));
    const res = await client.makeRequest('/me/sendMail', { body: {} }, 'POST');
    expect(client.getCalls()).toBe(1);
    expect(res.isError).toBe(true);
    expect(res._errorDetails?.ambiguousOutcome).toBe(true);
    expect(res.content[0].text).toMatch(/Sent Items/i);
  });

  it('does NOT retry a POST on a dropped connection (no HTTP status)', async () => {
    const client = stubClient(new Error('socket hang up')); // no .status
    const res = await client.makeRequest('/me/sendMail', { body: {} }, 'POST');
    expect(client.getCalls()).toBe(1);
    expect(res.isError).toBe(true);
    expect(res._errorDetails?.ambiguousOutcome).toBe(true);
  });

  it('DOES retry an idempotent GET on a 5xx', async () => {
    const client = stubClient(Object.assign(new Error('boom'), { status: 500 }));
    const res = await client.makeRequest('/me', {}, 'GET');
    expect(client.getCalls()).toBe(authConfig.retry.maxAttempts + 1);
    expect(res.isError).toBe(true);
    expect(res._errorDetails?.ambiguousOutcome).toBeUndefined();
  });

  it('still retries a POST on 429 (throttled before processing, so safe)', async () => {
    const client = stubClient(Object.assign(new Error('too many'), { status: 429 }));
    const res = await client.makeRequest('/me/sendMail', { body: {} }, 'POST');
    expect(client.getCalls()).toBe(authConfig.retry.maxAttempts + 1);
    expect(res.isError).toBe(true);
    expect(res._errorDetails?.ambiguousOutcome).toBeUndefined();
  });

  it('retries a POST 5xx when the caller marks it idempotent', async () => {
    const client = stubClient(Object.assign(new Error('boom'), { status: 503 }));
    const res = await client.makeRequest('/me/some-safe-op', { body: {}, idempotent: true }, 'POST');
    expect(client.getCalls()).toBe(authConfig.retry.maxAttempts + 1);
    expect(res.isError).toBe(true);
  });
});
