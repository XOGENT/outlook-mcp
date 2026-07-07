import { describe, it, expect, vi, beforeEach } from 'vitest';
import { withSendDedupe, buildSendKey, resetSendDedupe } from '../../tools/common/sendDedupe.js';
import { createToolError } from '../../utils/mcpErrorResponse.js';

const OK = { content: [{ type: 'text', text: 'sent' }] };

describe('buildSendKey', () => {
  it('is deterministic and order-sensitive on content', () => {
    const a = buildSendKey({ accountId: 't1', to: ['x@test.com'], subject: 'Hi' });
    const b = buildSendKey({ accountId: 't1', to: ['x@test.com'], subject: 'Hi' });
    const c = buildSendKey({ accountId: 't1', to: ['x@test.com'], subject: 'Bye' });
    expect(a).toBe(b);
    expect(a).not.toBe(c);
  });
});

describe('withSendDedupe', () => {
  beforeEach(() => resetSendDedupe());

  it('runs the send and returns its result on first call', async () => {
    const sendFn = vi.fn().mockResolvedValue(OK);
    const result = await withSendDedupe('k', sendFn);
    expect(result).toBe(OK);
    expect(sendFn).toHaveBeenCalledTimes(1);
  });

  it('suppresses a concurrent duplicate while the first send is in flight', async () => {
    let release;
    const sendFn = vi.fn().mockImplementation(() => new Promise(r => { release = () => r(OK); }));

    const first = withSendDedupe('k', sendFn);
    const second = await withSendDedupe('k', sendFn); // in-flight -> suppressed

    expect(second.isError).toBe(true);
    expect(second._errorDetails?.duplicateSuppressed).toBe(true);
    expect(sendFn).toHaveBeenCalledTimes(1);

    release();
    await first;
  });

  it('returns the cached success on a repeat call without re-sending', async () => {
    const sendFn = vi.fn().mockResolvedValue(OK);
    await withSendDedupe('k', sendFn);
    const again = await withSendDedupe('k', sendFn);
    expect(again).toBe(OK);
    expect(sendFn).toHaveBeenCalledTimes(1);
  });

  it('caches an ambiguous outcome so a retry does NOT re-send', async () => {
    const ambiguous = createToolError('outcome unknown — check Sent Items', { ambiguousOutcome: true });
    const sendFn = vi.fn().mockResolvedValue(ambiguous);
    const first = await withSendDedupe('k', sendFn);
    const again = await withSendDedupe('k', sendFn);
    expect(first).toBe(ambiguous);
    expect(again).toBe(ambiguous);
    expect(sendFn).toHaveBeenCalledTimes(1);
  });

  it('does NOT cache a hard (non-ambiguous) error, allowing a retry to re-send', async () => {
    const hardError = createToolError('bad request', { retryable: false });
    const sendFn = vi.fn()
      .mockResolvedValueOnce(hardError)
      .mockResolvedValueOnce(OK);
    const first = await withSendDedupe('k', sendFn);
    const second = await withSendDedupe('k', sendFn);
    expect(first).toBe(hardError);
    expect(second).toBe(OK);
    expect(sendFn).toHaveBeenCalledTimes(2);
  });

  it('does NOT cache a thrown error, allowing a retry to re-send', async () => {
    const sendFn = vi.fn()
      .mockRejectedValueOnce(new Error('boom'))
      .mockResolvedValueOnce(OK);
    await expect(withSendDedupe('k', sendFn)).rejects.toThrow('boom');
    const retry = await withSendDedupe('k', sendFn);
    expect(retry).toBe(OK);
    expect(sendFn).toHaveBeenCalledTimes(2);
  });

  it('re-sends once a cached entry has expired past its TTL', async () => {
    let clock = 1000;
    const now = () => clock;
    const sendFn = vi.fn().mockResolvedValue(OK);

    await withSendDedupe('k', sendFn, { ttlMs: 100, now });
    clock += 50;
    await withSendDedupe('k', sendFn, { ttlMs: 100, now }); // still cached
    expect(sendFn).toHaveBeenCalledTimes(1);

    clock += 1000; // past TTL
    await withSendDedupe('k', sendFn, { ttlMs: 100, now });
    expect(sendFn).toHaveBeenCalledTimes(2);
  });
});
