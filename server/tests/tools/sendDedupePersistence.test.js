import { describe, it, expect, vi, afterEach } from 'vitest';
import os from 'os';
import path from 'path';
import fs from 'fs';
import { withSendDedupe, resetSendDedupe, _setSendDedupeStorePath } from '../../tools/common/sendDedupe.js';

const OK = { content: [{ type: 'text', text: 'sent' }] };

function tmpJournal(tag) {
  return path.join(os.tmpdir(), 'send-dedupe-' + process.pid + '-' + tag + '.json');
}

afterEach(() => resetSendDedupe());

describe('sendDedupe persistence across process restart', () => {
  it('does NOT re-send a succeeded message after a simulated restart', async () => {
    const journal = tmpJournal('ok');
    _setSendDedupeStorePath(journal);
    resetSendDedupe();
    _setSendDedupeStorePath(journal);

    const send1 = vi.fn().mockResolvedValue(OK);
    await withSendDedupe('k', send1);
    expect(send1).toHaveBeenCalledTimes(1);
    expect(fs.existsSync(journal)).toBe(true);

    // Simulate a fresh process: same journal file, empty in-memory state.
    _setSendDedupeStorePath(journal);

    const send2 = vi.fn().mockResolvedValue(OK);
    const result = await withSendDedupe('k', send2);
    expect(send2).not.toHaveBeenCalled();
    expect(result.content[0].text).toBe('sent');
  });

  it('treats a send left pending by a crashed process as ambiguous (no re-send)', async () => {
    const journal = tmpJournal('pending');
    _setSendDedupeStorePath(journal);
    resetSendDedupe();
    _setSendDedupeStorePath(journal);

    // Start a send that never resolves; the pending marker is persisted before await.
    let hang;
    const stuck = vi.fn().mockImplementation(() => new Promise(r => { hang = r; }));
    void withSendDedupe('k', stuck);
    expect(fs.existsSync(journal)).toBe(true);

    // Simulate a restart while the send outcome is still unknown.
    _setSendDedupeStorePath(journal);

    const retry = vi.fn().mockResolvedValue(OK);
    const result = await withSendDedupe('k', retry);
    expect(retry).not.toHaveBeenCalled();
    expect(result.isError).toBe(true);
    expect(result._errorDetails?.ambiguousOutcome).toBe(true);

    if (hang) hang(OK);
  });
});
