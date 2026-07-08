import { describe, it, expect } from 'vitest';
import { readFileSync, readdirSync, statSync } from 'fs';
import { fileURLToPath } from 'url';
import path from 'path';

// This is a stdio MCP server: stdout carries the JSON-RPC protocol stream and
// must contain ONLY protocol frames. Any console.log / process.stdout.write in
// runtime code injects plain text into that stream, corrupting it — the client
// then drops the transport mid-call ("sent successfully but transport error")
// and retries, which duplicated email sends. This test fails if any runtime
// source file reintroduces a stdout write, so the bug can't regress silently.

const serverRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..', '..');

// index.js legitimately reassigns console.log -> console.error as the guard, so
// it is exempt (it is the fix, not a violation).
const EXEMPT = new Set([path.join(serverRoot, 'index.js')]);

function collectJsFiles(dir) {
  const out = [];
  for (const entry of readdirSync(dir)) {
    const full = path.join(dir, entry);
    const st = statSync(full);
    if (st.isDirectory()) {
      if (entry === 'test' || entry === 'tests' || entry === 'node_modules') continue;
      out.push(...collectJsFiles(full));
    } else if (entry.endsWith('.js')) {
      out.push(full);
    }
  }
  return out;
}

describe('stdio safety: no stdout writes in runtime code', () => {
  const files = collectJsFiles(serverRoot).filter(f => !EXEMPT.has(f));

  it('finds runtime source files to scan', () => {
    expect(files.length).toBeGreaterThan(0);
  });

  for (const file of files) {
    const rel = path.relative(serverRoot, file);
    it(`${rel} does not write to stdout`, () => {
      const src = readFileSync(file, 'utf8');
      // Strip line comments so a comment mentioning console.log doesn't trip us.
      const code = src.replace(/\/\/.*$/gm, '');
      expect(code, `${rel} must log to stderr (console.error), not stdout`).not.toMatch(/console\s*\.\s*log\s*\(/);
      expect(code, `${rel} must not write directly to process.stdout`).not.toMatch(/process\s*\.\s*stdout\s*\.\s*write\s*\(/);
    });
  }
});
