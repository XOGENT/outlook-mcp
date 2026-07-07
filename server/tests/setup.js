// Global test setup: guarantee tests never touch the real per-user data dir
// (~/.outlook-mcp). If a test file doesn't set MCP_OUTLOOK_DATA_DIR itself,
// default it to a unique temp dir so any accountRegistry/tokenManager writes
// stay sandboxed. Without this, a test that exercises the registry with the env
// unset would create ~/.outlook-mcp on the developer's machine — which then
// silently blocks the real install-store migration on the next server start.
import os from 'os';
import path from 'path';
import fs from 'fs';

if (!process.env.MCP_OUTLOOK_DATA_DIR) {
  const dir = path.join(os.tmpdir(), `outlook-mcp-test-${process.pid}`);
  fs.mkdirSync(dir, { recursive: true });
  process.env.MCP_OUTLOOK_DATA_DIR = dir;
}
