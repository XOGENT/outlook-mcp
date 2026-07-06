import { describe, it, expect } from 'vitest';
import { resolveScopes, CORE_DELEGATED_SCOPES, SHAREPOINT_SCOPES } from '../../auth/config.js';

function withEnv(vars, fn) {
  const saved = {};
  for (const [k, v] of Object.entries(vars)) {
    saved[k] = process.env[k];
    if (v === undefined) delete process.env[k];
    else process.env[k] = v;
  }
  try {
    return fn();
  } finally {
    for (const [k, v] of Object.entries(saved)) {
      if (v === undefined) delete process.env[k];
      else process.env[k] = v;
    }
  }
}

describe('OAuth scope resolution', () => {
  it('defaults to the admin-consent-free core set (no SharePoint Sites.*)', () => {
    withEnv({ OUTLOOK_MCP_ENABLE_SHAREPOINT: undefined, OUTLOOK_MCP_EXTRA_SCOPES: undefined }, () => {
      const scopes = resolveScopes().split(' ');
      for (const s of CORE_DELEGATED_SCOPES) expect(scopes).toContain(s);
      for (const s of SHAREPOINT_SCOPES) expect(scopes).not.toContain(s);
    });
  });

  it('adds SharePoint scopes when explicitly enabled', () => {
    withEnv({ OUTLOOK_MCP_ENABLE_SHAREPOINT: 'true', OUTLOOK_MCP_EXTRA_SCOPES: undefined }, () => {
      const scopes = resolveScopes().split(' ');
      for (const s of SHAREPOINT_SCOPES) expect(scopes).toContain(s);
    });
  });

  it('appends and de-duplicates extra scopes', () => {
    withEnv({ OUTLOOK_MCP_ENABLE_SHAREPOINT: undefined, OUTLOOK_MCP_EXTRA_SCOPES: 'Mail.Read, Presence.Read' }, () => {
      const scopes = resolveScopes().split(' ');
      expect(scopes).toContain('Presence.Read');
      expect(scopes.filter((s) => s === 'Mail.Read')).toHaveLength(1);
    });
  });
});
