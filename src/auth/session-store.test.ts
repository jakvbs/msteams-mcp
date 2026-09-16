/**
 * Tests for getTeamsOrigin — especially New Teams (teams.cloud.microsoft)
 * SubstrateSearch token selection when multiple origins are present.
 */

import { describe, it, expect } from 'vitest';
import { getTeamsOrigin, type SessionState } from './session-store.js';

function makeOrigin(
  origin: string,
  localStorage: SessionState['origins'][number]['localStorage'] = []
): SessionState['origins'][number] {
  return { origin, localStorage };
}

function makeSubstrateEntry(expiresAt = Math.floor(Date.now() / 1000) + 3600) {
  const payload = Buffer.from(JSON.stringify({ sub: 'test', exp: expiresAt })).toString('base64url');
  return {
    name: 'msal-substrate-token',
    value: JSON.stringify({
      credentialType: 'AccessToken',
      target: 'https://substrate.office.com/SubstrateSearch-Internal.ReadWrite',
      secret: `eyJhbGciOiJub25lIn0.${payload}.sig`,
    }),
  };
}

function makeChatSvcEntry() {
  return {
    name: 'msal-chatsvc-token',
    value: JSON.stringify({
      credentialType: 'AccessToken',
      target: 'https://chatsvcagg.teams.microsoft.com/.default',
      secret: 'eyJhbGciOiJub25lIn0.eyJzdWIiOiJjaGF0In0.sig',
    }),
  };
}

describe('getTeamsOrigin', () => {
  it('prefers the origin that holds a SubstrateSearch token', () => {
    const state: SessionState = {
      cookies: [],
      origins: [
        makeOrigin('https://teams.microsoft.com', [makeChatSvcEntry()]),
        makeOrigin('https://teams.cloud.microsoft', [makeSubstrateEntry()]),
      ],
    };

    const chosen = getTeamsOrigin(state);
    expect(chosen?.origin).toBe('https://teams.cloud.microsoft');
  });

  it.each([
    ['legacy first', ['https://teams.microsoft.com', 'https://teams.cloud.microsoft']],
    ['cloud first', ['https://teams.cloud.microsoft', 'https://teams.microsoft.com']],
  ])('prefers fresh cloud credentials over expired legacy credentials (%s)', (_label, origins) => {
    const expired = Math.floor(Date.now() / 1000) - 3600;
    const state: SessionState = {
      cookies: [],
      origins: origins.map(origin => makeOrigin(origin, [
        makeSubstrateEntry(origin === 'https://teams.microsoft.com' ? expired : undefined),
      ])),
    };

    expect(getTeamsOrigin(state)?.origin).toBe('https://teams.cloud.microsoft');
  });

  it('ignores untrusted origins when preferring fresh credentials', () => {
    const expired = Math.floor(Date.now() / 1000) - 3600;
    const state: SessionState = {
      cookies: [],
      origins: [
        makeOrigin('https://untrusted.example', [makeSubstrateEntry()]),
        makeOrigin('https://teams.microsoft.com', [makeSubstrateEntry(expired)]),
      ],
    };

    expect(getTeamsOrigin(state)?.origin).toBe('https://teams.microsoft.com');
  });

  it('retains an expired token origin for refresh metadata when none are fresh', () => {
    const expired = Math.floor(Date.now() / 1000) - 3600;
    const state: SessionState = {
      cookies: [],
      origins: [
        makeOrigin('https://teams.cloud.microsoft', [makeChatSvcEntry()]),
        makeOrigin('https://teams.microsoft.com', [makeSubstrateEntry(expired)]),
      ],
    };

    expect(getTeamsOrigin(state)?.origin).toBe('https://teams.microsoft.com');
  });

  it('falls back to teams.cloud.microsoft when no Substrate token exists', () => {
    const state: SessionState = {
      cookies: [],
      origins: [
        makeOrigin('https://teams.microsoft.com', [makeChatSvcEntry()]),
        makeOrigin('https://teams.cloud.microsoft', [makeChatSvcEntry()]),
      ],
    };

    const chosen = getTeamsOrigin(state);
    expect(chosen?.origin).toBe('https://teams.cloud.microsoft');
  });

  it('returns teams.microsoft.com when it is the only known origin', () => {
    const state: SessionState = {
      cookies: [],
      origins: [makeOrigin('https://teams.microsoft.com', [makeChatSvcEntry()])],
    };

    const chosen = getTeamsOrigin(state);
    expect(chosen?.origin).toBe('https://teams.microsoft.com');
  });

  it('returns null when no Teams origins are present', () => {
    const state: SessionState = {
      cookies: [],
      origins: [makeOrigin('https://login.microsoftonline.com', [])],
    };

    expect(getTeamsOrigin(state)).toBeNull();
  });
});
