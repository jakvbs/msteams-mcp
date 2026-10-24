/**
 * Tests for getTeamsOrigin — especially New Teams (teams.cloud.microsoft)
 * SubstrateSearch token selection when multiple origins are present.
 */

import { describe, it, expect } from 'vitest';
import { getTeamsOrigin, scopeSessionStateToTeams, type SessionState } from './session-store.js';

function makeOrigin(
  origin: string,
  localStorage: SessionState['origins'][number]['localStorage'] = []
): SessionState['origins'][number] {
  return { origin, localStorage };
}

function makeSubstrateEntry(target = 'https://substrate.office.com/SubstrateSearch-Internal.ReadWrite') {
  return {
    name: 'msal-substrate-token',
    value: JSON.stringify({
      credentialType: 'AccessToken',
      target,
      // JWT-shaped secret (header.payload.sig) — extractor only checks prefix
      secret: 'eyJhbGciOiJub25lIn0.eyJzdWIiOiJ0ZXN0In0.sig',
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

describe('scopeSessionStateToTeams', () => {
  it('drops cookies and local storage of non-Teams sites from an attached browser profile', () => {
    const cookie = (name: string, domain: string) => ({ name, value: 'v', domain, path: '/' });
    const state: SessionState = {
      cookies: [
        cookie('skypetoken_asm', '.asyncgw.teams.microsoft.com'),
        cookie('authtoken', 'teams.microsoft.com'),
        cookie('msal.cache.encryption', 'teams.cloud.microsoft'),
        cookie('SID', '.google.com'),
        cookie('ESTSAUTH', '.login.microsoftonline.com'),
        cookie('session', 'evilteams.microsoft.com.example'),
      ],
      origins: [
        makeOrigin('https://teams.cloud.microsoft', [makeChatSvcEntry()]),
        makeOrigin('https://www.youtube.com', [{ name: 'yt', value: 'v' }]),
      ],
    };

    expect(scopeSessionStateToTeams(state)).toEqual({
      cookies: [
        cookie('skypetoken_asm', '.asyncgw.teams.microsoft.com'),
        cookie('authtoken', 'teams.microsoft.com'),
        cookie('msal.cache.encryption', 'teams.cloud.microsoft'),
      ],
      origins: [makeOrigin('https://teams.cloud.microsoft', [makeChatSvcEntry()])],
    });
  });
});
