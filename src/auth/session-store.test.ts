/**
 * Tests for Teams session decoding and origin selection.
 */

import { webcrypto } from 'node:crypto';
import { describe, it, expect } from 'vitest';
import {
  decodeMsalCache,
  getTeamsOrigin,
  selectLiveTeamsOrigin,
  type SessionState,
} from './session-store.js';

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

  it('prefers fresh trusted cloud credentials', () => {
    const expired = Math.floor(Date.now() / 1000) - 3600;
    const state: SessionState = {
      cookies: [],
      origins: [
        makeOrigin('https://untrusted.example', [makeSubstrateEntry()]),
        makeOrigin('https://teams.microsoft.com', [makeSubstrateEntry(expired)]),
        makeOrigin('https://teams.cloud.microsoft', [makeSubstrateEntry()]),
      ],
    };

    expect(getTeamsOrigin(state)?.origin).toBe('https://teams.cloud.microsoft');
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

  it('selects only a trusted origin with a fresh token for live navigation', () => {
    const expired = Math.floor(Date.now() / 1000) - 3600;
    const state: SessionState = {
      cookies: [],
      origins: [
        makeOrigin('https://untrusted.example', [makeSubstrateEntry()]),
        makeOrigin('https://teams.cloud.microsoft', [makeSubstrateEntry(expired)]),
        makeOrigin('https://teams.microsoft.com', [makeSubstrateEntry()]),
      ],
    };

    expect(selectLiveTeamsOrigin(state)).toBe('https://teams.microsoft.com');
    expect(selectLiveTeamsOrigin({
      cookies: [],
      origins: [makeOrigin('https://teams.cloud.microsoft', [makeSubstrateEntry(expired)])],
    })).toBeNull();
  });
});

async function makeEncryptedSession() {
  const clientId = 'test-client';
  const key = new Uint8Array(32).fill(7);
  const nonce = new Uint8Array(16).fill(3);
  const baseKey = await webcrypto.subtle.importKey('raw', key, 'HKDF', false, ['deriveKey']);
  const derivedKey = await webcrypto.subtle.deriveKey({
    name: 'HKDF', salt: nonce, hash: 'SHA-256', info: new TextEncoder().encode(clientId),
  }, baseKey, { name: 'AES-GCM', length: 256 }, false, ['encrypt']);
  const token = makeSubstrateEntry();
  const ciphertext = await webcrypto.subtle.encrypt(
    { name: 'AES-GCM', iv: new Uint8Array(12) }, derivedKey, new TextEncoder().encode(token.value));
  const entry = {
    name: `msal.2|home|login.windows.net|accesstoken|${clientId}|tenant|scope|`,
    value: JSON.stringify({
      id: 'key-id', nonce: Buffer.from(nonce).toString('base64url'),
      data: Buffer.from(ciphertext).toString('base64url'), lastUpdatedAt: 'test',
    }),
  };
  const state: SessionState = {
    cookies: [{
      name: 'msal.cache.encryption', domain: 'teams.cloud.microsoft', path: '/',
      value: encodeURIComponent(JSON.stringify({ id: 'key-id', key: Buffer.from(key).toString('base64url') })),
    }],
    origins: [makeOrigin('https://teams.cloud.microsoft', [
      { name: `msal.2.token.keys.${clientId}`, value: '{}' }, entry, makeChatSvcEntry(),
    ])],
  };
  return { state, entry, token };
}

describe('decodeMsalCache', () => {
  it('decodes browser-encrypted credentials for existing origin selection without changing plaintext entries', async () => {
    const { state, entry, token } = await makeEncryptedSession();
    expect(selectLiveTeamsOrigin(state)).toBeNull();
    decodeMsalCache(state);
    expect(entry.value).toBe(token.value);
    expect(selectLiveTeamsOrigin(state)).toBe('https://teams.cloud.microsoft');
    expect(state.origins[0].localStorage[2]).toEqual(makeChatSvcEntry());
  });

  it.each(['cookie domain', 'cookie id', 'ciphertext', 'app context', 'untrusted origin'])(
    'does not decode credentials when %s is wrong',
    async (failure) => {
      const { state, entry } = await makeEncryptedSession();
      if (failure === 'cookie domain') state.cookies[0].domain = 'untrusted.example';
      if (failure === 'cookie id') state.cookies[0].value = state.cookies[0].value.replace('key-id', 'rotated-key');
      if (failure === 'ciphertext') entry.value = entry.value.replace('"data":"', '"data":"AQ');
      if (failure === 'app context') entry.name = entry.name.replace('test-client', 'another-client');
      if (failure === 'untrusted origin') {
        state.origins[0].origin = 'https://untrusted.example';
        state.cookies[0].domain = 'untrusted.example';
      }
      const encryptedValue = entry.value;
      decodeMsalCache(state);
      expect(entry.value).toBe(encryptedValue);
      expect(selectLiveTeamsOrigin(state)).toBeNull();
    },
  );
});
