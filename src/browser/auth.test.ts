import { beforeEach, describe, expect, it, vi } from 'vitest';
import type { Page } from 'playwright';

vi.mock('./context.js', () => ({
  saveSessionState: vi.fn(),
  readOriginStorage: (page: { evaluate: () => unknown }) => page.evaluate(),
}));
vi.mock('../auth/token-extractor.js', () => ({ extractSubstrateToken: vi.fn() }));
vi.mock('../auth/session-store.js', async importOriginal => ({
  ...await importOriginal<typeof import('../auth/session-store.js')>(),
  getLiveTeamsOrigin: () => null,
}));

import { extractSubstrateToken } from '../auth/token-extractor.js';
import { ensureAuthenticated, getAuthStatus, isTeamsUrl } from './auth.js';

/** Minutes left on the MSAL token the page holds; -1 when there is none. */
let tokenMinutes: () => number;

beforeEach(() => {
  // The browser snapshot carries the token; the saved-session read finds none.
  vi.mocked(extractSubstrateToken).mockImplementation(state => {
    if (!state) return null;
    const minutes = tokenMinutes();
    return minutes > 0 ? { token: 'token', expiry: new Date(Date.now() + minutes * 60_000) } as never : null;
  });
});

/** A page on `url` with no Teams DOM markers whose storage holds the token `tokenMinutes` reports. */
function fakePage(url: string, overrides: Partial<Record<keyof Page, unknown>> = {}): Page {
  return {
    url: () => url,
    locator: () => ({ count: async () => 0 }),
    evaluate: async () => ({ origin: new URL(url).origin, localStorage: [] }),
    context: () => ({ cookies: async () => [] }),
    ...overrides,
  } as unknown as Page;
}

describe('isTeamsUrl', () => {
  it('covers every host Microsoft redirects Teams users to', () => {
    // New Teams settles on the bare host, with no path.
    for (const url of [
      'https://teams.microsoft.com/',
      'https://teams.cloud.microsoft/',
      'https://teams.chat.microsoft.com/',
      'https://teams.microsoft.us/',
      'https://dod.teams.microsoft.us/',
    ]) {
      expect(isTeamsUrl(url), url).toBe(true);
    }
    expect(isTeamsUrl('https://outlook.office.com/')).toBe(false);
  });

  it('matches on the host, whatever path the app has routed to', () => {
    expect(isTeamsUrl('https://teams.microsoft.com/v2/#/conversations/19:abc@thread.tacv2')).toBe(true);
  });
});

describe('getAuthStatus', () => {
  it('detects a completed login from the MSAL token alone', async () => {
    tokenMinutes = () => 55;
    const status = await getAuthStatus(fakePage('https://teams.cloud.microsoft/'));

    expect(status.isAuthenticated).toBe(true);
    expect(status.isOnLoginPage).toBe(false);
  });

  it('stays unauthenticated on a Teams host with no token', async () => {
    tokenMinutes = () => -1;
    const status = await getAuthStatus(fakePage('https://teams.cloud.microsoft/'));

    expect(status.isAuthenticated).toBe(false);
  });

  it('reports the login page', async () => {
    tokenMinutes = () => 99;
    const status = await getAuthStatus(fakePage('https://login.microsoftonline.com/common/oauth2'));

    expect(status).toMatchObject({ isAuthenticated: false, isOnLoginPage: true });
  });
});

describe('ensureAuthenticated', () => {
  it('keeps a visible browser open for sign-in when Teams loads with expired tokens', async () => {
    // Teams can render without redirecting to login while the Microsoft SSO
    // session is gone; only a fresh MSAL token proves the user has signed in.
    const minutes = [-1, -1, 55];
    tokenMinutes = () => minutes.length > 1 ? minutes.shift()! : minutes[0];
    const page = fakePage('https://teams.cloud.microsoft/', {
      goto: async () => null,
      waitForURL: async () => { throw new Error('timeout'); },
      waitForTimeout: async () => {},
    });
    const messages: string[] = [];

    await ensureAuthenticated(page, {} as never, msg => messages.push(msg), false, false);

    expect(messages).toContain('Waiting for manual login...');
    expect(messages).toContain('Authentication successful!');
    expect(minutes).toEqual([55]);
  });
});
