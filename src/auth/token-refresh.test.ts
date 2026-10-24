import { beforeEach, describe, expect, it, vi } from 'vitest';

vi.mock('./token-extractor.js', () => ({
  extractSubstrateToken: vi.fn(),
}));

vi.mock('./session-store.js', () => ({
  clearTokenCache: vi.fn(),
}));

vi.mock('../browser/context.js', () => ({ createBrowserContext: vi.fn(), closeBrowser: vi.fn() }));
vi.mock('../browser/auth.js', () => ({ ensureAuthenticated: vi.fn() }));

vi.mock('./token-refresh-http.js', () => ({
  refreshTokensViaHttp: vi.fn(),
}));

import { refreshTokensViaBrowser, withTokenRefreshScope } from './token-refresh.js';
import { extractSubstrateToken } from './token-extractor.js';
import { refreshTokensViaHttp } from './token-refresh-http.js';
import { createBrowserContext } from '../browser/context.js';

describe('refreshTokensViaBrowser', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it('uses HTTP refresh after the cached access token has expired', async () => {
    const refreshedExpiry = new Date(Date.now() + 60 * 60 * 1000);

    vi.mocked(extractSubstrateToken)
      .mockReturnValueOnce({ token: 'refreshed-access-token', expiry: refreshedExpiry });
    vi.mocked(refreshTokensViaHttp).mockResolvedValue({ ok: true, value: undefined });

    const result = await refreshTokensViaBrowser();

    expect(refreshTokensViaHttp).toHaveBeenCalledOnce();
    expect(createBrowserContext).not.toHaveBeenCalled();
    expect(result.ok).toBe(true);
  });

  it('shares one failed recovery within a request scope', async () => {
    vi.mocked(extractSubstrateToken).mockReturnValue(null);
    vi.mocked(refreshTokensViaHttp).mockResolvedValue({ ok: true, value: undefined });
    vi.mocked(createBrowserContext).mockRejectedValue(new Error('no browser'));

    await withTokenRefreshScope(async () => {
      const first = await refreshTokensViaBrowser();
      const second = await refreshTokensViaBrowser();
      expect(first.ok).toBe(false);
      expect(second).toBe(first);
    });

    expect(refreshTokensViaHttp).toHaveBeenCalledOnce();
    expect(createBrowserContext).toHaveBeenCalledOnce();
  });
});
