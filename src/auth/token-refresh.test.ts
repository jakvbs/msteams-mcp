import { beforeEach, describe, expect, it, vi } from 'vitest';
import { ok } from '../types/result.js';

vi.mock('./token-extractor.js', () => ({
  extractSubstrateToken: vi.fn(),
  clearTokenCache: vi.fn(),
}));

vi.mock('./token-refresh-http.js', () => ({
  refreshTokensViaHttp: vi.fn(),
}));

import { extractSubstrateToken } from './token-extractor.js';
import { refreshTokensViaHttp } from './token-refresh-http.js';
import { refreshTokensViaBrowser } from './token-refresh.js';

describe('refreshTokensViaBrowser', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it('attempts HTTP refresh when the Substrate token is missing', async () => {
    const refreshedToken = {
      token: 'refreshed-token',
      expiry: new Date(Date.now() + 60 * 60_000),
    };
    vi.mocked(extractSubstrateToken)
      .mockReturnValueOnce(null)
      .mockReturnValueOnce(refreshedToken);
    vi.mocked(refreshTokensViaHttp).mockResolvedValue(ok({
      tokensRefreshed: 3,
      skypeTokenRefreshed: true,
      refreshTokenRotated: false,
    }));

    const result = await refreshTokensViaBrowser();

    expect(refreshTokensViaHttp).toHaveBeenCalledOnce();
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value.previousExpiry).toEqual(new Date(0));
      expect(result.value.refreshNeeded).toBe(true);
    }
  });
});
