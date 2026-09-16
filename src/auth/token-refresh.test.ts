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

  it.each([
    { state: 'missing', beforeToken: null },
    {
      state: 'expired',
      beforeToken: {
        token: 'expired-token',
        expiry: new Date(Date.now() - 60_000),
      },
    },
  ])('attempts HTTP refresh when the Substrate token is $state', async ({ beforeToken }) => {
    const refreshedToken = {
      token: 'refreshed-token',
      expiry: new Date(Date.now() + 60 * 60_000),
    };
    vi.mocked(extractSubstrateToken)
      .mockReturnValueOnce(beforeToken)
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
      expect(result.value.method).toBe('http');
      expect(result.value.previousExpiry).toBeInstanceOf(Date);
      expect(result.value.previousExpiry).toEqual(beforeToken?.expiry ?? new Date(0));
      expect(result.value.refreshNeeded).toBe(true);
    }
  });
});
