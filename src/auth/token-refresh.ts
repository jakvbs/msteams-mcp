/**
 * Token refresh orchestrator.
 * 
 * Tries HTTP-based refresh first (direct OAuth2 token endpoint call, ~100ms),
 * then falls back to headless browser refresh (~8s) if HTTP fails.
 * 
 * HTTP refresh works by extracting the MSAL refresh token from session state
 * and exchanging it for new access tokens via Azure AD's token endpoint.
 * This works identically for standard Microsoft login and corporate SSO
 * (ADFS/Okta federation) — refresh tokens are standard Azure AD tokens
 * regardless of how the user originally authenticated.
 * 
 * Browser fallback covers cases where:
 * - The refresh token has expired (typically after days/weeks of inactivity)
 * - Conditional Access policies require interactive auth
 * - The MSAL cache format has changed unexpectedly
 * 
 * First login always requires a browser — there's no refresh token to use yet.
 */

import { AsyncLocalStorage } from 'node:async_hooks';
import { TOKEN_REFRESH_THRESHOLD_MS } from '../constants.js';
import { AuthenticationRequiredError, ErrorCode, SignInLoopError, createError } from '../types/errors.js';
import { type Result, ok, err } from '../types/result.js';
import {
  extractSubstrateToken,
  clearTokenCache,
} from './token-extractor.js';
import { refreshTokensViaHttp } from './token-refresh-http.js';
import * as log from '../utils/logger.js';

/** Result of a successful token refresh. */
export interface TokenRefreshResult {
  /** New token expiry time. */
  newExpiry: Date;
  /** Previous expiry time, or the Unix epoch when no valid prior token existed. */
  previousExpiry: Date;
  /** Minutes gained by refresh. */
  minutesGained: number;
  /** Whether a refresh was actually needed (token was close to expiry). */
  refreshNeeded: boolean;
  /** Which method was used: 'http' or 'browser'. */
  method: 'http' | 'browser';
}

/** Module-level flag to prevent concurrent refresh attempts. */
let refreshInProgress = false;

// A tool guard and the server's auth retry share one recovery result, including
// failure. A later MCP request gets a fresh scope and may try again.
const requestRefresh = new AsyncLocalStorage<{
  result?: Promise<Result<TokenRefreshResult>>;
}>();

export function withTokenRefreshScope<T>(run: () => Promise<T>): Promise<T> {
  return requestRefresh.run({}, run);
}

/**
 * Refreshes tokens using HTTP-first strategy with browser fallback.
 * 
 * 1. Try direct HTTP refresh via OAuth2 token endpoint (~100ms)
 * 2. If HTTP fails, fall back to headless browser refresh (~8s)
 * 3. If both fail, return error directing to teams_login
 */
export function refreshTokensViaBrowser(): Promise<Result<TokenRefreshResult>> {
  const scope = requestRefresh.getStore();
  if (scope?.result) return scope.result;
  // A different request holding the browser is not a recovery attempt by this
  // request; do not cache that transient "in progress" rejection.
  if (!scope || refreshInProgress) return refreshTokens();
  scope.result = refreshTokens();
  return scope.result;
}

async function refreshTokens(): Promise<Result<TokenRefreshResult>> {
  // Prevent concurrent refresh attempts
  if (refreshInProgress) {
    return err(createError(
      ErrorCode.UNKNOWN,
      'Token refresh already in progress. Please wait and try again.',
      { retryable: true, suggestions: ['Wait a moment and retry the request'] }
    ));
  }

  // Keep the current token expiry for comparison when one is available.
  // HTTP refresh only requires an MSAL refresh token, so a missing or expired
  // access token must not prevent recovery.
  const beforeToken = extractSubstrateToken();
  const previousExpiry = beforeToken?.expiry;
  if (previousExpiry) {
    log.debug('token-refresh', `Current token expires at ${previousExpiry.toISOString()} (${Math.round((previousExpiry.getTime() - Date.now()) / 60000)} mins remaining)`);
  } else {
    log.debug('token-refresh', 'No valid Substrate token found; attempting HTTP refresh');
  }
  refreshInProgress = true;

  try {
    // ── Strategy 1: HTTP refresh (fast, no browser needed) ──────────────
    const httpResult = await refreshTokensViaHttp();

    if (httpResult.ok) {
      log.info('token-refresh', `HTTP refresh succeeded: ${httpResult.value.tokensRefreshed} tokens refreshed` +
        (httpResult.value.skypeTokenRefreshed ? ', skype token refreshed' : '') +
        (httpResult.value.refreshTokenRotated ? ', refresh token rotated' : ''));

      // Verify we now have a valid Substrate token
      const afterToken = extractSubstrateToken();
      if (afterToken && afterToken.expiry.getTime() > Date.now()) {
        const minutesGained = previousExpiry
          ? Math.round((afterToken.expiry.getTime() - previousExpiry.getTime()) / 1000 / 60)
          : 0;
        const wasCloseToExpiry = !previousExpiry ||
          previousExpiry.getTime() - Date.now() < TOKEN_REFRESH_THRESHOLD_MS;

        return ok({
          newExpiry: afterToken.expiry,
          previousExpiry: previousExpiry ?? new Date(0),
          minutesGained,
          refreshNeeded: wasCloseToExpiry,
          method: 'http',
        });
      }

      // HTTP refresh reported success but we can't extract a valid token — fall through
      log.warn('token-refresh', 'HTTP refresh reported success but no valid Substrate token found, falling back to browser');
    } else {
      log.warn('token-refresh', `HTTP refresh failed: ${httpResult.error.message}, falling back to browser`);

      // If the error is definitively an auth error (expired refresh token),
      // don't bother with browser fallback — it won't help either
      if (httpResult.error.code === ErrorCode.AUTH_EXPIRED) {
        // Still try browser — the persistent profile's session cookies may work
        log.info('token-refresh', 'Auth expired, but trying browser fallback (session cookies may still be valid)');
      }
    }

    // ── Strategy 2: Browser refresh (fallback) ──────────────────────────
    const browserResult = await refreshTokensViaBrowserImpl(previousExpiry);
    if (!browserResult.ok) {
      log.warn('token-refresh', browserResult.error.message);
    }
    return browserResult;

  } finally {
    refreshInProgress = false;
  }
}

/**
 * Browser-based token refresh implementation.
 * Opens a headless browser with the persistent profile, lets MSAL
 * silently refresh tokens using session cookies.
 */
async function refreshTokensViaBrowserImpl(
  previousExpiry?: Date,
): Promise<Result<TokenRefreshResult>> {
  // Import browser functions dynamically to avoid circular dependencies
  const { createBrowserContext, closeBrowser } = await import('../browser/context.js');

  let manager: Awaited<ReturnType<typeof createBrowserContext>> | null = null;

  try {
    // Open headless browser with persistent profile
    manager = await createBrowserContext({ headless: true });

    // Import auth functions
    const { ensureAuthenticated } = await import('../browser/auth.js');

    // Use the same auth flow that works for login - this triggers token acquisition
    // showOverlay: false since headless browser has no visible window
    // headless: true to fail fast if user interaction is required
    await ensureAuthenticated(manager.page, manager.context, (msg) => {
      // Silent logging for headless refresh
      log.debug('token-refresh', msg);
    }, false, true);

    // Close browser (ensureAuthenticated already saved the session)
    await closeBrowser(manager, false);
    manager = null;

    // Clear our token cache to force re-extraction from the new session
    clearTokenCache();

    // Extract the new token to verify we still have valid tokens
    const afterToken = extractSubstrateToken();
    if (!afterToken) {
      return err(createError(
        ErrorCode.AUTH_EXPIRED,
        'ACTION REQUIRED: Token refresh failed. You MUST call teams_login to re-authenticate.',
      ));
    }

    const newExpiry = afterToken.expiry;
    const minutesGained = previousExpiry
      ? Math.round((newExpiry.getTime() - previousExpiry.getTime()) / 1000 / 60)
      : 0;

    // Missing tokens always need refresh; otherwise use the normal threshold.
    const wasCloseToExpiry = !previousExpiry ||
      previousExpiry.getTime() - Date.now() < TOKEN_REFRESH_THRESHOLD_MS;

    // If we had a previous token that needed refresh, require a newer one.
    if (previousExpiry && wasCloseToExpiry && newExpiry.getTime() <= previousExpiry.getTime()) {
      return err(createError(
        ErrorCode.AUTH_EXPIRED,
        'ACTION REQUIRED: Token was not refreshed despite being close to expiry. You MUST call teams_login to re-authenticate.',
      ));
    }

    return ok({
      newExpiry,
      previousExpiry: previousExpiry ?? new Date(0),
      minutesGained,
      refreshNeeded: wasCloseToExpiry,
      method: 'browser',
    });

  } catch (error) {
    const pageClosed = manager?.page.isClosed() ?? false;
    // Clean up browser if still open
    if (manager) {
      try {
        await closeBrowser(manager, false);
      } catch {
        // Ignore cleanup errors
      }
    }

    const message = error instanceof Error ? error.message : 'Unknown error';
    if (error instanceof SignInLoopError) {
      return err(error.toMcpError());
    }
    if (error instanceof AuthenticationRequiredError) {
      return err(createError(ErrorCode.AUTH_REQUIRED, error.message));
    }
    return err(createError(
      ErrorCode.BROWSER_ERROR,
      pageClosed
        ? 'Browser sign-in was interrupted: the tab or browser was closed.'
        : `Token refresh via browser failed: ${message}`,
    ));
  }
}
