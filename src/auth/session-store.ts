/**
 * Secure session state storage.
 * 
 * Handles reading and writing session state with:
 * - Encryption at rest
 * - Restricted file permissions
 * - Automatic migration from plaintext
 * 
 * Session files are stored in a user-specific config directory (~/.teams-mcp-server/)
 * to ensure consistency regardless of how the server is invoked (npx, global install, etc.).
 */

import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { createDecipheriv, hkdfSync } from 'node:crypto';
import { z } from 'zod';
import { encrypt, decrypt } from './crypto.js';
import { SESSION_EXPIRY_HOURS, TEAMS_HOSTS } from '../constants.js';
import * as log from '../utils/logger.js';

/**
 * User-specific config directory for teams-mcp-server.
 * - Windows: %APPDATA%\teams-mcp-server\
 * - macOS/Linux: ~/.teams-mcp-server/
 */
export const CONFIG_DIR = process.platform === 'win32'
  ? path.join(process.env.APPDATA || path.join(os.homedir(), 'AppData', 'Roaming'), 'teams-mcp-server')
  : path.join(os.homedir(), '.teams-mcp-server');
export const SESSION_STATE_PATH = path.join(CONFIG_DIR, 'session-state.json');
export const TOKEN_CACHE_PATH = path.join(CONFIG_DIR, 'token-cache.json');

/** File permission mode: owner read/write only. */
const SECURE_FILE_MODE = 0o600;

/**
 * Ensures the config directory exists with secure permissions.
 */
export function ensureConfigDir(): void {
  fs.mkdirSync(CONFIG_DIR, { recursive: true, mode: 0o700 });
}

/** Session state as stored by Playwright. */
export interface SessionState {
  cookies: Array<{
    name: string;
    value: string;
    domain?: string;
    path?: string;
    expires?: number;
    httpOnly?: boolean;
    secure?: boolean;
    sameSite?: 'Strict' | 'Lax' | 'None';
  }>;
  origins: Array<{
    origin: string;
    localStorage: Array<{ name: string; value: string }>;
  }>;
}

/** Token cache structure. */
export interface TokenCache {
  substrateToken: string;
  substrateTokenExpiry: number;
  extractedAt: number;
}

/**
 * Writes data securely with encryption and file permissions.
 */
function writeSecure(filePath: string, data: unknown): void {
  const json = JSON.stringify(data, null, 2);
  const encrypted = encrypt(json);
  
  fs.writeFileSync(filePath, JSON.stringify(encrypted, null, 2), { 
    mode: SECURE_FILE_MODE,
    encoding: 'utf8',
  });
}

/**
 * Reads and decrypts data. Returns null if missing or undecryptable
 * (different machine, corrupted).
 */
function readSecure<T>(filePath: string): T | null {
  if (!fs.existsSync(filePath)) return null;
  try {
    return JSON.parse(decrypt(JSON.parse(fs.readFileSync(filePath, 'utf8')))) as T;
  } catch (error) {
    log.error('session-store', `Failed to read ${filePath}: ${error instanceof Error ? error.message : error}`);
    return null;
  }
}

/**
 * Checks if session state file exists.
 */
export function hasSessionState(): boolean {
  return fs.existsSync(SESSION_STATE_PATH);
}

/**
 * Reads the session state.
 */
export function readSessionState(): SessionState | null {
  const state = readSecure<SessionState>(SESSION_STATE_PATH);
  return state ? decodeMsalCache(state) : null;
}

export function isTeamsHostname(hostname: string): boolean {
  return TEAMS_HOSTS.some(host => hostname === host || hostname.endsWith(`.${host}`));
}

/**
 * Keeps only Teams cookies and Teams origins. A browser attached over CDP is
 * the user's own profile, so its storage state also carries every other
 * site's cookies and local storage; nothing outside Teams is read from the
 * session, so nothing outside Teams is persisted.
 */
export function scopeSessionStateToTeams(state: SessionState): SessionState {
  return {
    cookies: state.cookies.filter(cookie => isTeamsHostname((cookie.domain ?? '').replace(/^\./, ''))),
    origins: state.origins.filter(origin => isTeamsHostname(new URL(origin.origin).hostname)),
  };
}

/**
 * Writes the session state securely, limited to Teams.
 */
export function writeSessionState(state: SessionState): void {
  ensureConfigDir();
  writeSecure(SESSION_STATE_PATH, scopeSessionStateToTeams(state));
}

/**
 * Deletes the session state file.
 */
export function clearSessionState(): void {
  if (fs.existsSync(SESSION_STATE_PATH)) {
    fs.unlinkSync(SESSION_STATE_PATH);
  }
}

/**
 * Deletes everything stored locally: browser profile (cookies, MSAL caches,
 * every account ever signed in), session state and token cache.
 */
export function clearAllLocalData(): void {
  fs.rmSync(CONFIG_DIR, { recursive: true, force: true });
}

/**
 * Gets the age of the session state in hours.
 */
export function getSessionAge(): number | null {
  if (!hasSessionState()) {
    return null;
  }

  const stats = fs.statSync(SESSION_STATE_PATH);
  const ageMs = Date.now() - stats.mtimeMs;
  return ageMs / (1000 * 60 * 60);
}

/**
 * Checks if session is likely expired (>12 hours old).
 */
export function isSessionLikelyExpired(): boolean {
  const age = getSessionAge();
  if (age === null) return true;
  return age > SESSION_EXPIRY_HOURS;
}

/**
 * Reads the token cache.
 */
export function readTokenCache(): TokenCache | null {
  return readSecure<TokenCache>(TOKEN_CACHE_PATH);
}

/**
 * Writes the token cache securely.
 */
export function writeTokenCache(cache: TokenCache): void {
  ensureConfigDir();
  writeSecure(TOKEN_CACHE_PATH, cache);
}

/**
 * Clears the token cache.
 */
export function clearTokenCache(): void {
  if (fs.existsSync(TOKEN_CACHE_PATH)) {
    fs.unlinkSync(TOKEN_CACHE_PATH);
  }
}

/**
 * Known Teams origins (commercial and government clouds).
 * Used to find the correct origin in session state.
 */
const TEAMS_ORIGINS = TEAMS_HOSTS.map(host => `https://${host}`);

const MsalEncryptionCookieSchema = z.object({ id: z.string(), key: z.string() });
const MsalEncryptedEntrySchema = z.object({
  id: z.string(),
  nonce: z.string(),
  data: z.string(),
});

/**
 * Decodes MSAL's encrypted cache in a detached session snapshot, never in the browser.
 * Wire format: msal-browser-v5.6.3 LocalStorage.getContext and BrowserCrypto.decrypt.
 */
export function decodeMsalCache(state: SessionState): SessionState {
  for (const origin of state.origins) {
    if (!TEAMS_ORIGINS.includes(origin.origin)) continue;
    const hostname = new URL(origin.origin).hostname;
    const cookie = state.cookies.find(candidate =>
      candidate.name === 'msal.cache.encryption'
      && candidate.domain?.replace(/^\./, '') === hostname
      && (candidate.path === '/' || candidate.path === undefined));
    if (!cookie) continue;

    try {
      const key = MsalEncryptionCookieSchema.parse(JSON.parse(decodeURIComponent(cookie.value)));
      const baseKey = Buffer.from(key.key, 'base64url');
      if (baseKey.length !== 32) continue;
      const clientIds = origin.localStorage.flatMap(entry =>
        entry.name.match(/^msal(?:\.\d+)?\.token\.keys\.(.+)$/)?.slice(1) ?? []);

      for (const entry of origin.localStorage) {
        if (!entry.name.startsWith('msal.')) continue;
        try {
          const encrypted = MsalEncryptedEntrySchema.safeParse(JSON.parse(entry.value));
          if (!encrypted.success || encrypted.data.id !== key.id) continue;
          const nonce = Buffer.from(encrypted.data.nonce, 'base64url');
          if (nonce.length !== 16) continue;
          const context = clientIds.find(clientId => entry.name.includes(clientId)) ?? '';
          const derivedKey = hkdfSync('sha256', baseKey, nonce, context, 32);
          const data = Buffer.from(encrypted.data.data, 'base64url');
          // MSAL derives a new key per nonce, so AES-GCM uses a zero IV.
          const decipher = createDecipheriv('aes-256-gcm', Buffer.from(derivedKey), Buffer.alloc(12));
          decipher.setAuthTag(data.subarray(-16));
          const plaintext = Buffer.concat([decipher.update(data.subarray(0, -16)), decipher.final()]).toString('utf8');
          JSON.parse(plaintext);
          entry.value = plaintext;
        } catch {
          // Leave corrupt entries unreadable rather than treating them as credentials.
        }
      }
    } catch {
      // Missing or rotated keys cannot decrypt this browser session's cache.
    }
  }
  return state;
}

/**
 * Classifies the best SubstrateSearch MSAL token in localStorage.
 * Expired tokens remain useful because their MSAL metadata can drive refresh.
 */
function getSubstrateSearchTokenState(
  localStorage: SessionState['origins'][number]['localStorage'] | undefined
): 'fresh' | 'expired' | null {
  let hasExpiredToken = false;

  for (const item of localStorage ?? []) {
    try {
      const entry: unknown = JSON.parse(item.value);
      if (typeof entry !== 'object' || entry === null) continue;
      if (!('target' in entry) || typeof entry.target !== 'string') continue;
      if (!entry.target.includes('SubstrateSearch')) continue;
      if (!('secret' in entry) || typeof entry.secret !== 'string') continue;

      const payloadSegment = entry.secret.split('.')[1];
      if (!payloadSegment) continue;

      const payload: unknown = JSON.parse(Buffer.from(payloadSegment, 'base64url').toString());
      if (
        typeof payload !== 'object' ||
        payload === null ||
        !('exp' in payload) ||
        typeof payload.exp !== 'number'
      ) {
        continue;
      }

      if (payload.exp * 1000 > Date.now()) return 'fresh';
      hasExpiredToken = true;
    } catch {
      // ignore non-JSON localStorage values and malformed JWTs
    }
  }

  return hasExpiredToken ? 'expired' : null;
}

/**
 * Selects a trusted Teams origin with a fresh SubstrateSearch token.
 */
export function selectLiveTeamsOrigin(state: SessionState): string | null {
  for (const knownOrigin of TEAMS_ORIGINS) {
    const origin = state.origins.find(candidate => candidate.origin === knownOrigin);
    if (getSubstrateSearchTokenState(origin?.localStorage) === 'fresh') return knownOrigin;
  }

  return null;
}

/**
 * Gets a trusted, live Teams origin from the persisted session.
 */
export function getLiveTeamsOrigin(): string | null {
  try {
    const state = readSessionState();
    return state ? selectLiveTeamsOrigin(state) : null;
  } catch {
    return null;
  }
}

/**
 * Gets the Teams origin from session state.
 * Checks multiple known Teams domains to support government clouds.
 *
 * Prefers an origin that has a SubstrateSearch token so Outlook/email and
 * Substrate search keep working when both teams.cloud.microsoft and
 * teams.microsoft.com are present in the saved session (common after the
 * New Teams host migration).
 */
export function getTeamsOrigin(state: SessionState): SessionState['origins'][number] | null {
  if (!state.origins) return null;

  // Prefer fresh credentials only from exact, trusted Teams origins. If none
  // are fresh, retain an expired token's MSAL metadata for HTTP refresh.
  let expiredOrigin: SessionState['origins'][number] | null = null;
  for (const knownOrigin of TEAMS_ORIGINS) {
    const origin = state.origins.find(o => o.origin === knownOrigin);
    if (!origin) continue;

    const tokenState = getSubstrateSearchTokenState(origin.localStorage);
    if (tokenState === 'fresh') return origin;
    if (tokenState === 'expired' && !expiredOrigin) expiredOrigin = origin;
  }
  if (expiredOrigin) return expiredOrigin;

  // Try known Teams origins in priority order
  for (const knownOrigin of TEAMS_ORIGINS) {
    const origin = state.origins.find(o => o.origin === knownOrigin);
    if (origin) return origin;
  }

  // Fallback: any origin on a known Teams host
  return state.origins.find(o => TEAMS_HOSTS.some(host => o.origin.includes(host))) ?? null;
}
