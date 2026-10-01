import { execFile } from 'node:child_process';
import { createCipheriv, hkdfSync } from 'node:crypto';
import { mkdtemp, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { chromium, type BrowserContext, type Page } from 'playwright';
import { expect, it } from 'vitest';

// Microsoft's sign-in page after AADSTS50196: it reports the code and keeps reloading.
const loopingSignInPage = '<html><body><script>window.$Config = { pgid: "ConvergedSignIn", sErrorCode: "50196" };'
  + 'setTimeout(() => location.reload(), 300);</script>Trying to sign you in</body></html>';
// Teams starts sign-in by script, like MSAL; a routed 302 target would bypass the route.
const teamsStartingSignIn = '<html><body><script>'
  + 'location.href = "https://login.microsoftonline.com/common/reprocess";</script></body></html>';

async function addEncryptedToken(page: Page): Promise<void> {
  const clientId = 'fixture-client';
  const cacheKey = `msal.2.fixture-accesstoken-${clientId}-substrate`;
  const key = Buffer.alloc(32, 1);
  const nonce = Buffer.alloc(16, 2);
  const payload = Buffer.from(JSON.stringify({ exp: Math.floor(Date.now() / 1000) + 3600 })).toString('base64url');
  const token = JSON.stringify({ target: 'https://substrate.office.com/search/SubstrateSearch',
    secret: `eyJhbGciOiJub25lIn0.${payload}.fixture` });
  const cipher = createCipheriv('aes-256-gcm', Buffer.from(hkdfSync('sha256', key, nonce, clientId, 32)), Buffer.alloc(12));
  const data = Buffer.concat([cipher.update(token), cipher.final(), cipher.getAuthTag()]).toString('base64url');
  const entry = JSON.stringify({ id: 'fixture-key', nonce: nonce.toString('base64url'), data });
  await page.context().addCookies([{
    name: 'msal.cache.encryption',
    value: encodeURIComponent(JSON.stringify({ id: 'fixture-key', key: key.toString('base64url') })),
    url: 'https://teams.cloud.microsoft', secure: true, sameSite: 'Lax',
  }]);
  await page.evaluate(({ cacheKey, entry, clientId }) => {
    localStorage.setItem(`msal.2.token.keys.${clientId}`, JSON.stringify({ accessToken: [cacheKey] }));
    localStorage.setItem(cacheKey, entry);
  }, { cacheKey, entry, clientId });
}

it.each([
  { scenario: 'signed-out', command: 'tool', code: 'AUTH_REQUIRED', message: 'Teams sign-in is required' },
  { scenario: 'closed-tab', command: 'tool', code: 'BROWSER_ERROR', message: 'tab or browser was closed' },
  // This subprocess/CDP boundary deliberately exercises the real token deadline;
  // a parent fake clock cannot control the separate CLI process's polling clock.
  { scenario: 'missing-token', command: 'tool', code: 'AUTH_REQUIRED', message: 'Teams did not provide a valid authentication token' },
  { scenario: 'late-login', command: 'tool', code: 'AUTH_REQUIRED', message: 'Teams sign-in is required' },
  { scenario: 'sign-in-loop', command: 'tool', code: 'AUTH_REQUIRED', message: 'AADSTS50196' },
  // Explicit login used to wait five minutes on the looping page.
  { scenario: 'sign-in-loop', command: 'login', code: 'AUTH_REQUIRED', message: 'AADSTS50196' },
  { scenario: 'interactive-missing-token', command: 'login', code: 'AUTH_REQUIRED', message: 'Teams did not provide a valid authentication token' },
  { scenario: 'encrypted-token', command: 'login', code: null, message: 'Token refresh detected' },
])('reports $scenario from $command through the real browser and CLI', async ({ scenario, command, code, message }) => {
  const home = await mkdtemp(join(tmpdir(), 'teams-browser-proof-'));
  const profile = join(home, 'browser');
  let context: BrowserContext | undefined;
  try {
    context = await chromium.launchPersistentContext(profile, {
      headless: true, args: ['--remote-debugging-port=0'],
    });
    const port = (await readFile(join(profile, 'DevToolsActivePort'), 'utf8')).split('\n')[0];
    let teamsNavigations = 0;
    await context.route('**/*', async route => {
      const url = new URL(route.request().url());
      if (scenario === 'signed-out' && url.hostname === 'teams.cloud.microsoft') {
        await route.fulfill({ status: 302, headers: { location: 'https://login.microsoftonline.com/test' } });
      } else if (scenario === 'sign-in-loop') {
        const body = url.hostname === 'teams.cloud.microsoft' ? teamsStartingSignIn : loopingSignInPage;
        await route.fulfill({ status: 200, contentType: 'text/html', body });
      } else if (scenario === 'interactive-missing-token' && url.hostname === 'teams.cloud.microsoft'
        && route.request().isNavigationRequest() && teamsNavigations++ === 0) {
        // Fail headless promptly, then let visible login reach Teams without tokens.
        await route.fulfill({ status: 200, contentType: 'text/html', body: teamsStartingSignIn });
      } else {
        await route.fulfill({ status: 200, contentType: 'text/html', body: '<html><body>Sign in</body></html>' });
      }
    });
    if (scenario === 'closed-tab') {
      context.on('page', page => {
        page.once('domcontentloaded', () => { void page.close(); });
      });
    }
    let browserActionError: unknown;
    const result = await new Promise<{ stdout: string; stderr: string; exitCode: string | number }>(resolve => {
      const child = execFile(process.execPath, ['--import', 'tsx', 'src/cli.ts',
        ...(command === 'login' ? ['login', '--force', 'true'] : ['teams_set_out_of_office', '--enabled', 'true'])], {
          env: { ...process.env, HOME: home, USERPROFILE: home, APPDATA: home,
            LOG_LEVEL: scenario === 'late-login' || scenario === 'encrypted-token' ? 'debug' : 'warn',
            TEAMS_MCP_CDP_URL: `http://127.0.0.1:${port}` },
          timeout: 35_000,
        }, (error, stdout, stderr) => resolve({ stdout, stderr, exitCode: error?.code ?? 0 }));
      if (scenario === 'late-login' || scenario === 'encrypted-token') {
        child.stdout?.on('data', (chunk: string) => {
          if (chunk.includes('Waiting for MSAL to refresh tokens...')) {
            const page = context?.pages().find(candidate => candidate.url().startsWith('https://teams.cloud.microsoft'));
            if (!page) return;
            // Inject only after the CLI saved its initial snapshot and began polling.
            const action = scenario === 'late-login' ? page.goto('https://login.microsoftonline.com/test') : addEncryptedToken(page);
            void action.catch(error => { browserActionError = error; child.kill(); });
          }
        });
      }
    });
    expect(browserActionError).toBeUndefined();
    expect.soft(result.exitCode).toBe(code === null ? 0 : 1);
    expect.soft(result.stdout).toContain(message);
    if (code === null) {
      expect.soft(result.stdout).toContain('Login completed silently via SSO');
      expect.soft(result.stdout).toContain('Done.');
    } else {
      expect.soft(result.stdout).toContain(code);
      expect.soft(result.stdout).toContain('login --force true');
      expect.soft(result.stdout).toContain(`TEAMS_MCP_CDP_URL='http://127.0.0.1:${port}'`);
      expect.soft(result.stdout).not.toContain('Done.');
      expect.soft(result.stdout).not.toContain('Complete');
    }
    expect(context.pages()).toHaveLength(1);
  } finally {
    try {
      await context?.close();
    } finally {
      await rm(home, { recursive: true, force: true });
    }
  }
}, 40_000);
