import { execFile } from 'node:child_process';
import { mkdtemp, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { chromium, type BrowserContext } from 'playwright';
import { expect, it } from 'vitest';

it.each([
  { scenario: 'signed-out', code: 'AUTH_REQUIRED', message: 'Teams sign-in is required' },
  { scenario: 'closed-tab', code: 'BROWSER_ERROR', message: 'tab or browser was closed' },
  // This subprocess/CDP boundary deliberately exercises the real token deadline;
  // a parent fake clock cannot control the separate CLI process's polling clock.
  { scenario: 'missing-token', code: 'AUTH_REQUIRED', message: 'Teams did not provide a valid authentication token' },
  { scenario: 'late-login', code: 'AUTH_REQUIRED', message: 'Teams sign-in is required' },
])('reports $scenario through the real browser and CLI', async ({ scenario, code, message }) => {
  const home = await mkdtemp(join(tmpdir(), 'teams-browser-proof-'));
  const profile = join(home, 'browser');
  let context: BrowserContext | undefined;
  try {
    context = await chromium.launchPersistentContext(profile, {
      headless: true, args: ['--remote-debugging-port=0'],
    });
    const port = (await readFile(join(profile, 'DevToolsActivePort'), 'utf8')).split('\n')[0];
    await context.route('**/*', async route => {
      const url = new URL(route.request().url());
      if (scenario === 'signed-out' && url.hostname === 'teams.cloud.microsoft') {
        await route.fulfill({ status: 302, headers: { location: 'https://login.microsoftonline.com/test' } });
      } else {
        await route.fulfill({ status: 200, contentType: 'text/html', body: '<html><body>Sign in</body></html>' });
      }
    });
    if (scenario === 'closed-tab') {
      context.on('page', page => {
        page.once('domcontentloaded', () => { void page.close(); });
      });
    }
    const result = await new Promise<{ stdout: string; stderr: string; exitCode: string | number }>(resolve => {
      const child = execFile(process.execPath,
        ['--import', 'tsx', 'src/cli.ts', 'teams_set_out_of_office', '--enabled', 'true'], {
          env: { ...process.env, HOME: home, USERPROFILE: home, APPDATA: home,
            LOG_LEVEL: scenario === 'late-login' ? 'debug' : 'warn',
            TEAMS_MCP_CDP_URL: `http://127.0.0.1:${port}` },
          timeout: 35_000,
        }, (error, stdout, stderr) => resolve({ stdout, stderr, exitCode: error?.code ?? 0 }));
      if (scenario === 'late-login') {
        child.stdout?.on('data', (chunk: string) => {
          if (chunk.includes('Waiting for MSAL to refresh tokens...')) {
            const page = context?.pages().find(candidate => candidate.url().startsWith('https://teams.cloud.microsoft'));
            void page?.goto('https://login.microsoftonline.com/test');
          }
        });
      }
    });
    expect.soft(result.exitCode).toBe(1);
    expect.soft(result.stdout).toContain(code);
    expect.soft(result.stdout).toContain(message);
    expect.soft(result.stdout).toContain('login --force true');
    expect.soft(result.stdout).toContain(`TEAMS_MCP_CDP_URL='http://127.0.0.1:${port}'`);
    expect.soft(result.stdout).not.toContain('Done.');
    expect.soft(result.stdout).not.toContain('Complete');
    expect(context.pages()).toHaveLength(1);
  } finally {
    try {
      await context?.close();
    } finally {
      await rm(home, { recursive: true, force: true });
    }
  }
}, 40_000);
