import { execFile } from 'node:child_process';
import { mkdtemp, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { chromium, type BrowserContext } from 'playwright';
import { expect, it } from 'vitest';

const jwtPart = (value: object) => Buffer.from(JSON.stringify(value)).toString('base64url');
const token = JSON.stringify({
  target: 'https://substrate.office.com/search/SubstrateSearch',
  secret: [jwtPart({ alg: 'none' }), jwtPart({ exp: Math.floor(Date.now() / 1000) + 3600 }), 'fixture'].join('.'),
});
// Signed-in Teams with MSAL's hidden sign-in iframe, so the context has seen a second origin.
const signedInTeams = '<html><body><iframe src="https://login.microsoftonline.com/silent"></iframe><script>'
  + `localStorage.setItem('msal.fixture-accesstoken', ${JSON.stringify(token)});</script></body></html>`;

it('signs in through an attached browser using only its own tab', async () => {
  const home = await mkdtemp(join(tmpdir(), 'teams-attach-proof-'));
  const profile = join(home, 'browser');
  let context: BrowserContext | undefined;
  try {
    context = await chromium.launchPersistentContext(profile, {
      headless: true, args: ['--remote-debugging-port=0'],
    });
    const port = (await readFile(join(profile, 'DevToolsActivePort'), 'utf8')).split('\n')[0];
    await context.route('**/*', route => route.fulfill({
      status: 200,
      contentType: 'text/html',
      body: new URL(route.request().url()).hostname.startsWith('teams.') ? signedInTeams : '<html></html>',
    }));
    let openedTabs = 0;
    context.on('page', () => { openedTabs++; });

    const result = await new Promise<{ stdout: string; stderr: string; exitCode: string | number }>(resolve => {
      execFile(process.execPath, ['--import', 'tsx', 'src/cli.ts', 'login'], {
        env: { ...process.env, HOME: home, USERPROFILE: home, APPDATA: home, LOG_LEVEL: 'warn',
          TEAMS_MCP_CDP_URL: `http://127.0.0.1:${port}` },
        timeout: 35_000,
      }, (error, stdout, stderr) => resolve({ stdout, stderr, exitCode: error?.code ?? 0 }));
    });

    expect.soft(result.stdout + result.stderr).toContain('Login completed silently via SSO');
    expect.soft(result.exitCode).toBe(0);
    // Playwright's storageState() would open another tab for the iframe's origin.
    expect(openedTabs).toBe(1);
    expect(context.pages()).toHaveLength(1);
  } finally {
    try {
      await context?.close();
    } finally {
      await rm(home, { recursive: true, force: true });
    }
  }
}, 40_000);
