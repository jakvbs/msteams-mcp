import { execFile } from 'node:child_process';
import { mkdtemp, rm } from 'node:fs/promises';
import { createServer } from 'node:http';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { expect, it } from 'vitest';

it.each([
  { json: true },
  { json: false },
])('reports failed recovery once through the real CLI (json=$json)', async ({ json }) => {
  const home = await mkdtemp(join(tmpdir(), 'teams-auth-proof-'));
  let discoveryRequests = 0;
  const endpoint = createServer((request, response) => {
    if (request.url?.startsWith('/json/version')) discoveryRequests++;
    response.writeHead(503);
    response.end('Browser unavailable for regression proof');
  });
  await new Promise<void>(resolve => endpoint.listen(0, '127.0.0.1', resolve));
  try {
    const address = endpoint.address();
    if (!address || typeof address === 'string') throw new Error('Missing test endpoint');
    const { stdout, stderr, code } = await new Promise<{ stdout: string; stderr: string; code: string | number }>(resolve => {
      execFile(process.execPath,
        ['--import', 'tsx', join(process.cwd(), 'src/cli.ts'), 'teams_search', '--query', 'test', ...(json ? ['--json'] : [])], {
          env: { ...process.env, HOME: home, USERPROFILE: home, APPDATA: home, LOG_LEVEL: 'warn',
            TEAMS_MCP_CDP_URL: `http://127.0.0.1:${address.port}` },
          timeout: 20_000,
        }, (error, stdout, stderr) => resolve({ stdout, stderr, code: error?.code ?? 0 }));
    });
    expect.soft(code).toBe(1);
    if (json) {
      expect.soft(JSON.parse(stdout)).toMatchObject({ success: false, errorCode: 'BROWSER_ERROR' });
    } else {
      expect.soft(stderr).toContain('code: BROWSER_ERROR');
    }
    expect.soft(json ? stdout : stderr).toContain('503');
    expect.soft(discoveryRequests).toBe(1);
    expect(stderr).toContain('Token refresh via browser failed:');
  } finally {
    await new Promise<void>((resolve, reject) => endpoint.close(error => error ? reject(error) : resolve()));
    await rm(home, { recursive: true, force: true });
  }
}, 25_000);
