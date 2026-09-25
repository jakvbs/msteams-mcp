import { execFile } from 'node:child_process';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { expect, it } from 'vitest';

it('forwards forceNew from the recommended login command', async () => {
  const home = await mkdtemp(join(tmpdir(), 'teams-login-proof-'));
  try {
    const result = await new Promise<{ stdout: string; code: string | number }>(resolve => {
      execFile(process.execPath, ['--import', 'tsx', 'src/cli.ts', 'login', '--force', 'true'], {
        env: { ...process.env, HOME: home, USERPROFILE: home, APPDATA: home,
          LOG_LEVEL: 'warn', TEAMS_MCP_CDP_URL: 'http://127.0.0.1:0' },
        timeout: 15_000,
      }, (error, stdout) => resolve({ stdout, code: error?.code ?? 0 }));
    });
    expect(result.code).toBe(1);
    expect(result.stdout).toContain('Calling: teams_login');
    expect(result.stdout).toContain('Arguments: {"forceNew":true}');
    expect(result.stdout).not.toContain('INVALID_INPUT');
    expect(result.stdout).not.toContain('Done.');
  } finally {
    await rm(home, { recursive: true, force: true });
  }
}, 20_000);
