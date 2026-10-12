import { execFile } from 'node:child_process';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { expect, it } from 'vitest';

it('executes the suggested login command outside the repository with the same CDP endpoint', async () => {
  const home = await mkdtemp(join(tmpdir(), 'teams-login-proof-'));
  try {
    const result = await new Promise<{ stdout: string; code: string | number }>(resolve => {
      execFile(process.execPath, ['--import', 'tsx', 'src/cli.ts', 'teams_search', '--query', 'test'], {
        env: { ...process.env, HOME: home, USERPROFILE: home, APPDATA: home,
          LOG_LEVEL: 'warn', TEAMS_MCP_CDP_URL: 'http://127.0.0.1:0' },
        timeout: 15_000,
      }, (error, stdout) => resolve({ stdout, code: error?.code ?? 0 }));
    });
    expect(result.code).toBe(1);
    const command = result.stdout.match(/Sign in: (.+)/)?.[1];
    expect(command).toBeDefined();
    if (!command) throw new Error('Missing recovery command');
    const login = await new Promise<{ stdout: string; code: string | number }>(resolve => {
      execFile('/bin/sh', ['-c', command], {
        cwd: tmpdir(),
        env: { ...process.env, HOME: home, USERPROFILE: home, APPDATA: home,
          LOG_LEVEL: 'warn', TEAMS_MCP_CDP_URL: '' },
        timeout: 15_000,
      }, (error, stdout) => resolve({ stdout, code: error?.code ?? 0 }));
    });
    expect(login.code).toBe(1);
    expect(login.stdout).toContain('Calling: teams_login');
    expect(login.stdout).toContain('Arguments: {"forceNew":true}');
    expect(login.stdout).toContain('127.0.0.1:0');
    expect(login.stdout).not.toContain('Could not launch');
    expect(result.stdout).not.toContain('INVALID_INPUT');
    expect(result.stdout).not.toContain('Done.');
  } finally {
    await rm(home, { recursive: true, force: true });
  }
}, 20_000);
