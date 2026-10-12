import { execFile } from 'node:child_process';
import { mkdtemp, rm, symlink } from 'node:fs/promises';
import { createServer } from 'node:http';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { CallToolRequestSchema, ListToolsRequestSchema } from '@modelcontextprotocol/sdk/types.js';
import { expect, it, vi } from 'vitest';
import { callTool } from './cli.js';

it.each([
  { entry: 'direct', json: true },
  { entry: 'symlink', json: false },
])('reports failed recovery through the real CLI ($entry, json=$json)', async ({ entry, json }) => {
  const home = await mkdtemp(join(tmpdir(), 'teams-auth-proof-'));
  const cli = join(process.cwd(), 'src/cli.ts');
  const linkedCli = join(home, 'msteams');
  if (entry === 'symlink') await symlink(cli, linkedCli);
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
        ['--import', 'tsx', entry === 'symlink' ? linkedCli : cli, 'teams_search', '--query', 'test', ...(json ? ['--json'] : [])], {
          env: { ...process.env, HOME: home, USERPROFILE: home, APPDATA: home, LOG_LEVEL: 'warn',
            TEAMS_MCP_CDP_URL: `http://127.0.0.1:${address.port}` },
          timeout: 20_000,
        }, (error, stdout, stderr) => resolve({ stdout, stderr, code: error?.code ?? 0 }));
    });
    expect.soft(code).toBe(1);
    if (json) expect(JSON.parse(stdout)).toMatchObject({ isError: true });
    expect.soft(stdout).toContain('BROWSER_ERROR');
    expect.soft(stdout).toContain('503');
    expect.soft(stdout).not.toContain('Done.');
    expect.soft(stdout).not.toContain('Complete');
    expect.soft(discoveryRequests).toBe(1);
    expect.soft(stderr).toContain('Token refresh via browser failed:');
    expect(stderr).toContain('503');
  } finally {
    await new Promise<void>((resolve, reject) => endpoint.close(error => error ? reject(error) : resolve()));
    await rm(home, { recursive: true, force: true });
  }
}, 25_000);

it.each([
  { elapsed: 60_001, expected: 'completed' },
  { elapsed: 300_001, expected: 'McpError: MCP error -32001: Request timed out' },
])('enforces the CLI authentication budget at $elapsed ms', async ({ elapsed, expected }) => {
  const client = new Client({ name: 'timeout-proof', version: '1' });
  const server = new Server({ name: 'auth-fixture', version: '1' }, { capabilities: { tools: {} } });
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  let started = () => {};
  const entered = new Promise<void>(resolve => { started = resolve; });
  let finish = () => {};
  const recovered = new Promise<void>(resolve => { finish = resolve; });
  server.setRequestHandler(ListToolsRequestSchema, async () => ({ tools: [
    { name: 'teams_get_out_of_office', inputSchema: { type: 'object' } },
  ] }));
  server.setRequestHandler(CallToolRequestSchema, async () => {
    started();
    await recovered;
    return { content: [{ type: 'text', text: 'authentication recovered' }] };
  });
  await server.connect(serverTransport);
  await client.connect(clientTransport);
  const output = vi.spyOn(console, 'log').mockImplementation(() => {});
  vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout', 'Date'] });
  try {
    const outcome = callTool(client, 'teams_get_out_of_office', {}, true)
      .then(() => 'completed', error => String(error));
    await entered;
    await vi.advanceTimersByTimeAsync(elapsed);
    finish();
    expect(await outcome).toBe(expected);
    if (expected === 'completed') {
      expect(output).toHaveBeenCalledWith(expect.stringContaining('authentication recovered'));
    } else {
      expect(output).not.toHaveBeenCalled();
    }
  } finally {
    finish();
    vi.useRealTimers();
    output.mockRestore();
    await client.close();
    await server.close();
  }
});
