import assert from 'assert';
import spawn from 'cross-spawn-cb';
import * as fs from 'fs';
import { safeRmSync } from 'fs-remove-compat';
import * as path from 'path';

describe('stdio start failure (sheets)', () => {
  const projectDir = path.resolve('.tmp', `runtime-stdio-failure-${process.pid}`);
  after(() => safeRmSync(projectDir, { recursive: true, force: true }));

  it('prints the start error without a second sonic boom error', async () => {
    fs.mkdirSync(projectDir, { recursive: true });
    fs.writeFileSync(path.join(projectDir, '.mcp.json'), '{ "mcpServers": {} }\n');
    const env = { PATH: process.env.PATH ?? '', GOOGLE_CLIENT_ID: 'test-client-id', TOKEN_STORE_URI: 'bogus:///x', LOG_LEVEL: 'error' };
    let stderr: string | undefined;
    try {
      await spawn(process.execPath, [path.resolve('bin/server.js'), '--stdio', '--headless'], { cwd: projectDir, env, encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'] });
    } catch (error) {
      stderr = String((error as { stderr?: unknown }).stderr ?? '');
    }
    assert.notStrictEqual(stderr, undefined, 'the server should not start with an unknown store protocol');
    assert.ok(stderr?.includes('Failed to create store for bogus:'), `stderr: ${stderr}`);
    assert.ok(!stderr?.includes('sonic boom is not ready yet'), `stderr: ${stderr}`);
  });
});
