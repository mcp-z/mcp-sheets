import { createServerRegistry, type ManagedClient, type ServerRegistry } from '@mcp-z/client';
import assert from 'assert';
import { randomUUID } from 'crypto';
import * as fs from 'fs';
import { safeRmSync } from 'fs-remove-compat';
import * as path from 'path';
import { fileURLToPath, pathToFileURL } from 'url';

export function createStdioFixture(mode: 'local' | 'live') {
  const serverRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');
  const projectDir = path.join(serverRoot, '.tmp', `sheets-stdio-${randomUUID()}`);
  const stateDir = path.join(projectDir, '.mcp-z');
  let registry: ServerRegistry | undefined;
  let spawnAttempted = false;

  async function connect(): Promise<ManagedClient> {
    fs.mkdirSync(projectDir, { recursive: true });
    fs.writeFileSync(path.join(projectDir, '.mcp.json'), '{ "mcpServers": {} }\n');
    const localEnv: Record<string, string> = {};
    for (const key of ['PATH', 'SystemRoot', 'WINDIR', 'ComSpec', 'TEMP', 'TMP', 'TMPDIR']) {
      const value = process.env[key];
      if (value !== undefined) localEnv[key] = value;
    }

    spawnAttempted = true;
    registry = createServerRegistry(
      {
        sheets: {
          command: process.execPath,
          args: [path.join(serverRoot, 'bin/server.js'), '--stdio', '--headless', '--auth=loopback-oauth'],
          env: {
            ...(mode === 'local' ? { GOOGLE_CLIENT_ID: 'test-client-id', GOOGLE_CLIENT_SECRET: '' } : {}),
            NODE_ENV: 'test',
            HEADLESS: 'true',
            LOG_LEVEL: 'error',
            TOKEN_STORE_URI: pathToFileURL(mode === 'live' ? path.join(serverRoot, '.tokens/store.json') : path.join(stateDir, 'tokens.json')).href,
            DCR_STORE_URI: pathToFileURL(path.join(stateDir, 'dcr.json')).href,
            RESOURCE_STORE_URI: pathToFileURL(path.join(stateDir, 'files')).href,
          },
        },
      },
      { cwd: projectDir, ...(mode === 'local' ? { env: localEnv } : {}) }
    );
    return registry.connect('sheets');
  }

  async function close(): Promise<void> {
    const failures: unknown[] = [];
    let childClosed = !spawnAttempted;
    if (registry) {
      try {
        const result = await registry.close();
        childClosed = true;
        assert.deepStrictEqual({ timedOut: result.timedOut, killedCount: result.killedCount }, { timedOut: false, killedCount: 0 }, 'Server should shut down cooperatively');
      } catch (error) {
        failures.push(error);
      }
    }
    if (childClosed) {
      try {
        safeRmSync(projectDir, { recursive: true, force: true });
      } catch (error) {
        failures.push(error);
      }
    } else {
      failures.push(new Error(`Server closure unconfirmed; retained scratch ${projectDir}`));
    }
    if (failures.length === 1) throw failures[0];
    if (failures.length > 1) {
      const error = new Error(`Stdio fixture teardown failed: ${failures.map((failure) => (failure instanceof Error ? failure.message : String(failure))).join('; ')}`);
      throw Object.assign(error, { errors: failures });
    }
  }

  return { projectDir, connect, close };
}
