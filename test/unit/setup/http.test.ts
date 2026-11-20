import '../../lib/env-loader.ts';
import { pathToFileURL } from 'node:url';
import type { ServerConfig } from '@mcp-z/mcp-sheets';
import { setup } from '@mcp-z/mcp-sheets';
import { dcrUtils } from '@mcp-z/oauth';
import { Client as ModernClient, StreamableHTTPClientTransport as ModernTransport, type OAuthClientProvider, UnauthorizedError } from '@modelcontextprotocol/client';
import assert from 'assert';
import { randomUUID } from 'crypto';
import * as fs from 'fs';
import { safeRmSync } from 'fs-remove-compat';
import getPort from 'get-port';
import * as path from 'path';
import createStore from '../../../src/lib/create-store.ts';

describe('setup.createHTTPServer - transport initialization', () => {
  // Note: stdio transport tests are skipped because stdio initialization blocks waiting for input.
  // The stdio transport is tested indirectly through integration tests and manual CLI testing.

  const servers: Awaited<ReturnType<typeof setup.createHTTPServer>>[] = [];
  let testContextPath: string;

  before(async () => {
    // Create isolated test context with pre-configured account
    const testId = randomUUID();
    testContextPath = path.join(process.cwd(), '.tmp', `.mcp-z-test-${testId}`);
    fs.mkdirSync(testContextPath, { recursive: true });
  });

  after(async () => {
    // Use close function to properly shut down all transports
    for (const result of servers) {
      await result.close();
    }

    // Clean up isolated test context
    if (testContextPath && fs.existsSync(testContextPath)) {
      safeRmSync(testContextPath, { recursive: true, force: true });
    }
  });

  it('initializes single HTTP transport with OAuth', async () => {
    const port = await getPort({ host: '127.0.0.1' });
    const config: ServerConfig = {
      name: 'test-server',
      version: '0.0.0-test',
      transport: {
        type: 'http',
        port,
      },
      baseDir: testContextPath,
      clientId: 'test-client-id',
      clientSecret: 'test-client-secret',
      headless: true,
      logLevel: 'error',
      auth: 'loopback-oauth',
      repositoryUrl: 'https://github.com/mcp-z/mcp-sheets',
      resourceStoreUri: pathToFileURL(path.join(testContextPath, 'files')).href,
    };

    const result = await setup.createHTTPServer(config);
    servers.push(result);

    assert.ok('httpServer' in result && result.httpServer, 'HTTP server should be initialized');
  });

  it('includes logger in server result', async () => {
    const port = await getPort({ host: '127.0.0.1' });
    const config: ServerConfig = {
      name: 'test-server',
      version: '0.0.0-test',
      transport: { type: 'http', port },
      baseDir: testContextPath,
      clientId: 'test-client-id',
      clientSecret: 'test-client-secret',
      headless: true,
      logLevel: 'error',
      auth: 'loopback-oauth',
      repositoryUrl: 'https://github.com/mcp-z/mcp-sheets',
      resourceStoreUri: pathToFileURL(path.join(testContextPath, 'files')).href,
    };

    const result = await setup.createHTTPServer(config);
    servers.push(result);

    assert.ok(result.logger, 'Result should have logger');
    assert.strictEqual(typeof result.logger.info, 'function', 'Logger should have info method');
    assert.strictEqual(typeof result.logger.error, 'function', 'Logger should have error method');
  });

  it('creates server with MCP server instance', async () => {
    const port = await getPort({ host: '127.0.0.1' });
    const config: ServerConfig = {
      name: 'test-server',
      version: '0.0.0-test',
      transport: { type: 'http', port },
      baseDir: testContextPath,
      clientId: 'test-client-id',
      clientSecret: 'test-client-secret',
      headless: true,
      logLevel: 'error',
      auth: 'loopback-oauth',
      repositoryUrl: 'https://github.com/mcp-z/mcp-sheets',
      resourceStoreUri: pathToFileURL(path.join(testContextPath, 'files')).href,
    };

    const result = await setup.createHTTPServer(config);
    servers.push(result);

    assert.strictEqual(typeof result.close, 'function', 'Result should have close function');
  });

  describe('DCR bearer gate', () => {
    const MCP_HEADERS = { 'content-type': 'application/json', accept: 'application/json, text/event-stream', 'mcp-protocol-version': '2025-06-18' };
    const initialize = { jsonrpc: '2.0', id: 1, method: 'initialize', params: { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'http-test', version: '0.0.0-test' } } };
    const TOKEN = 'seeded-dcr-token';

    // A DCR store holding one access token whose provider tokens never reach Google.
    async function seedDcrStore(): Promise<string> {
      const dcrFile = pathToFileURL(path.join(testContextPath, `dcr-${randomUUID()}.json`)).href;
      const seed = await createStore<unknown>(dcrFile);
      await dcrUtils.setAccessToken(seed as never, TOKEN, {
        access_token: TOKEN,
        token_type: 'Bearer',
        expires_in: 3600,
        scope: 'openid',
        client_id: 'seeded-client',
        providerTokens: { accessToken: 'never-sent-to-google', expiresAt: Date.now() + 3600_000 },
        created_at: Date.now(),
      } as never);
      await seed.disconnect();
      return dcrFile;
    }

    // A DCR server on loopback whose store holds the seeded token.
    async function startDcrServer(): Promise<{ base: string }> {
      const port = await getPort({ host: '127.0.0.1' });
      const base = `http://127.0.0.1:${port}`;
      const dcrFile = await seedDcrStore();
      const saved = process.env.DCR_STORE_URI;
      process.env.DCR_STORE_URI = dcrFile;
      try {
        const config: ServerConfig = {
          name: 'test-server',
          version: '0.0.0-test',
          transport: { type: 'http', port },
          baseDir: testContextPath,
          baseUrl: base,
          clientId: 'test-client-id',
          clientSecret: 'test-client-secret',
          headless: true,
          logLevel: 'error',
          auth: 'dcr',
          resourceStoreUri: pathToFileURL(path.join(testContextPath, 'files')).href,
          repositoryUrl: 'https://github.com/mcp-z/mcp-sheets',
        };
        servers.push(await setup.createHTTPServer(config));
      } finally {
        if (saved === undefined) delete process.env.DCR_STORE_URI;
        else process.env.DCR_STORE_URI = saved;
      }
      return { base };
    }

    const post = (base: string, body: unknown, authorization?: string) => fetch(`${base}/mcp`, { method: 'POST', headers: { ...MCP_HEADERS, ...(authorization && { authorization }) }, body: JSON.stringify(body) });

    it('answers 401 with a resource_metadata challenge when there is no token', async () => {
      const { base } = await startDcrServer();
      const response = await post(base, initialize);
      assert.strictEqual(response.status, 401);
      const challenge = response.headers.get('www-authenticate') ?? '';
      assert.ok(challenge.includes(`resource_metadata="${base}/.well-known/oauth-protected-resource/mcp"`), challenge);
    });

    it('answers 401 invalid_token for an unknown token', async () => {
      const { base } = await startDcrServer();
      const response = await post(base, initialize, 'Bearer not-a-real-token');
      assert.strictEqual(response.status, 401);
      assert.ok((response.headers.get('www-authenticate') ?? '').includes('error="invalid_token"'));
    });

    for (const [label, mode] of [
      ['a legacy client', 'legacy'],
      ['a 2026-07-28 client', { pin: '2026-07-28' }],
    ] as const) {
      it(`lists tools for ${label} holding a seeded token`, async () => {
        const { base } = await startDcrServer();
        const client = new ModernClient({ name: 'http-test', version: '0.0.0-test' }, { versionNegotiation: { mode } });
        const transport = new ModernTransport(new URL(`${base}/mcp`), { requestInit: { headers: { Authorization: `Bearer ${TOKEN}` } } });
        try {
          await client.connect(transport);
          const { tools } = await client.listTools();
          assert.ok(tools.some((tool) => tool.name === 'values-search'));
        } finally {
          await client.close();
        }
      });
    }

    it('sends an SDK client without a token to the authorization endpoint without opening it', async () => {
      const { base } = await startDcrServer();
      let authorizationUrl: URL | undefined;
      let registered: { client_id: string } | undefined;
      let verifier: string | undefined;
      const redirectUrl = 'http://127.0.0.1:53682/callback';
      const provider: OAuthClientProvider = {
        redirectUrl,
        clientMetadata: { client_name: 'http-test', redirect_uris: [redirectUrl], grant_types: ['authorization_code', 'refresh_token'], response_types: ['code'], token_endpoint_auth_method: 'none' },
        clientInformation: () => registered,
        saveClientInformation: (info) => {
          registered = info as { client_id: string };
        },
        tokens: () => undefined,
        saveTokens: () => {},
        redirectToAuthorization: (url) => {
          authorizationUrl = url;
        },
        saveCodeVerifier: (value) => {
          verifier = value;
        },
        codeVerifier: () => verifier as string,
      };
      const client = new ModernClient({ name: 'http-test', version: '0.0.0-test' });
      try {
        await assert.rejects(client.connect(new ModernTransport(new URL(`${base}/mcp`), { authProvider: provider })), UnauthorizedError);
      } finally {
        await client.close();
      }
      assert.ok(authorizationUrl, 'redirectToAuthorization was reached');
      assert.strictEqual(`${authorizationUrl.origin}${authorizationUrl.pathname}`, `${base}/oauth/authorize`);
    });
  });
});
