/**
 * Sheets Server Spawn Integration Test
 *
 */

import { createServerRegistry, type ManagedClient, type ServerRegistry } from '@mcp-z/client';
import assert from 'assert';

describe('Sheets Server Spawn Integration', () => {
  let client: ManagedClient;
  let registry: ServerRegistry;

  before(async () => {
    registry = createServerRegistry(
      {
        sheets: {
          command: 'node',
          args: ['bin/server.js', '--headless'],
          env: {
            NODE_ENV: 'test',
            GOOGLE_CLIENT_ID: process.env.GOOGLE_CLIENT_ID || '',
            GOOGLE_CLIENT_SECRET: process.env.GOOGLE_CLIENT_SECRET || '',
            HEADLESS: 'true',
            LOG_LEVEL: 'error',
          },
        },
      },
      { cwd: process.cwd() }
    );

    client = await registry.connect('sheets');
  });

  after(async () => {
    if (client) {
      await client.close();
    }

    if (registry) {
      await registry.close();
    }
  });

  it('should connect to Sheets server', async () => {
    // Client is already connected via registry.connect() in before hook
    assert.ok(client, 'Should have connected Sheets client');
  });

  it('should list tools via MCP protocol', async () => {
    const result = await client.listTools();

    assert.ok(result.tools, 'Should return tools');
    assert.ok(result.tools.length > 0, 'Should have at least one tool');

    // Verify specific tools exist
    const includes = (name: string) => result.tools.some((t) => t.name.includes(name));
    assert.ok(includes('values-search'), 'Should have values-search tool');
    assert.ok(includes('spreadsheet-find'), 'Should have spreadsheet-find tool');
    assert.ok(includes('rows-get'), 'Should have rows-get tool');
    assert.ok(includes('rows-append'), 'Should have rows-append tool');
  });
});
