import '../lib/env-loader.ts';
import type { ManagedClient } from '@mcp-z/client';
import assert from 'assert';
import { createStdioFixture } from '../lib/create-stdio-fixture.ts';

describe('Sheets Server Spawn Integration', () => {
  let client: ManagedClient;
  let fixture: ReturnType<typeof createStdioFixture>;

  before(async () => {
    fixture = createStdioFixture('local');
    client = await fixture.connect();
  });

  after(async () => {
    if (fixture) await fixture.close();
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
