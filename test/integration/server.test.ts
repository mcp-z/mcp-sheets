// Network: the nested live metadata invocation suite uses Google Sheets and requires configured credentials.
import '../lib/env-loader.ts';
import { pathToFileURL } from 'node:url';
import { mcp } from '@mcp-z/mcp-sheets';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';
import { CallToolResultSchema } from '@modelcontextprotocol/sdk/types.js';
import assert from 'assert';
import * as path from 'path';
import { createMetadataFixture, METADATA_SHEET_ID, METADATA_SHEET_TITLE } from '../lib/metadata-fixture.ts';

describe('Sheets MCP Server Component Tests', () => {
  let client: Client;
  let transport: StdioClientTransport;

  before(async () => {
    // Resolve paths relative to server root
    const serverRoot = path.resolve(import.meta.dirname, '../..');
    const serverPath = path.join(serverRoot, 'bin/server.js');

    // StdioClientTransport spawns the server automatically
    transport = new StdioClientTransport({
      command: process.execPath,
      args: [serverPath, '--stdio', '--headless', '--auth=loopback-oauth'],
      env: {
        ...process.env,
        NODE_ENV: 'test',
        TOKEN_STORE_URI: pathToFileURL(path.join(serverRoot, '.tokens/store.json')).href,
      } as Record<string, string>,
    });

    client = new Client({ name: 'test-client', version: '1.0.0' }, { capabilities: {} });

    await client.connect(transport);
  });

  after(async () => {
    if (client) await client.close();
  });

  describe('MCP Protocol Component Testing', () => {
    it('should respond to MCP tools/list request', async () => {
      const result = await client.listTools();

      assert(Array.isArray(result.tools), 'Should return tools array');
      assert(result.tools.length > 0, 'Should have at least one tool');
    });

    it('should respond to MCP prompts/list request', async () => {
      const result = await client.listPrompts();

      assert(Array.isArray(result.prompts) || result.prompts === undefined, 'Should return prompts array or undefined');
    });

    it('should respond to MCP resources/list request', async () => {
      const result = await client.listResources();

      assert(Array.isArray(result.resources), 'Should return resources array');
    });

    it('should have expected Sheets tools available', async () => {
      const result = await client.listTools();

      const toolNames = result.tools.map((tool) => tool.name);

      // Expected Sheets tools based on servers/mcp-sheets/src/mcp/tools/index.ts
      const expectedTools = ['rows-append', 'rows-get', 'values-search', 'sheet-create', 'sheet-delete', 'sheet-find', 'spreadsheet-create', 'spreadsheet-find', 'values-batch-update', 'spreadsheet-metadata-get', 'cells-metadata-get'];

      // Verify each expected tool is registered
      for (const expectedTool of expectedTools) {
        assert(toolNames.includes(expectedTool), `Should have ${expectedTool} tool registered`);
      }
    });

    it('should have properly configured tool schemas', async () => {
      const result = await client.listTools();

      // Verify each tool has required MCP schema fields
      for (const tool of result.tools) {
        assert(typeof tool.name === 'string', `Tool ${tool.name} should have string name`);
        assert(typeof tool.description === 'string', `Tool ${tool.name} should have string description`);
        assert(typeof tool.inputSchema === 'object', `Tool ${tool.name} should have inputSchema object`);

        // Verify inputSchema is properly structured
        const inputSchema = tool.inputSchema;
        assert.strictEqual(inputSchema.type, 'object', `Tool ${tool.name} inputSchema should be object type`);
        assert(typeof inputSchema.properties === 'object', `Tool ${tool.name} should have properties in inputSchema`);
      }
    });

    it('discovers both metadata contracts with read-only annotations', async () => {
      const { tools } = await client.listTools();
      const workbook = tools.find((tool) => tool.name === 'spreadsheet-metadata-get');
      const cells = tools.find((tool) => tool.name === 'cells-metadata-get');
      assert.ok(workbook && cells);
      assert.equal(workbook.annotations?.readOnlyHint, true);
      assert.equal(cells.annotations?.readOnlyHint, true);
      assert.deepEqual(Object.keys(workbook.inputSchema.properties ?? {}), ['id']);
      assert.deepEqual(cells.inputSchema.required, ['id', 'gid', 'ranges']);
    });

    it('rejects invalid metadata ranges at the registered protocol boundary', async () => {
      const response = await client.callTool({ name: 'cells-metadata-get', arguments: { id: 'metadata-schema-test', gid: METADATA_SHEET_ID, ranges: ['A:A'] } });
      assert.equal(response.isError, true);
    });

    describe('Live metadata invocation', () => {
      let metadataFixture: Awaited<ReturnType<typeof createMetadataFixture>>;

      before(async () => {
        metadataFixture = await createMetadataFixture();
      });

      after(async () => {
        if (metadataFixture) await metadataFixture.close();
      });

      it('invokes both metadata tools through the built package and preserves structured and text results', async () => {
        const workbook = CallToolResultSchema.parse(await client.callTool({ name: 'spreadsheet-metadata-get', arguments: { id: metadataFixture.id } }));
        assert.ok(!workbook.isError);
        const workbookResult = mcp.toolFactories.spreadsheetMetadataGet().config.outputSchema.parse(workbook.structuredContent).result;
        if (workbookResult.type !== 'success') assert.fail('Expected authenticated workbook metadata');
        assert.equal(workbookResult.metadata.spreadsheetId, metadataFixture.id);
        assert.equal(workbookResult.metadata.properties?.locale, 'en_CA');
        assert.equal(workbookResult.metadata.sheets?.length, 2);
        const workbookSheets = workbookResult.metadata.sheets ?? [];
        assert.deepEqual(workbookSheets.map((sheet) => sheet.properties?.title).sort(), ['Empty tab', METADATA_SHEET_TITLE].sort());
        for (const sheet of workbookSheets) assert.equal(sheet.data, undefined);

        const cells = CallToolResultSchema.parse(await client.callTool({ name: 'cells-metadata-get', arguments: { id: metadataFixture.id, gid: METADATA_SHEET_ID, ranges: ['E8'] } }));
        assert.ok(!cells.isError);
        const cellsResult = mcp.toolFactories.cellsMetadataGet().config.outputSchema.parse(cells.structuredContent).result;
        if (cellsResult.type !== 'success') assert.fail('Expected authenticated cell metadata');
        assert.equal(cellsResult.gid, String(METADATA_SHEET_ID));
        const sheet = cellsResult.metadata.sheets?.[0];
        assert.equal(sheet?.properties?.title, METADATA_SHEET_TITLE);
        assert.equal(sheet?.data?.[0]?.startRow, 7);
        assert.equal(sheet?.data?.[0]?.startColumn, 4);
        const note = sheet?.data?.[0]?.rowData?.[0]?.values?.[0];
        assert.equal(note?.note, 'Selected note');
        assert.equal(note?.userEnteredValue, undefined);
        for (const { response, result } of [
          { response: workbook, result: workbookResult },
          { response: cells, result: cellsResult },
        ]) {
          const text = response.content.find((item) => item.type === 'text');
          assert.ok(text && text.type === 'text');
          assert.deepEqual(JSON.parse(text.text), result);
        }
      });
    });
  });

  describe('Component Health and Status', () => {
    it('should be accessible as a single component', async () => {
      // Simple health check - any successful MCP response indicates the server is running
      const result = await client.listTools();

      // Any valid MCP response means the server is accessible
      assert(result.tools, 'Should return tools from MCP server');
    });
  });
});
