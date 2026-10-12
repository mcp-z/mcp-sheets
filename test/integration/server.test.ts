import '../lib/env-loader.ts';
import { type ManagedClient, ToolResponseError } from '@mcp-z/client';
import assert from 'assert';
import { createStdioFixture } from '../lib/create-stdio-fixture.ts';
import { METADATA_SHEET_ID } from '../lib/metadata-fixture.ts';

describe('Sheets MCP Server Component Tests', () => {
  let client: ManagedClient;
  let fixture: ReturnType<typeof createStdioFixture>;

  before(async () => {
    fixture = createStdioFixture('local');
    client = await fixture.connect();
  });

  after(async () => {
    if (fixture) await fixture.close();
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
      await assert.rejects(
        () => client.callTool({ name: 'cells-metadata-get', arguments: { id: 'metadata-schema-test', gid: METADATA_SHEET_ID, ranges: ['A:A'] } }),
        (error: unknown) => {
          assert.ok(error instanceof ToolResponseError);
          assert.equal(error.response.isError, true);
          return true;
        }
      );
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
