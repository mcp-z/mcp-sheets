import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';
import assert from 'assert';
import * as path from 'path';

describe('Sheets MCP Server Component Tests', () => {
  let client: Client;
  let transport: StdioClientTransport;

  before(async () => {
    // Resolve paths relative to server root
    const serverRoot = path.resolve(import.meta.dirname, '../../..');
    const envFile = path.join(serverRoot, '.env.test');
    const serverPath = path.join(serverRoot, 'bin/server.js');

    // StdioClientTransport spawns the server automatically
    transport = new StdioClientTransport({
      command: 'node',
      args: [`--env-file=${envFile}`, serverPath],
      env: {
        ...process.env,
        NODE_ENV: 'test',
      } as Record<string, string>,
    });

    client = new Client({ name: 'test-client', version: '1.0.0' }, { capabilities: {} });

    await client.connect(transport);
  });

  after(async () => {
    await client.close();
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
      const expectedTools = ['rows-append', 'rows-get', 'values-search', 'sheet-create', 'sheet-delete', 'sheet-find', 'spreadsheet-create', 'spreadsheet-find', 'values-batch-update'];

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
