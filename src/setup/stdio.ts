import { composeMiddleware, connectStdio, registerPrompts, registerResources, registerTools } from '@mcp-z/server';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import type { RuntimeOverrides, ServerConfig } from '../types.ts';
import { createDefaultRuntime } from './runtime.ts';

export async function createStdioServer(config: ServerConfig, overrides?: RuntimeOverrides) {
  const runtime = await createDefaultRuntime(config, overrides);
  const modules = runtime.createDomainModules();
  const layers = runtime.middlewareFactories.map((factory) => factory(runtime.deps));
  const composed = composeMiddleware(modules, layers);
  const logger = runtime.deps.logger;

  const tools = [...composed.tools, ...runtime.deps.oauthAdapters.accountTools];
  const prompts = [...composed.prompts, ...runtime.deps.oauthAdapters.accountPrompts];

  const mcpServer = new McpServer({ name: config.name, version: config.version });
  registerTools(mcpServer, tools);
  registerResources(mcpServer, composed.resources);
  registerPrompts(mcpServer, prompts);

  logger.info(`Starting ${config.name} MCP server (stdio)`);
  const { close } = await connectStdio(mcpServer, { logger });
  logger.info('stdio transport ready');

  return {
    mcpServer,
    logger,
    close: async () => {
      await close();
      await runtime.close();
    },
  };
}
