import { composeMiddleware, connectHttp, McpServer, registerPrompts, registerResources, registerTools } from '@mcp-z/server';
import cors from 'cors';
import express from 'express';
import type { RuntimeOverrides, ServerConfig } from '../types.ts';
import { createDefaultRuntime } from './runtime.ts';

export async function createHTTPServer(config: ServerConfig, overrides?: RuntimeOverrides) {
  const runtime = await createDefaultRuntime(config, overrides);
  const modules = runtime.createDomainModules();
  const layers = runtime.middlewareFactories.map((factory) => factory(runtime.deps));
  const composed = composeMiddleware(modules, layers);
  const logger = runtime.deps.logger;
  const port = config.transport.port;
  if (!port) throw new Error('Port is required for HTTP transport');

  const tools = [...composed.tools, ...runtime.deps.oauthAdapters.accountTools];
  const prompts = [...composed.prompts, ...runtime.deps.oauthAdapters.accountPrompts];

  // Built per request (HTTP) and per connection (stdio), not shared. The SDK caches the
  // negotiated protocol revision on the McpServer instance - its own docs say a negotiated
  // session never re-routes a method onto the other era - so one shared instance pins itself
  // to whichever revision reaches it first and answers the other with -32601.
  const buildServer = () => {
    const mcpServer = new McpServer({ name: config.name, version: config.version });
    registerTools(mcpServer, tools);
    registerResources(mcpServer, composed.resources);
    registerPrompts(mcpServer, prompts);
    return mcpServer;
  };

  const app = express();
  app.use(express.json({ limit: '10mb' }));

  // Must mount '/mcp' before any permissive app-level cors(), or that layer answers its
  // preflight first; pass baseUrl's origin/host below or a public deployment 403s itself.
  const publicUrl = config.baseUrl ? new URL(config.baseUrl) : undefined;
  logger.info(`Starting ${config.name} MCP server (http)`);
  const { close, httpServer } = await connectHttp(buildServer, {
    logger,
    app,
    port,
    allowedOrigins: publicUrl ? [publicUrl.origin] : undefined,
    allowedHosts: publicUrl ? [publicUrl.host] : undefined,
  });

  // The loopback OAuth callback and the DCR discovery/registration endpoints are
  // meant to be reached from a browser on another origin. Keep cors() for them,
  // scoped to their own routes - never in front of '/mcp'.
  if (runtime.deps.oauthAdapters.loopbackRouter) {
    app.use('/', cors(), runtime.deps.oauthAdapters.loopbackRouter);
    logger.info('Mounted loopback OAuth callback router');
  }

  if (runtime.deps.oauthAdapters.dcrRouter) {
    app.use('/', cors(), runtime.deps.oauthAdapters.dcrRouter);
    logger.info('Mounted DCR router with OAuth endpoints');
  }

  logger.info('http transport ready');

  return {
    httpServer,
    logger,
    close: async () => {
      await close();
      await runtime.close();
    },
  };
}
