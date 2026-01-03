import { sanitizeForLoggingFormatter } from '@mcp-z/oauth';
import type { CachedToken } from '@mcp-z/oauth-google';
import type { Logger, MiddlewareLayer } from '@mcp-z/server';
import { createLoggingMiddleware } from '@mcp-z/server';
import * as fs from 'fs';
import * as path from 'path';
import pino from 'pino';
import createStore from '../lib/create-store.ts';
import * as mcp from '../mcp/index.ts';
import type { CommonRuntime, RuntimeDeps, RuntimeOverrides, ServerConfig } from '../types.ts';
import { createOAuthAdapters, type OAuthAdapters } from './oauth-google.ts';

export function createLogger(config: ServerConfig): Logger {
  const hasStdio = config.transport.type === 'stdio';
  const logsPath = path.join(config.baseDir, 'logs', `${config.name}.log`);
  if (hasStdio) fs.mkdirSync(path.dirname(logsPath), { recursive: true });
  return pino({ level: config.logLevel ?? 'info', formatters: sanitizeForLoggingFormatter() }, hasStdio ? pino.destination({ dest: logsPath, sync: false }) : pino.destination(1));
}

export async function createTokenStore(baseDir: string) {
  const storeUri = process.env.STORE_URI || `file://${path.join(baseDir, 'tokens.json')}`;
  return createStore<CachedToken>(storeUri);
}

export async function createDcrStore(baseDir: string, required: boolean) {
  if (!required) return undefined;
  const dcrStoreUri = process.env.DCR_STORE_URI || `file://${path.join(baseDir, 'dcr.json')}`;
  return createStore<unknown>(dcrStoreUri);
}

export function createAuthLayer(authMiddleware: OAuthAdapters['middleware']): MiddlewareLayer {
  return {
    withTool: authMiddleware.withToolAuth,
    withResource: authMiddleware.withResourceAuth,
    withPrompt: authMiddleware.withPromptAuth,
  };
}

export function createLoggingLayer(logger: Logger): MiddlewareLayer {
  const logging = createLoggingMiddleware({ logger });
  return {
    withTool: logging.withToolLogging,
    withResource: logging.withResourceLogging,
    withPrompt: logging.withPromptLogging,
  };
}

export async function createDefaultRuntime(config: ServerConfig, overrides?: RuntimeOverrides): Promise<CommonRuntime> {
  if (config.auth === 'dcr' && config.transport.type !== 'http') throw new Error('DCR mode requires an HTTP transport');

  const logger = createLogger(config);
  const tokenStore = await createTokenStore(config.baseDir);
  const baseUrl = config.baseUrl ?? (config.transport.type === 'http' && config.transport.port ? `http://localhost:${config.transport.port}` : undefined);
  const dcrStore = await createDcrStore(config.baseDir, config.auth === 'dcr');
  const oauthAdapters = await createOAuthAdapters(config, { logger, tokenStore, dcrStore }, baseUrl);
  const deps: RuntimeDeps = { config, logger, tokenStore, oauthAdapters, baseUrl };
  const createDomainModules =
    overrides?.createDomainModules ??
    (() => ({
      tools: Object.values(mcp.toolFactories).map((factory) => factory()),
      resources: Object.values(mcp.resourceFactories).map((factory) => factory()),
      prompts: Object.values(mcp.promptFactories).map((factory) => factory()),
    }));
  const middlewareFactories = overrides?.middlewareFactories ?? [() => createAuthLayer(oauthAdapters.middleware), () => createLoggingLayer(logger)];

  return {
    deps,
    middlewareFactories,
    createDomainModules,
    close: async () => {},
  };
}
