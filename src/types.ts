import type { CachedToken, DcrConfig, OAuthConfig } from '@mcp-z/oauth-google';
import type { BaseServerConfig, MiddlewareLayer, PromptModule, ResourceModule, Logger as ServerLogger, ToolModule } from '@mcp-z/server';
import type { Keyv } from 'keyv';
import type { OAuthAdapters } from './setup/oauth-google.ts';

export type Logger = Pick<Console, 'info' | 'error' | 'warn' | 'debug'>;

/**
 * Composes transport config, OAuth config, and application-level config
 */
export interface ServerConfig extends BaseServerConfig, OAuthConfig {
  logLevel: string;
  baseDir: string;
  name: string;
  version: string;
  repositoryUrl: string;

  // File serving configuration for CSV exports
  resourceStoreUri: string;
  baseUrl?: string;

  // DCR configuration (when auth === 'dcr')
  dcrConfig?: DcrConfig;
}

export interface StorageContext {
  resourceStoreUri: string;
  baseUrl?: string;
  transport: BaseServerConfig['transport'];
}

export interface StorageExtra {
  storageContext: StorageContext;
}

export interface GoogleApiError {
  response?: { status?: number };
  status?: number;
  statusCode?: number;
  code?: number | string;
  message?: string;
}

/** Runtime dependencies exposed to middleware/factories. */
export interface RuntimeDeps {
  config: ServerConfig;
  logger: ServerLogger;
  tokenStore: Keyv<CachedToken>;
  oauthAdapters: OAuthAdapters;
  baseUrl?: string;
}

/** Collections of MCP modules produced by domain factories. */
export type DomainModules = {
  tools: ToolModule[];
  resources: ResourceModule[];
  prompts: PromptModule[];
};

/** Factory that produces a middleware layer given runtime dependencies. */
export type MiddlewareFactory = (deps: RuntimeDeps) => MiddlewareLayer;

/** Shared runtime configuration returned by `createDefaultRuntime`. */
export interface CommonRuntime {
  deps: RuntimeDeps;
  middlewareFactories: MiddlewareFactory[];
  createDomainModules: () => DomainModules;
  close: () => Promise<void>;
}

export interface RuntimeOverrides {
  middlewareFactories?: MiddlewareFactory[];
  createDomainModules?: () => DomainModules;
}
