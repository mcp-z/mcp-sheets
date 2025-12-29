import { AccountServer, type AuthEmailProvider } from '@mcp-z/oauth';
import type { CachedToken } from '@mcp-z/oauth-google';
import { createDcrRouter, DcrOAuthProvider, LoopbackOAuthProvider, ServiceAccountProvider } from '@mcp-z/oauth-google';
import type { Logger, PromptModule, ToolModule } from '@mcp-z/server';
import type { Router } from 'express';
import type { Keyv } from 'keyv';
import { GOOGLE_SCOPE } from '../constants.ts';
import type { ServerConfig } from '../types.js';

/**
 * Gmail OAuth runtime dependencies.
 */
export interface OAuthRuntimeDeps {
  logger: Logger;
  tokenStore: Keyv<CachedToken>;
  dcrStore?: Keyv<unknown>;
}

/**
 * Auth middleware helpers used to wrap MCP modules.
 */
export interface AuthMiddleware {
  withToolAuth<T extends { name: string; config: unknown; handler: unknown }>(module: T): T;
  withResourceAuth<T extends { name: string; template?: unknown; config?: unknown; handler: unknown }>(module: T): T;
  withPromptAuth<T extends { name: string; config: unknown; handler: unknown }>(module: T): T;
}

/**
 * Result returned by createOAuthAdapters.
 */
export interface OAuthAdapters {
  primary: LoopbackOAuthProvider | ServiceAccountProvider | DcrOAuthProvider;
  middleware: AuthMiddleware;
  authAdapter: AuthEmailProvider;
  accountTools: ToolModule[];
  accountPrompts: PromptModule[];
  dcrRouter?: Router;
}

/**
 * Create Sheets OAuth adapters and helpers.
 *
 * @param config Sheets server configuration.
 * @param deps Runtime dependencies (logger, token store, optional DCR store).
 */
export async function createOAuthAdapters(config: ServerConfig, deps: OAuthRuntimeDeps, baseUrl?: string): Promise<OAuthAdapters> {
  const { logger, tokenStore, dcrStore } = deps;
  const oauthStaticConfig = {
    service: config.name,
    clientId: config.clientId,
    clientSecret: config.clientSecret,
    scope: GOOGLE_SCOPE,
    auth: config.auth,
    headless: config.headless,
    redirectUri: config.transport.type === 'stdio' ? undefined : config.redirectUri,
    ...(config.serviceAccountKeyFile && { serviceAccountKeyFile: config.serviceAccountKeyFile }),
    ...(baseUrl && { baseUrl }),
  };

  let primary: LoopbackOAuthProvider | ServiceAccountProvider | DcrOAuthProvider;

  if (oauthStaticConfig.auth === 'dcr') {
    logger.debug('Creating DCR provider', { service: oauthStaticConfig.service });

    if (!dcrStore) {
      throw new Error('DCR mode requires dcrStore to be configured');
    }
    if (!oauthStaticConfig.baseUrl) {
      throw new Error('DCR mode requires baseUrl to be configured');
    }

    primary = new DcrOAuthProvider({
      clientId: oauthStaticConfig.clientId,
      ...(oauthStaticConfig.clientSecret && { clientSecret: oauthStaticConfig.clientSecret }),
      scope: oauthStaticConfig.scope,
      verifyEndpoint: `${oauthStaticConfig.baseUrl}/oauth/verify`,
      logger,
    });

    const dcrRouter = createDcrRouter({
      store: dcrStore,
      issuerUrl: oauthStaticConfig.baseUrl,
      baseUrl: oauthStaticConfig.baseUrl,
      scopesSupported: oauthStaticConfig.scope.split(' '),
      clientConfig: {
        clientId: oauthStaticConfig.clientId,
        ...(oauthStaticConfig.clientSecret && { clientSecret: oauthStaticConfig.clientSecret }),
      },
    });

    const middleware = primary.authMiddleware();
    const authAdapter: AuthEmailProvider = {
      getUserEmail: () => {
        throw new Error('DCR mode does not support getUserEmail - tokens are provided via bearer auth');
      },
    };

    return {
      primary,
      middleware: middleware as unknown as AuthMiddleware,
      authAdapter,
      accountTools: [],
      accountPrompts: [],
      dcrRouter,
    };
  }

  if (oauthStaticConfig.auth === 'service-account') {
    if (!oauthStaticConfig.serviceAccountKeyFile) {
      throw new Error('Service account key file is required when auth mode is "service-account". Set GOOGLE_SERVICE_ACCOUNT_KEY_FILE environment variable or use --service-account-key-file flag.');
    }

    logger.debug('Creating service account provider', { service: oauthStaticConfig.service });
    primary = new ServiceAccountProvider({
      keyFilePath: oauthStaticConfig.serviceAccountKeyFile,
      scopes: oauthStaticConfig.scope.split(' '),
      logger,
    });
  } else {
    logger.debug('Creating loopback OAuth provider', { service: oauthStaticConfig.service });
    primary = new LoopbackOAuthProvider({
      service: oauthStaticConfig.service,
      clientId: oauthStaticConfig.clientId,
      clientSecret: oauthStaticConfig.clientSecret,
      scope: oauthStaticConfig.scope,
      headless: oauthStaticConfig.headless,
      logger,
      tokenStore,
      ...(oauthStaticConfig.redirectUri !== undefined && { redirectUri: oauthStaticConfig.redirectUri }),
    });
  }

  const authAdapter: AuthEmailProvider = {
    getUserEmail: (accountId) => primary.getUserEmail(accountId),
    ...('authenticateNewAccount' in primary && primary.authenticateNewAccount
      ? {
          authenticateNewAccount: () => primary.authenticateNewAccount?.(),
        }
      : {}),
  };

  let middleware: ReturnType<LoopbackOAuthProvider['authMiddleware']>;
  let accountTools: ToolModule[];
  let accountPrompts: PromptModule[];

  if (oauthStaticConfig.auth === 'service-account') {
    middleware = primary.authMiddleware();
    accountTools = [];
    accountPrompts = [];
    logger.debug('Service account mode - no account tools', { service: oauthStaticConfig.service });
  } else {
    middleware = primary.authMiddleware();

    const result = AccountServer.createLoopback({
      service: oauthStaticConfig.service,
      store: tokenStore,
      logger,
      auth: authAdapter,
    });
    accountTools = result.tools as ToolModule[];
    accountPrompts = result.prompts as PromptModule[];
    logger.debug('Loopback OAuth (multi-account mode)', { service: oauthStaticConfig.service });
  }

  return {
    primary,
    middleware: middleware as unknown as AuthMiddleware,
    authAdapter,
    accountTools,
    accountPrompts,
  };
}
