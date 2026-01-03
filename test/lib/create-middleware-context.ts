/**
 * Create auth middleware for testing
 *
 * This helper sets up the auth middleware in single-account mode,
 * which is appropriate for unit tests. It replaces the old context pattern.
 */

import { listAccountIds } from '@mcp-z/oauth';
import { LoopbackOAuthProvider } from '@mcp-z/oauth-google';
import type { Keyv } from 'keyv';
import { GOOGLE_SCOPE } from '../../src/constants.ts';
import createStore from '../../src/lib/create-store.ts';
import type { Logger } from '../../src/types.ts';
import { createConfig } from './config.ts';

/**
 * Validate exactly one account exists (0/1/Many strategy)
 *
 * Uses listAccountIds utility from @mcp-z/oauth for proper key parsing.
 * For single-account tests, expects exactly one account
 */
async function validateSingleAccount(tokenStore: Keyv, service: string): Promise<string> {
  const accountIds = await listAccountIds(tokenStore, service);

  if (accountIds.length === 0) {
    throw new Error(`No test account found for ${service}. Run \`npm run test:setup\` to generate OAuth token.`);
  }

  if (accountIds.length > 1) {
    throw new Error(`Multiple test accounts found for ${service} (${accountIds.length} accounts). Tests require exactly one account for determinism. Clean and regenerate:\n  rm -rf .tokens\n  npm run test:setup`);
  }

  const accountId = accountIds[0];
  if (!accountId) {
    throw new Error('Internal error: validated account array has no first element');
  }
  return accountId;
}

/**
 * Create test middleware with per-package local token storage.
 * Uses local .tokens/ directory in single-account mode.
 *
 * Test Token Strategy:
 * - Location: <package-root>/.tokens/google/
 * - Validation: Exactly 1 account (0/1/many validation)
 * - Mode: single-account (simplest for unit tests)
 * - Isolated from other packages and production tokens
 *
 * @returns Middleware wrapper function (middleware)
 */
export default async function createMiddlewareContext() {
  const config = createConfig();
  const logger: Logger = {
    debug: (_msg: string, _meta?: Record<string, unknown>) => {},
    info: (_msg: string, _meta?: Record<string, unknown>) => {},
    warn: (msg: string, meta?: Record<string, unknown>) => console.warn(msg, meta),
    error: (msg: string, meta?: Record<string, unknown>) => console.error(msg, meta),
  };

  // Local .tokens/ directory at package root
  const tokenStore = await createStore<unknown>('file://.//.tokens/store.json');

  // Validate exactly 1 account exists
  const accountId = await validateSingleAccount(tokenStore, config.name);

  const authProvider = new LoopbackOAuthProvider({
    service: config.name,
    clientId: config.clientId,
    clientSecret: config.clientSecret,
    scope: GOOGLE_SCOPE,
    headless: true, // No browser in tests
    logger,
    tokenStore,
  });

  logger.info({ accountId }, 'Creating Sheets test middleware');

  // Create middleware in single-account mode (simplest for unit tests)
  const middleware = authProvider.authMiddleware();

  // Return everything needed - middleware, auth client, and logger
  // This allows tests to use middleware pattern for tools AND get direct auth for test setup
  return {
    middleware,
    auth: authProvider.toAuth(accountId),
    authProvider,
    logger,
    accountId,
  };
}
