#!/usr/bin/env node

/**
 * Sheets Test Token Setup
 *
 * Generates OAuth token for Sheets server tests in local .tokens/ directory.
 *
 * Usage:
 *   npm run test:setup
 */

import { LoopbackOAuthProvider } from '@mcp-z/oauth-google';
import { GOOGLE_SCOPE } from '../src/constants.ts';
import createStore from '../src/lib/create-store.ts';
import { createConfig } from '../src/setup/config.ts';

async function setupTest(): Promise<void> {
  const config = createConfig();
  console.log('🔐 Sheets Test Token Setup');
  console.log('');

  // Use local .tokens/ directory at package root (no subdirectories)
  const tokenStore = await createStore<unknown>('file://.//.tokens/store.json');

  const auth = new LoopbackOAuthProvider({
    service: config.name,
    clientId: config.clientId,
    clientSecret: config.clientSecret,
    scope: GOOGLE_SCOPE,
    headless: false,
    logger: console,
    tokenStore,
  });

  console.log('Starting OAuth flow...');
  console.log('');

  // Trigger OAuth flow via middleware (handles auth_url by opening browser + polling)
  const middleware = auth.authMiddleware();
  const setupTool = middleware.withToolAuth({
    name: 'test-setup',
    config: {},
    handler: async () => {
      return { ok: true };
    },
  });
  await setupTool.handler({}, {});

  console.log('✓ OAuth flow completed, fetching user email...');

  // Get email for display (from active account)
  const email = await auth.getUserEmail();

  console.log('');
  console.log('✅ OAuth token generated successfully!');
  console.log(`📧 Authenticated as: ${email}`);
  console.log('📁 Token saved to: .tokens/store.json');
  console.log(`   Token key: ${email}:sheets:token`);
  console.log('');
  console.log('Run `npm run test:unit` to verify Sheets API integration');
}

// Run if executed directly
if (import.meta.main) {
  setupTest()
    .then(() => {
      process.exit(0);
    })
    .catch((error) => {
      console.error('\n❌ Token setup failed:', error.message);
      process.exit(1);
    });
}
