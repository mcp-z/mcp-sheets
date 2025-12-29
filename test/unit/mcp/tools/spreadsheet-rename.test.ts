import type { Logger, LoopbackOAuthProvider } from '@mcp-z/oauth-google';
import assert from 'assert';
import crypto from 'crypto';
import fs from 'fs/promises';
import type { OAuth2Client } from 'google-auth-library';
import { google } from 'googleapis';
import * as path from 'path';
import createTool, { type Input, type Output } from '../../../../src/mcp/tools/spreadsheet-rename.js';
import { createExtra, type TypedHandler } from '../../../lib/create-extra.js';
import createMiddlewareContext from '../../../lib/create-middleware-context.js';
import { createTestSpreadsheet, deleteTestSpreadsheet } from '../../../lib/spreadsheet-helpers.js';

/**
 * OPTIMIZATION NOTES - spreadsheet-rename.test.ts
 *
 * API calls: 5 total (was 6)
 * - Setup: 1 (createSpreadsheet)
 * - Test 1: 3 (handler:get + handler:batchUpdate + verify:get)
 * - Teardown: 1 (deleteSpreadsheet)
 *
 * Optimizations applied:
 * - Removed "get original title" call before handler - handler already returns oldTitle in response
 * - We know the original title from createTestSpreadsheet, so we can compare directly
 */

// Shared test resources
let sharedSpreadsheetId: string;
let originalTitle: string;
let auth: OAuth2Client;
let authProvider: LoopbackOAuthProvider;
let logger: Logger;
let accountId: string;
let handler: TypedHandler<Input>;
let tmpDir: string;

before(async () => {
  try {
    // Create temporary directory
    tmpDir = path.join('.tmp', `spreadsheet-rename-tests-${crypto.randomUUID()}`);
    await fs.mkdir(tmpDir, { recursive: true });

    // Get middleware for tool creation and auth for close operations
    const middlewareContext = await createMiddlewareContext();
    auth = middlewareContext.auth;
    logger = middlewareContext.logger;
    const middleware = middlewareContext.middleware;
    const tool = createTool();
    const wrappedTool = middleware.withToolAuth(tool);
    handler = wrappedTool.handler;
    authProvider = middlewareContext.authProvider;
    accountId = middlewareContext.accountId;

    // Create shared spreadsheet for all tests - remember the title for later verification
    originalTitle = `ci-spreadsheet-rename-tests-${Date.now()}`;
    const accessToken = await authProvider.getAccessToken(accountId);
    sharedSpreadsheetId = await createTestSpreadsheet(accessToken, { title: originalTitle });
  } catch (error) {
    logger.error('Failed to initialize test resources:', { error });
    throw error;
  }
});

after(async () => {
  // Cleanup resources - fail fast on errors
  const accessToken = await authProvider.getAccessToken(accountId);
  await deleteTestSpreadsheet(accessToken, sharedSpreadsheetId, logger);
  await fs.rm(tmpDir, { recursive: true, force: true });
});

it('spreadsheet_rename renames an existing spreadsheet', async () => {
  const newTitle = `renamed-spreadsheet-${Date.now()}`;

  // OPTIMIZATION: Skip "get original title" - we know it from createTestSpreadsheet
  // and handler returns oldTitle in its response anyway

  // Rename the spreadsheet
  const res = await handler({ id: sharedSpreadsheetId, newTitle }, createExtra());
  assert.ok(res && res.structuredContent && res.content, 'missing structured result for spreadsheet_rename');

  const branch = res.structuredContent?.result as Output | undefined;
  assert.ok(branch, 'missing structured result for spreadsheet_rename');
  assert.equal(branch.type, 'success');

  if (branch.type === 'success') {
    assert.equal(branch.id, sharedSpreadsheetId, 'spreadsheet_rename did not return spreadsheet id');
    assert.equal(branch.oldTitle, originalTitle, 'oldTitle should match original');
    assert.equal(branch.newTitle, newTitle, 'newTitle should match requested');
    assert.ok(branch.spreadsheetUrl.includes(sharedSpreadsheetId), 'spreadsheetUrl should contain spreadsheet id');

    // Verify spreadsheet was renamed via Sheets API
    const client = google.sheets({ version: 'v4', auth });
    const info = await client.spreadsheets.get({ spreadsheetId: sharedSpreadsheetId, fields: 'properties.title' });
    assert.equal(info.data.properties?.title, newTitle, 'Spreadsheet title should be updated in Google Sheets');
  }
});
