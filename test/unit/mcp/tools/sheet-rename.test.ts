import type { Logger, LoopbackOAuthProvider } from '@mcp-z/oauth-google';
import assert from 'assert';
import crypto from 'crypto';
import fs from 'fs/promises';
import type { OAuth2Client } from 'google-auth-library';
import { google } from 'googleapis';
import * as path from 'path';
import createTool, { type Input, type Output } from '../../../../src/mcp/tools/sheet-rename.ts';
import { createExtra, type TypedHandler } from '../../../lib/create-extra.ts';
import createMiddlewareContext from '../../../lib/create-middleware-context.ts';
import { createTestSpreadsheet, deleteTestSpreadsheet } from '../../../lib/spreadsheet-helpers.ts';

/**
 * OPTIMIZATION NOTES - sheet-rename.test.ts
 *
 * API calls: 6 total (was 8)
 * - Setup: 1 (createSpreadsheet)
 * - Test 1: 3 (handler:get + handler:batchUpdate + verify:get)
 * - Test 2: 1 (handler:get - fails immediately)
 * - Teardown: 1 (deleteSpreadsheet)
 *
 * Optimizations applied:
 * - Use default sheet (gid=0) instead of creating a new sheet (saves 1 API call)
 * - Removed createTestSheet import (no longer needed)
 */

// Shared test resources
let sharedSpreadsheetId: string;
let auth: OAuth2Client;
let authProvider: LoopbackOAuthProvider;
let logger: Logger;
let accountId: string;
let handler: TypedHandler<Input>;
let tmpDir: string;
let defaultSheetOriginalTitle: string;

before(async () => {
  try {
    // Create temporary directory
    tmpDir = path.join('.tmp', `sheet-rename-tests-${crypto.randomUUID()}`);
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

    // Create shared spreadsheet for all tests
    const title = `ci-sheet-rename-tests-${Date.now()}`;
    const accessToken = await authProvider.getAccessToken(accountId);
    sharedSpreadsheetId = await createTestSpreadsheet(accessToken, { title });

    // Note: New spreadsheets come with a default sheet named "Sheet1" (gid=0)
    defaultSheetOriginalTitle = 'Sheet1';
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

it('sheet_rename renames an existing sheet', async () => {
  // OPTIMIZATION: Use the default sheet (gid=0) instead of creating a new one
  const defaultSheetGid = '0';
  const newTitle = `renamed-sheet-${Date.now()}`;

  // Rename the default sheet
  const res = await handler({ id: sharedSpreadsheetId, gid: defaultSheetGid, newTitle }, createExtra());
  assert.ok(res && res.structuredContent && res.content, 'missing structured result for sheet_rename');

  const branch = res.structuredContent?.result as Output | undefined;
  assert.ok(branch, 'missing structured result for sheet_rename');
  assert.equal(branch.type, 'success');

  if (branch.type === 'success') {
    assert.equal(branch.id, sharedSpreadsheetId, 'sheet_rename did not return spreadsheet id');
    assert.equal(branch.gid, defaultSheetGid, 'sheet_rename did not return correct gid');
    assert.equal(branch.oldTitle, defaultSheetOriginalTitle, 'oldTitle should match original');
    assert.equal(branch.newTitle, newTitle, 'newTitle should match requested');
    assert.ok(branch.sheetUrl.includes(sharedSpreadsheetId), 'sheetUrl should contain spreadsheet id');

    // Verify sheet was renamed via Sheets API
    const client = google.sheets({ version: 'v4', auth });
    const info = await client.spreadsheets.get({ spreadsheetId: sharedSpreadsheetId, fields: 'sheets.properties' });
    const renamedSheet = info.data.sheets?.find((s) => s?.properties?.sheetId === 0);
    assert.equal(renamedSheet?.properties?.title, newTitle, 'Sheet title should be updated in Google Sheets');
  }
});

it('sheet_rename fails for non-existent sheet', async () => {
  const nonExistentGid = '999999999';

  try {
    await handler({ id: sharedSpreadsheetId, gid: nonExistentGid, newTitle: 'should-fail' }, createExtra());
    assert.fail('Expected error for non-existent sheet');
  } catch (error) {
    assert.ok(error instanceof Error, 'Should throw an error');
    assert.ok(error.message.includes('not found'), 'Error should mention sheet not found');
  }
});
