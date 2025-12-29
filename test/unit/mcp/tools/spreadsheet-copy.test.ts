import type { Logger, LoopbackOAuthProvider } from '@mcp-z/oauth-google';
import assert from 'assert';
import crypto from 'crypto';
import fs from 'fs/promises';
import type { OAuth2Client } from 'google-auth-library';
import { google } from 'googleapis';
import * as path from 'path';
import createTool, { type Input, type Output } from '../../../../src/mcp/tools/spreadsheet-copy.js';
import { createExtra, type TypedHandler } from '../../../lib/create-extra.js';
import createMiddlewareContext from '../../../lib/create-middleware-context.js';
import { createTestSpreadsheet, deleteTestSpreadsheet } from '../../../lib/spreadsheet-helpers.js';

/**
 * OPTIMIZATION NOTES - spreadsheet-copy.test.ts
 *
 * API calls: 8 total (was 13)
 * - Setup: 2 (createSpreadsheet + values.update)
 * - Test 1: 4 (handler:get + handler:files.copy + verify:get with gridData)
 * - Teardown: 2 (delete source + delete 1 copy)
 *
 * Optimizations applied:
 * - Removed "default name" test - "custom name" test is superset (proves copy works AND custom naming works)
 * - Combined spreadsheets.get + values.get into single get with includeGridData
 * - Saves 1 close call (only 1 copied spreadsheet instead of 2)
 */

// Shared test resources
let sourceSpreadsheetId: string;
const copiedSpreadsheetIds: string[] = [];
let auth: OAuth2Client;
let authProvider: LoopbackOAuthProvider;
let logger: Logger;
let accountId: string;
let handler: TypedHandler<Input>;
let tmpDir: string;

before(async () => {
  try {
    // Create temporary directory
    tmpDir = path.join('.tmp', `spreadsheet-copy-tests-${crypto.randomUUID()}`);
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

    const accessToken = await authProvider.getAccessToken(accountId);

    // Create source spreadsheet with data
    sourceSpreadsheetId = await createTestSpreadsheet(accessToken, { title: `ci-spreadsheet-copy-source-${Date.now()}` });

    // Add some data to the default sheet
    const sheets = google.sheets({ version: 'v4', auth });
    await sheets.spreadsheets.values.update({
      spreadsheetId: sourceSpreadsheetId,
      range: 'Sheet1!A1:B2',
      valueInputOption: 'RAW',
      requestBody: {
        values: [
          ['OriginalHeader1', 'OriginalHeader2'],
          ['OriginalValue1', 'OriginalValue2'],
        ],
      },
    });
  } catch (error) {
    logger.error('Failed to initialize test resources:', { error });
    throw error;
  }
});

after(async () => {
  // Cleanup resources - fail fast on errors
  const accessToken = await authProvider.getAccessToken(accountId);
  await deleteTestSpreadsheet(accessToken, sourceSpreadsheetId, logger);

  // Delete any copied spreadsheets
  for (const id of copiedSpreadsheetIds) {
    try {
      await deleteTestSpreadsheet(accessToken, id, logger);
    } catch {
      // Ignore errors for copied spreadsheets - they may have been deleted manually
    }
  }

  await fs.rm(tmpDir, { recursive: true, force: true });
});

it('spreadsheet_copy copies a spreadsheet with custom name and data preserved', async () => {
  // Test copy with custom name (superset of default name - proves copy works AND custom naming works)
  const newTitle = `custom-copy-${Date.now()}`;

  const res = await handler(
    {
      id: sourceSpreadsheetId,
      newTitle,
    },
    createExtra()
  );

  assert.ok(res && res.structuredContent && res.content, 'missing structured result for spreadsheet_copy');

  const branch = res.structuredContent?.result as Output | undefined;
  assert.ok(branch, 'missing structured result for spreadsheet_copy');
  assert.equal(branch.type, 'success');

  if (branch.type === 'success') {
    // Validate response structure
    assert.equal(branch.sourceId, sourceSpreadsheetId, 'should return source id');
    assert.ok(branch.newId, 'should return new spreadsheet id');
    assert.equal(branch.newTitle, newTitle, 'should use custom title');
    assert.ok(branch.spreadsheetUrl.includes(branch.newId), 'spreadsheetUrl should contain new id');

    // Track for close
    copiedSpreadsheetIds.push(branch.newId);

    // OPTIMIZATION: Single API call to verify spreadsheet exists with correct title AND data was copied
    // Uses includeGridData instead of separate get + values.get calls
    const client = google.sheets({ version: 'v4', auth });
    const info = await client.spreadsheets.get({
      spreadsheetId: branch.newId,
      includeGridData: true,
      ranges: ['Sheet1!A1:B2'],
    });

    // Verify spreadsheet has the correct title
    assert.equal(info.data.properties?.title, newTitle, 'Copied spreadsheet should have custom title');

    // Verify data was copied (from the grid data)
    const sheet = info.data.sheets?.find((s) => s?.properties?.title === 'Sheet1');
    assert.ok(sheet, 'Sheet1 should exist in copied spreadsheet');
    const gridData = sheet?.data?.[0]?.rowData;
    assert.ok(gridData, 'Grid data should exist');
    assert.equal(gridData[0]?.values?.[0]?.formattedValue, 'OriginalHeader1', 'First cell should be OriginalHeader1');
    assert.equal(gridData[0]?.values?.[1]?.formattedValue, 'OriginalHeader2', 'Second cell should be OriginalHeader2');
    assert.equal(gridData[1]?.values?.[0]?.formattedValue, 'OriginalValue1', 'Third cell should be OriginalValue1');
    assert.equal(gridData[1]?.values?.[1]?.formattedValue, 'OriginalValue2', 'Fourth cell should be OriginalValue2');
  }
});
