import type { Logger, LoopbackOAuthProvider } from '@mcp-z/oauth-google';
import assert from 'assert';
import crypto from 'crypto';
import fs from 'fs/promises';
import type { OAuth2Client } from 'google-auth-library';
import { google } from 'googleapis';
import * as path from 'path';
import createTool, { type Input, type Output } from '../../../../src/mcp/tools/sheet-copy-to.js';
import { createExtra, type TypedHandler } from '../../../lib/create-extra.js';
import createMiddlewareContext from '../../../lib/create-middleware-context.js';
import { createTestSheet, createTestSpreadsheet, deleteTestSpreadsheet } from '../../../lib/spreadsheet-helpers.js';

/**
 * OPTIMIZATION NOTES - sheet-copy-to.test.ts
 *
 * API calls: 11 total (was 16)
 * - Setup: 4 (2×createSpreadsheet + createSheet + values.update)
 * - Test 1: 4 (handler:get + handler:copyTo + handler:batchUpdate(rename) + verify:get with gridData)
 * - Test 2: 1 (handler:get - fails immediately)
 * - Teardown: 2 (2×deleteSpreadsheet)
 *
 * Optimizations applied:
 * - Removed "copy without rename" test - "copy with rename" is superset (proves copy works AND rename works)
 * - Combined spreadsheets.get + values.get into single get with includeGridData
 */

// Shared test resources
let sourceSpreadsheetId: string;
let destinationSpreadsheetId: string;
let auth: OAuth2Client;
let authProvider: LoopbackOAuthProvider;
let logger: Logger;
let accountId: string;
let handler: TypedHandler<Input>;
let tmpDir: string;
let sourceSheetId: number;

before(async () => {
  try {
    // Create temporary directory
    tmpDir = path.join('.tmp', `sheet-copy-to-tests-${crypto.randomUUID()}`);
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

    // Create source spreadsheet
    sourceSpreadsheetId = await createTestSpreadsheet(accessToken, { title: `ci-sheet-copy-to-source-${Date.now()}` });

    // Create destination spreadsheet
    destinationSpreadsheetId = await createTestSpreadsheet(accessToken, { title: `ci-sheet-copy-to-dest-${Date.now()}` });

    // Create a source sheet to copy from
    sourceSheetId = await createTestSheet(accessToken, sourceSpreadsheetId, { title: 'SourceSheet' });

    // Add some data to the source sheet
    const sheets = google.sheets({ version: 'v4', auth });
    await sheets.spreadsheets.values.update({
      spreadsheetId: sourceSpreadsheetId,
      range: 'SourceSheet!A1:B2',
      valueInputOption: 'RAW',
      requestBody: {
        values: [
          ['SourceHeader1', 'SourceHeader2'],
          ['SourceValue1', 'SourceValue2'],
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
  await deleteTestSpreadsheet(accessToken, destinationSpreadsheetId, logger);
  await fs.rm(tmpDir, { recursive: true, force: true });
});

it('sheet_copy_to copies a sheet to another spreadsheet with rename and data preserved', async () => {
  // Test copy with rename (superset of copy without rename - proves both copy AND rename work)
  const newTitle = `renamed-copy-${Date.now()}`;

  const res = await handler(
    {
      sourceId: sourceSpreadsheetId,
      sourceGid: String(sourceSheetId),
      destinationId: destinationSpreadsheetId,
      newTitle,
    },
    createExtra()
  );

  assert.ok(res && res.structuredContent && res.content, 'missing structured result for sheet_copy_to');

  const branch = res.structuredContent?.result as Output | undefined;
  assert.ok(branch, 'missing structured result for sheet_copy_to');
  assert.equal(branch.type, 'success');

  if (branch.type === 'success') {
    // Validate response structure
    assert.equal(branch.sourceId, sourceSpreadsheetId, 'should return source spreadsheet id');
    assert.equal(branch.sourceGid, String(sourceSheetId), 'should return source gid');
    assert.equal(branch.sourceTitle, 'SourceSheet', 'should return source title');
    assert.equal(branch.destinationId, destinationSpreadsheetId, 'should return destination id');
    assert.ok(branch.destinationGid, 'should return destination gid');
    assert.equal(branch.destinationTitle, newTitle, 'destination title should be the new title');
    assert.equal(branch.renamed, true, 'should be marked as renamed');
    assert.ok(branch.sheetUrl.includes(destinationSpreadsheetId), 'sheetUrl should contain destination id');

    // OPTIMIZATION: Single API call to verify sheet exists with correct title AND data was copied
    // Uses includeGridData instead of separate get + values.get calls
    const client = google.sheets({ version: 'v4', auth });
    const info = await client.spreadsheets.get({
      spreadsheetId: destinationSpreadsheetId,
      includeGridData: true,
      ranges: [`'${newTitle}'!A1:B2`],
    });

    // Verify sheet exists with correct title
    const copiedSheet = info.data.sheets?.find((s) => s?.properties?.title === newTitle);
    assert.ok(copiedSheet, `Sheet with title "${newTitle}" should exist in destination`);
    assert.equal(String(copiedSheet?.properties?.sheetId), branch.destinationGid, 'Sheet gid should match');

    // Verify data was copied (from the grid data)
    const gridData = copiedSheet?.data?.[0]?.rowData;
    assert.ok(gridData, 'Grid data should exist');
    assert.equal(gridData[0]?.values?.[0]?.formattedValue, 'SourceHeader1', 'First cell should be SourceHeader1');
    assert.equal(gridData[0]?.values?.[1]?.formattedValue, 'SourceHeader2', 'Second cell should be SourceHeader2');
    assert.equal(gridData[1]?.values?.[0]?.formattedValue, 'SourceValue1', 'Third cell should be SourceValue1');
    assert.equal(gridData[1]?.values?.[1]?.formattedValue, 'SourceValue2', 'Fourth cell should be SourceValue2');
  }
});

it('sheet_copy_to fails for non-existent source sheet', async () => {
  const nonExistentGid = '999999999';

  try {
    await handler(
      {
        sourceId: sourceSpreadsheetId,
        sourceGid: nonExistentGid,
        destinationId: destinationSpreadsheetId,
      },
      createExtra()
    );
    assert.fail('Expected error for non-existent source sheet');
  } catch (error) {
    assert.ok(error instanceof Error, 'Should throw an error');
    assert.ok(error.message.includes('not found'), 'Error should mention sheet not found');
  }
});
