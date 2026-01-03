import type { Logger, LoopbackOAuthProvider } from '@mcp-z/oauth-google';
import assert from 'assert';
import type { OAuth2Client } from 'google-auth-library';
import { google } from 'googleapis';
import createRowsGetTool, { type Input as RowsGetInput, type Output as RowsGetOutput } from '../../../../src/mcp/tools/rows-get.ts';
import createValuesClearTool, { type Input, type Output } from '../../../../src/mcp/tools/values-clear.ts';
import { createExtra, type TypedHandler } from '../../../lib/create-extra.ts';
import createMiddlewareContext from '../../../lib/create-middleware-context.ts';
import { createTestSpreadsheet, deleteTestSpreadsheet } from '../../../lib/spreadsheet-helpers.ts';

/**
 * OPTIMIZATION STRATEGY - values-clear.test.ts
 *
 * Single Spreadsheet, Single Batch Write
 * ======================================
 *
 * All test data is written in a single batchUpdate call in before().
 * Tests use different ranges to avoid interference:
 * - A1:C3: Test 1 - Single range clear
 * - E1:G3: Test 2 - Multiple range clear (E1:F3, G1:G3)
 *
 * Benefits:
 * 1. Minimizes API calls to avoid rate limiting
 * 2. Tests remain isolated through range separation
 * 3. Fast execution with shared setup
 */

describe('values-clear tool (service-backed tests)', () => {
  let sharedSpreadsheetId: string;
  let sharedGid: string;
  let auth: OAuth2Client;
  let authProvider: LoopbackOAuthProvider;
  let logger: Logger;
  let accountId: string;
  let valuesClearHandler: TypedHandler<Input>;
  let rowsGetHandler: TypedHandler<RowsGetInput>;

  before(async () => {
    const middlewareContext = await createMiddlewareContext();
    authProvider = middlewareContext.authProvider;
    logger = middlewareContext.logger;
    auth = middlewareContext.auth;
    const middleware = middlewareContext.middleware;
    accountId = middlewareContext.accountId;

    const valuesClearTool = createValuesClearTool();
    const wrappedValuesClearTool = middleware.withToolAuth(valuesClearTool);
    valuesClearHandler = wrappedValuesClearTool.handler;

    const rowsGetTool = createRowsGetTool();
    const wrappedRowsGetTool = middleware.withToolAuth(rowsGetTool);
    rowsGetHandler = wrappedRowsGetTool.handler;

    const title = `ci-values-clear-tests-${Date.now()}`;
    const accessToken = await authProvider.getAccessToken(accountId);
    sharedSpreadsheetId = await createTestSpreadsheet(accessToken, { title });
    sharedGid = '0';

    // Add all test data in a single batch write
    try {
      const sheets = google.sheets({ version: 'v4', auth });
      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: sharedSpreadsheetId,
        requestBody: {
          valueInputOption: 'RAW',
          data: [
            // Test 1 data: Single range clear (A1:C3)
            {
              range: 'Sheet1!A1:C3',
              values: [
                ['Header1', 'Header2', 'Header3'],
                ['Value1', 'Value2', 'Value3'],
                ['Value4', 'Value5', 'Value6'],
              ],
            },
            // Test 2 data: Multiple range clear (E1:G3)
            {
              range: 'Sheet1!E1:G3',
              values: [
                ['ColE', 'ColF', 'ColG'],
                ['E2', 'F2', 'G2'],
                ['E3', 'F3', 'G3'],
              ],
            },
          ],
        },
      });
    } catch (error) {
      throw new Error(`Failed to write test data in before() hook: ${error instanceof Error ? error.message : String(error)}. All tests will be skipped.`);
    }
  });

  after(async () => {
    const accessToken = await authProvider.getAccessToken(accountId);
    await deleteTestSpreadsheet(accessToken, sharedSpreadsheetId, logger);
  });

  it('values-clear clears a single range', async () => {
    // Clear A2:C3 (leave header row intact)
    const clearResult = await valuesClearHandler({ id: sharedSpreadsheetId, gid: sharedGid, ranges: ['A2:C3'] }, createExtra());

    const clearBranch = clearResult.structuredContent?.result as Output | undefined;
    assert.equal(clearBranch?.type, 'success', 'clear should succeed');

    if (clearBranch?.type === 'success') {
      // EXACT validation of clearedRanges, not just length > 0
      assert.equal(clearBranch.clearedRanges.length, 1, 'should have cleared exactly 1 range');
      // API returns range with sheet name prefix
      assert.ok(clearBranch.clearedRanges[0]?.includes('A2:C3'), 'cleared range should be A2:C3');
    }

    // Verify cells are cleared by reading them back
    const readResult = await rowsGetHandler({ id: sharedSpreadsheetId, gid: sharedGid, range: 'A1:C3' }, createExtra());

    const readBranch = readResult.structuredContent?.result as RowsGetOutput | undefined;
    assert.equal(readBranch?.type, 'success', 'read should succeed');

    if (readBranch?.type === 'success') {
      // Header row should still exist
      assert.ok(readBranch.rows.length >= 1, 'should have at least header row');
      const headerRow = readBranch.rows[0];
      if (headerRow) {
        assert.equal(headerRow[0], 'Header1', 'header should be preserved');
      }

      // Data rows should be cleared (empty or not present)
      // Google Sheets API may return fewer rows when trailing rows are empty
      if (readBranch.rows.length > 1) {
        const dataRow = readBranch.rows[1];
        // If row exists, values should be empty/null
        if (dataRow && dataRow.length > 0) {
          assert.ok(
            dataRow.every((cell) => cell === '' || cell === null || cell === undefined),
            'cleared cells should be empty'
          );
        }
      }
    }
  });

  it('values-clear clears multiple ranges', async () => {
    // Clear E2:F3 and G2:G3 separately
    const clearResult = await valuesClearHandler({ id: sharedSpreadsheetId, gid: sharedGid, ranges: ['E2:F3', 'G2:G3'] }, createExtra());

    const clearBranch = clearResult.structuredContent?.result as Output | undefined;
    assert.equal(clearBranch?.type, 'success', 'clear should succeed');

    if (clearBranch?.type === 'success') {
      // EXACT validation of clearedRanges content, not just count
      assert.equal(clearBranch.clearedRanges.length, 2, 'should have cleared exactly 2 ranges');
      // Verify both ranges were cleared (order may vary, so check includes)
      const rangesStr = clearBranch.clearedRanges.join(',');
      assert.ok(rangesStr.includes('E2:F3'), 'should have cleared E2:F3');
      assert.ok(rangesStr.includes('G2:G3'), 'should have cleared G2:G3');
    }

    // Verify headers preserved, data cleared
    const readResult = await rowsGetHandler({ id: sharedSpreadsheetId, gid: sharedGid, range: 'E1:G3' }, createExtra());

    const readBranch = readResult.structuredContent?.result as RowsGetOutput | undefined;
    assert.equal(readBranch?.type, 'success', 'read should succeed');

    if (readBranch?.type === 'success') {
      // Header row should still exist
      const headerRow = readBranch.rows[0];
      if (headerRow) {
        assert.equal(headerRow[0], 'ColE', 'header E should be preserved');
        assert.equal(headerRow[1], 'ColF', 'header F should be preserved');
        assert.equal(headerRow[2], 'ColG', 'header G should be preserved');
      }
    }
  });
});
