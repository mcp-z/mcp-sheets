import type { Logger, LoopbackOAuthProvider } from '@mcp-z/oauth-google';
import assert from 'assert';
import type { OAuth2Client } from 'google-auth-library';
import { google } from 'googleapis';
import createDimensionsMoveTool, { type Input, type Output } from '../../../../src/mcp/tools/dimensions-move.js';
import createRowsGetTool, { type Input as RowsGetInput, type Output as RowsGetOutput } from '../../../../src/mcp/tools/rows-get.js';
import { createExtra, type TypedHandler } from '../../../lib/create-extra.js';
import createMiddlewareContext from '../../../lib/create-middleware-context.js';
import { createTestSpreadsheet, deleteTestSpreadsheet } from '../../../lib/spreadsheet-helpers.js';

/**
 * OPTIMIZATION STRATEGY - dimensions-move.test.ts
 *
 * Single Spreadsheet, Single Batch Write
 * ======================================
 *
 * All test data is written in a single batchUpdate call in before().
 * Tests operate on different row/column ranges to avoid interference:
 * - Rows 1-10: Test data rows for row move tests (A1:C10)
 * - Columns E-J: Test data columns for column move tests (E1:J3)
 *
 * API calls:
 * - Setup: 2 (createSpreadsheet + batchUpdate for test data)
 * - Test 1: 1 (move rows) + 1 (verify read)
 * - Test 2: 1 (move columns) + 1 (verify read)
 * - Teardown: 1 (deleteSpreadsheet)
 *
 * Total: 7 API calls
 *
 * Benefits:
 * 1. Minimizes API calls to avoid rate limiting
 * 2. Tests remain isolated through range separation
 * 3. Fast execution with shared setup
 */

describe('dimensions-move tool (service-backed tests)', () => {
  let sharedSpreadsheetId: string;
  let sharedGid: string;
  let auth: OAuth2Client;
  let authProvider: LoopbackOAuthProvider;
  let logger: Logger;
  let accountId: string;
  let dimensionsMoveHandler: TypedHandler<Input>;
  let rowsGetHandler: TypedHandler<RowsGetInput>;

  before(async () => {
    const middlewareContext = await createMiddlewareContext();
    authProvider = middlewareContext.authProvider;
    logger = middlewareContext.logger;
    auth = middlewareContext.auth;
    const middleware = middlewareContext.middleware;
    accountId = middlewareContext.accountId;

    const dimensionsMoveTool = createDimensionsMoveTool();
    const wrappedDimensionsMoveTool = middleware.withToolAuth(dimensionsMoveTool);
    dimensionsMoveHandler = wrappedDimensionsMoveTool.handler;

    const rowsGetTool = createRowsGetTool();
    const wrappedRowsGetTool = middleware.withToolAuth(rowsGetTool);
    rowsGetHandler = wrappedRowsGetTool.handler;

    const title = `ci-dimensions-move-tests-${Date.now()}`;
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
            // Test 1 data: Row move test (A1:C10)
            // Rows labeled Row1-Row10 so we can verify order after move
            {
              range: 'Sheet1!A1:C10',
              values: [
                ['Row1', 'A1', 'B1'],
                ['Row2', 'A2', 'B2'],
                ['Row3', 'A3', 'B3'],
                ['Row4', 'A4', 'B4'],
                ['Row5', 'A5', 'B5'],
                ['Row6', 'A6', 'B6'],
                ['Row7', 'A7', 'B7'],
                ['Row8', 'A8', 'B8'],
                ['Row9', 'A9', 'B9'],
                ['Row10', 'A10', 'B10'],
              ],
            },
            // Test 2 data: Column move test (E1:J3)
            // Columns labeled ColE-ColJ so we can verify order after move
            {
              range: 'Sheet1!E1:J3',
              values: [
                ['ColE', 'ColF', 'ColG', 'ColH', 'ColI', 'ColJ'],
                ['E2', 'F2', 'G2', 'H2', 'I2', 'J2'],
                ['E3', 'F3', 'G3', 'H3', 'I3', 'J3'],
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

  it('dimensions-move moves rows down within sheet', async () => {
    // Move rows 2-4 (indices 1-4, 0-based) to position 8 (index 8)
    // Before: Row1, Row2, Row3, Row4, Row5, Row6, Row7, Row8, Row9, Row10
    // After:  Row1, Row5, Row6, Row7, Row8, Row2, Row3, Row4, Row9, Row10
    const moveResult = await dimensionsMoveHandler(
      {
        id: sharedSpreadsheetId,
        gid: sharedGid,
        dimension: 'ROWS',
        startIndex: 1,
        endIndex: 4,
        destinationIndex: 8,
      },
      createExtra()
    );

    const moveBranch = moveResult.structuredContent?.result as Output | undefined;
    assert.equal(moveBranch?.type, 'success', 'move should succeed');

    if (moveBranch?.type === 'success') {
      assert.equal(moveBranch.dimension, 'ROWS', 'dimension should be ROWS');
      assert.equal(moveBranch.movedCount, 3, 'should have moved 3 rows');
      assert.equal(moveBranch.sourceRange.startIndex, 1, 'source startIndex should be 1');
      assert.equal(moveBranch.sourceRange.endIndex, 4, 'source endIndex should be 4');
      assert.equal(moveBranch.destinationIndex, 8, 'destinationIndex should be 8');
    }

    // Verify the rows were actually moved by reading the data
    const readResult = await rowsGetHandler({ id: sharedSpreadsheetId, gid: sharedGid, range: 'A1:A10' }, createExtra());
    const readBranch = readResult.structuredContent?.result as RowsGetOutput | undefined;
    assert.equal(readBranch?.type, 'success', 'read should succeed');

    if (readBranch?.type === 'success') {
      // After moving rows 2-4 to position 8:
      // Expected order: Row1, Row5, Row6, Row7, Row8, Row2, Row3, Row4, Row9, Row10
      const rowLabels = readBranch.rows.map((row) => row[0]);
      assert.equal(rowLabels[0], 'Row1', 'Row1 should stay at position 0');
      assert.equal(rowLabels[1], 'Row5', 'Row5 should now be at position 1');
      assert.equal(rowLabels[5], 'Row2', 'Row2 should now be at position 5');
      assert.equal(rowLabels[6], 'Row3', 'Row3 should now be at position 6');
      assert.equal(rowLabels[7], 'Row4', 'Row4 should now be at position 7');
    }
  });

  it('dimensions-move moves columns right within sheet', async () => {
    // Move columns E-F (indices 4-6, 0-based) to position 9 (after column I)
    // Before columns at E-J: ColE, ColF, ColG, ColH, ColI, ColJ
    // After:                 ColG, ColH, ColI, ColE, ColF, ColJ
    const moveResult = await dimensionsMoveHandler(
      {
        id: sharedSpreadsheetId,
        gid: sharedGid,
        dimension: 'COLUMNS',
        startIndex: 4,
        endIndex: 6,
        destinationIndex: 9,
      },
      createExtra()
    );

    const moveBranch = moveResult.structuredContent?.result as Output | undefined;
    assert.equal(moveBranch?.type, 'success', 'move should succeed');

    if (moveBranch?.type === 'success') {
      assert.equal(moveBranch.dimension, 'COLUMNS', 'dimension should be COLUMNS');
      assert.equal(moveBranch.movedCount, 2, 'should have moved 2 columns');
      assert.equal(moveBranch.sourceRange.startIndex, 4, 'source startIndex should be 4');
      assert.equal(moveBranch.sourceRange.endIndex, 6, 'source endIndex should be 6');
      assert.equal(moveBranch.destinationIndex, 9, 'destinationIndex should be 9');
    }

    // Verify the columns were actually moved by reading the header row
    const readResult = await rowsGetHandler({ id: sharedSpreadsheetId, gid: sharedGid, range: 'E1:J1' }, createExtra());
    const readBranch = readResult.structuredContent?.result as RowsGetOutput | undefined;
    assert.equal(readBranch?.type, 'success', 'read should succeed');

    if (readBranch?.type === 'success' && readBranch.rows[0]) {
      // After moving columns E-F (indices 4-5) to position 9:
      // Expected order in E-J: ColG, ColH, ColI, ColE, ColF, ColJ
      const colLabels = readBranch.rows[0];
      assert.equal(colLabels[0], 'ColG', 'ColG should now be at column E');
      assert.equal(colLabels[1], 'ColH', 'ColH should now be at column F');
      assert.equal(colLabels[2], 'ColI', 'ColI should now be at column G');
      assert.equal(colLabels[3], 'ColE', 'ColE should now be at column H');
      assert.equal(colLabels[4], 'ColF', 'ColF should now be at column I');
      assert.equal(colLabels[5], 'ColJ', 'ColJ should stay at column J');
    }
  });
});
