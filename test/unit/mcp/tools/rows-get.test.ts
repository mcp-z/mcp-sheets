import type { Logger, LoopbackOAuthProvider } from '@mcp-z/oauth-google';
import assert from 'assert';
import type { OAuth2Client } from 'google-auth-library';
import { google } from 'googleapis';
import createRowsGetTool, { type Input, type Output } from '../../../../src/mcp/tools/rows-get.ts';
import { createExtra, type TypedHandler } from '../../../lib/create-extra.ts';
import createMiddlewareContext from '../../../lib/create-middleware-context.ts';
import { createTestSpreadsheet, deleteTestSpreadsheet } from '../../../lib/spreadsheet-helpers.ts';

/**
 * OPTIMIZATION STRATEGY - rows-get.test.ts
 *
 * Hybrid Testing Approach for Read Operations
 * ============================================
 *
 * This test file contains smoke tests to verify basic rows-get functionality.
 * The tool's comprehensive coverage is achieved through integration with write operation tests.
 *
 * Strategy:
 * - Smoke Tests (here): Verify core read capability and render parameter
 * - Integration Tests (rows-append.test.ts, etc.): Use rows-get for validation after writes
 * - Single batch write in before(): All test data written in one API call using batchUpdate
 *
 * Test Data Layout (single batchUpdate call):
 * - A1:C3: Basic test data (Name, Age, City rows)
 * - E1:F3: Formula test data (formulas in E3:F3 for render parameter tests)
 *
 * Benefits:
 * 1. Validates rows-get in real-world usage scenarios
 * 2. Tests write operations with read-back verification (higher confidence)
 * 3. Demonstrates tool integration patterns
 * 4. Minimizes API calls to avoid rate limiting
 *
 * Coverage by Integration:
 * - Single cell reads: Validated in rows-append.test.ts
 * - Row ranges: Validated in columns-update.test.ts
 * - Full columns: Validated in search.test.ts
 * - Rectangular ranges: Validated here (smoke test)
 * - Render parameter: Validated here (FORMULA, FORMATTED_VALUE)
 *
 * See: test/unit/mcp/tools/rows-append.test.ts for integration example
 */

describe('rows-get tool (service-backed tests)', () => {
  let sharedSpreadsheetId: string;
  let sharedGid: string;
  let auth: OAuth2Client;
  let authProvider: LoopbackOAuthProvider;
  let logger: Logger;
  let accountId: string;
  let rowsGetHandler: TypedHandler<Input>;

  before(async () => {
    const middlewareContext = await createMiddlewareContext();
    authProvider = middlewareContext.authProvider;
    logger = middlewareContext.logger;
    auth = middlewareContext.auth;
    const middleware = middlewareContext.middleware;
    accountId = middlewareContext.accountId;
    const tool = createRowsGetTool();
    const wrappedTool = middleware.withToolAuth(tool);
    rowsGetHandler = wrappedTool.handler;
    const title = `ci-rows-get-tests-${Date.now()}`;
    const accessToken = await authProvider.getAccessToken(accountId);
    sharedSpreadsheetId = await createTestSpreadsheet(accessToken, { title });
    // Use default sheet (gid: 0) to minimize write operations
    sharedGid = '0';

    // Add test data using googleapis directly (faster than using tool)
    // Using batchUpdate to write all test data in a single API call
    try {
      const sheets = google.sheets({ version: 'v4', auth });
      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: sharedSpreadsheetId,
        requestBody: {
          valueInputOption: 'USER_ENTERED', // Required for formulas to be interpreted
          data: [
            // Basic test data (A1:C3) - for smoke test
            {
              range: 'Sheet1!A1:C3',
              values: [
                ['Name', 'Age', 'City'],
                ['Alice', '30', 'NYC'],
                ['Bob', '25', 'LA'],
              ],
            },
            // Formula test data (E1:F3) - for render parameter tests
            {
              range: 'Sheet1!E1:F3',
              values: [
                ['Value1', 'Value2'],
                ['10', '20'],
                ['=E2+F2', '=E2*F2'], // Formulas that calculate to 30 and 200
              ],
            },
          ],
        },
      });
    } catch (error) {
      // FAIL FAST: Throw to skip all tests if data setup fails
      throw new Error(`Failed to write test data in before() hook: ${error instanceof Error ? error.message : String(error)}. All tests will be skipped.`);
    }
  });

  after(async () => {
    // Cleanup - fail fast on errors
    const accessToken = await authProvider.getAccessToken(accountId);
    await deleteTestSpreadsheet(accessToken, sharedSpreadsheetId, logger);
  });

  it('sheets-rows-get retrieves rectangular range (smoke test)', async () => {
    // SMOKE TEST: Verifies core rows-get functionality
    // Additional coverage through integration tests in rows-append.test.ts, etc.
    const result = await rowsGetHandler({ id: sharedSpreadsheetId, gid: sharedGid, range: 'A1:C3' }, createExtra());

    const branch = result.structuredContent?.result as Output | undefined;
    assert.equal(branch?.type, 'success', 'should succeed');

    if (branch?.type === 'success') {
      assert.ok(Array.isArray(branch.rows), 'should have rows array');
      // EXACT count, not "at least" - we wrote 3 rows, expect 3 rows
      assert.equal(branch.rows.length, 3, 'should return exactly 3 rows');

      // Verify complete rectangular structure (3 rows × 3 cols = 9 cells)
      const [headerRow, aliceRow, bobRow] = branch.rows;

      // Validate header row completely
      assert.ok(headerRow, 'should have header row');
      assert.equal(headerRow.length, 3, 'header row should have 3 columns');
      assert.equal(headerRow[0], 'Name', 'header col 1');
      assert.equal(headerRow[1], 'Age', 'header col 2');
      assert.equal(headerRow[2], 'City', 'header col 3');

      // Validate data rows completely
      assert.ok(aliceRow, 'should have Alice row');
      assert.equal(aliceRow[0], 'Alice', 'Alice name');
      assert.equal(aliceRow[1], '30', 'Alice age');
      assert.equal(aliceRow[2], 'NYC', 'Alice city');

      assert.ok(bobRow, 'should have Bob row');
      assert.equal(bobRow[0], 'Bob', 'Bob name');
      assert.equal(bobRow[1], '25', 'Bob age');
      assert.equal(bobRow[2], 'LA', 'Bob city');
    }
  });

  it('sheets-rows-get with render=FORMULA returns formula text', async () => {
    // Tests render parameter - formula data written in batch during before() hook
    const result = await rowsGetHandler({ id: sharedSpreadsheetId, gid: sharedGid, range: 'E3:F3', render: 'FORMULA' }, createExtra());

    const branch = result.structuredContent?.result as Output | undefined;
    assert.equal(branch?.type, 'success', 'should succeed');

    if (branch?.type === 'success') {
      assert.ok(Array.isArray(branch.rows), 'should have rows array');
      const row = branch.rows[0];
      if (row) {
        // Should return formula strings, not calculated values
        assert.equal(row[0], '=E2+F2', 'should return formula text for first cell');
        assert.equal(row[1], '=E2*F2', 'should return formula text for second cell');
      }
    }
  });

  it('sheets-rows-get with render=FORMATTED_VALUE returns calculated values', async () => {
    const result = await rowsGetHandler({ id: sharedSpreadsheetId, gid: sharedGid, range: 'E3:F3', render: 'FORMATTED_VALUE' }, createExtra());

    const branch = result.structuredContent?.result as Output | undefined;
    assert.equal(branch?.type, 'success', 'should succeed');

    if (branch?.type === 'success') {
      assert.ok(Array.isArray(branch.rows), 'should have rows array');
      const row = branch.rows[0];
      if (row) {
        // Should return calculated values, not formulas
        assert.equal(row[0], '30', 'should return calculated value for sum formula');
        assert.equal(row[1], '200', 'should return calculated value for product formula');
      }
    }
  });
}); // End describe block
