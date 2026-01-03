import type { Logger, LoopbackOAuthProvider } from '@mcp-z/oauth-google';
import assert from 'assert';
import crypto from 'crypto';
import fs from 'fs/promises';
import * as path from 'path';
import createTool, { type Input, type Output } from '../../../../src/mcp/tools/values-batch-update.ts';
import { createExtra, type TypedHandler } from '../../../lib/create-extra.ts';
import createMiddlewareContext from '../../../lib/create-middleware-context.ts';
import { createTestSpreadsheet, deleteTestSpreadsheet } from '../../../lib/spreadsheet-helpers.ts';

describe('values-batch-update tool (service-backed tests)', () => {
  // Shared test resources
  let sharedSpreadsheetId: string;
  let authProvider: LoopbackOAuthProvider;
  let logger: Logger;
  let accountId: string;
  let handler: TypedHandler<Input>;
  let tmpDir: string;

  before(async () => {
    try {
      // Create temporary directory
      tmpDir = path.join('.tmp', `values-batch-update-tests-${crypto.randomUUID()}`);
      await fs.mkdir(tmpDir, { recursive: true });

      // Get middleware for tool creation
      const middlewareContext = await createMiddlewareContext();
      authProvider = middlewareContext.authProvider;
      logger = middlewareContext.logger;
      const middleware = middlewareContext.middleware;
      accountId = middlewareContext.accountId;
      const tool = createTool();
      const wrappedTool = middleware.withToolAuth(tool);
      handler = wrappedTool.handler;

      // Create shared spreadsheet for all tests
      const title = `ci-values-batch-update-tests-${Date.now()}`;
      sharedSpreadsheetId = await createTestSpreadsheet(await authProvider.getAccessToken(accountId), { title });
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

  it('values_batchUpdate handles RAW vs USER_ENTERED value input options', async () => {
    // Use default sheet (gid: 0) to minimize write operations
    const testSheetId = 0;

    // Test with USER_ENTERED (formulas should be processed)
    const userEnteredRequests = [
      {
        range: 'A1:B2',
        values: [
          ['=1+1', '=TODAY()'],
          ['=SUM(1,2,3)', '=CONCATENATE("Hello", " World")'],
        ],
        majorDimension: 'ROWS' as const,
      },
    ];

    const userEnteredResp = await handler(
      {
        id: sharedSpreadsheetId,
        gid: String(testSheetId),
        requests: userEnteredRequests,
        valueInputOption: 'USER_ENTERED',
        includeData: false,
      },
      createExtra()
    );

    const userEnteredStructured = userEnteredResp.structuredContent?.result as Output | undefined;
    assert.ok(userEnteredStructured, 'Response missing structuredContent.result');
    assert.strictEqual(userEnteredStructured.type, 'success', 'USER_ENTERED should succeed');

    // Test with RAW (formulas should be treated as literal text)
    const rawRequests = [
      {
        range: 'D1:E2',
        values: [
          ['=1+1', '=TODAY()'],
          ['Raw Text', 'Another Text'],
        ],
        majorDimension: 'ROWS' as const,
      },
    ];

    const rawResp = await handler(
      {
        id: sharedSpreadsheetId,
        gid: String(testSheetId),
        requests: rawRequests,
        valueInputOption: 'RAW',
        includeData: false,
      },
      createExtra()
    );

    const rawStructured = rawResp.structuredContent?.result as Output | undefined;
    assert.ok(rawStructured, 'Response missing structuredContent.result');
    assert.strictEqual(rawStructured.type, 'success', 'RAW should succeed');

    // Both should update the same number of cells
    if (userEnteredStructured.type === 'success' && rawStructured.type === 'success') {
      assert.equal(userEnteredStructured.totalUpdatedCells, 4, 'USER_ENTERED should update 4 cells');
      assert.equal(rawStructured.totalUpdatedCells, 4, 'RAW should update 4 cells');
    }
  });

  it('values_batchUpdate handles batch updates with multiple ranges, dimensions, and includeData', async () => {
    // Use default sheet (gid: 0) to minimize write operations
    const testSheetId = 0;

    // Comprehensive test covering:
    // - Single range (1x1)
    // - Multiple ranges with different sizes
    // - ROWS major dimension
    // - COLUMNS major dimension
    // - Complex cell count validation
    // - includeData option (consolidated from previous test1)
    const requests = [
      {
        range: 'A1:A1', // 1x1 = 1 cell
        values: [['Single']],
        majorDimension: 'ROWS' as const,
      },
      {
        range: 'C1:E3', // 3x3 = 9 cells (ROWS)
        values: [
          ['C1', 'D1', 'E1'],
          ['C2', 'D2', 'E2'],
          ['C3', 'D3', 'E3'],
        ],
        majorDimension: 'ROWS' as const,
      },
      {
        range: 'A5:J5', // 1x10 = 10 cells
        values: [['A5', 'B5', 'C5', 'D5', 'E5', 'F5', 'G5', 'H5', 'I5', 'J5']],
        majorDimension: 'ROWS' as const,
      },
      {
        range: 'L1:N2', // 2x3 = 6 cells (COLUMNS)
        values: [
          ['L1', 'L2'], // Column L
          ['M1', 'M2'], // Column M
          ['N1', 'N2'], // Column N
        ],
        majorDimension: 'COLUMNS' as const,
      },
    ];

    const batchUpdateResp = await handler(
      {
        id: sharedSpreadsheetId,
        gid: String(testSheetId),
        requests,
        valueInputOption: 'USER_ENTERED',
        includeData: true, // Test includeData option (consolidated from test1)
      },
      createExtra()
    );

    // Validate complete response structure
    assert.ok(batchUpdateResp, 'Handler returned no result');

    const batchStructured = batchUpdateResp.structuredContent?.result as Output | undefined;
    assert.ok(batchStructured, 'Response missing structuredContent.result');
    assert.strictEqual(batchStructured.type, 'success', 'Expected success result');

    if (batchStructured.type === 'success') {
      // Validate response structure matches ValuesBatchUpdateResponseSchema
      assert.ok(typeof batchStructured.id === 'string', 'Item missing valid id');
      assert.ok(typeof batchStructured.spreadsheetTitle === 'string', 'Item missing valid spreadsheetTitle');
      assert.ok(typeof batchStructured.spreadsheetUrl === 'string', 'Item missing valid spreadsheetUrl');
      assert.ok(typeof batchStructured.sheetTitle === 'string', 'Item missing valid sheetTitle');
      assert.ok(typeof batchStructured.gid === 'string', 'Item missing valid gid');
      assert.ok(typeof batchStructured.sheetUrl === 'string', 'Item missing valid sheetUrl');
      assert.ok(typeof batchStructured.totalUpdatedRows === 'number', 'Item missing valid totalUpdatedRows');
      assert.ok(typeof batchStructured.totalUpdatedColumns === 'number', 'Item missing valid totalUpdatedColumns');
      assert.ok(typeof batchStructured.totalUpdatedCells === 'number', 'Item missing valid totalUpdatedCells');
      assert.ok(Array.isArray(batchStructured.updatedRanges), 'Item missing valid updatedRanges array');

      // Validate content array
      const content = batchUpdateResp.content;
      assert.ok(Array.isArray(content), 'Response missing content array');
      assert.ok(content.length > 0, 'Content array is empty');
      const firstContent = content[0];
      assert.ok(firstContent, 'First content item missing');
      assert.strictEqual(firstContent.type, 'text', 'Content item missing text type');
      if (firstContent.type === 'text') {
        assert.ok(typeof firstContent.text === 'string', 'Content item missing text field');
      }

      // Validate cell counts: 1 + 9 + 10 + 6 = 26 cells total
      assert.equal(batchStructured.totalUpdatedCells, 26, 'Should update 26 cells total across all ranges');
      assert.equal(batchStructured.updatedRanges.length, 4, 'Should have 4 updated ranges');

      // Validate includeData functionality (consolidated from test1)
      assert.ok(batchStructured.updatedData, 'Item should include updatedData when includeData is true');
      assert.ok(Array.isArray(batchStructured.updatedData), 'updatedData should be an array');
      assert.equal(batchStructured.updatedData.length, 4, 'Should have 4 updated data entries matching requests');

      // Validate each updatedData item has required fields
      for (const updatedDataItem of batchStructured.updatedData) {
        assert.ok(typeof updatedDataItem.range === 'string', 'Updated data item missing range');
        assert.ok(['ROWS', 'COLUMNS'].includes(updatedDataItem.majorDimension), 'Updated data item missing valid majorDimension');
        assert.ok(Array.isArray(updatedDataItem.values), 'Updated data item missing values array');
      }
    }
  });
}); // End describe block
