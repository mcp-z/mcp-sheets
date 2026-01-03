import type { Logger, LoopbackOAuthProvider } from '@mcp-z/oauth-google';
import assert from 'assert';
import crypto from 'crypto';
import fs from 'fs/promises';
import * as path from 'path';
import createTool, { type Input, type Output } from '../../../../src/mcp/tools/rows-append.ts';
import createRowsGetTool, { type Input as RowsGetInput, type Output as RowsGetOutput } from '../../../../src/mcp/tools/rows-get.ts';
import { createExtra, type TypedHandler } from '../../../lib/create-extra.ts';
import createMiddlewareContext from '../../../lib/create-middleware-context.ts';
import { createTestSpreadsheet, deleteTestSpreadsheet } from '../../../lib/spreadsheet-helpers.ts';

/**
 * OPTIMIZATION: Hybrid Testing with rows-get Integration
 * ==========================================================
 *
 * This test file demonstrates the hybrid testing approach where write operation tests
 * use the rows-get tool for validation instead of googleapis directly.
 *
 * Benefits:
 * 1. Validates BOTH rows-append AND rows-get tools simultaneously
 * 2. Tests real-world tool integration patterns
 * 3. Provides additional rows-get coverage (see rows-get.test.ts for strategy)
 *
 * Pattern: Write with rows-append → Validate with rows-get
 */

describe('rows-append tool (service-backed tests)', () => {
  // Shared test resources
  let sharedSpreadsheetId: string;
  let authProvider: LoopbackOAuthProvider;
  let accountId: string;
  let logger: Logger;
  let handler: TypedHandler<Input>;
  let rowsGetHandler: TypedHandler<RowsGetInput>;
  let tmpDir: string;

  before(async () => {
    try {
      // Create temporary directory
      tmpDir = path.join('.tmp', `rows-append-tests-${crypto.randomUUID()}`);
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
      // Initialize rows-get tool for validation
      const rowsGetTool = createRowsGetTool();
      const wrappedRowsGetTool = middleware.withToolAuth(rowsGetTool);
      rowsGetHandler = wrappedRowsGetTool.handler;

      // Create shared spreadsheet for all tests (use default sheet to save operations)
      const title = `ci-rows-append-tests-${Date.now()}`;
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

  it('[Shared:test1-*,test2-*,test3-*] rows_append basic operations (consolidated)', async () => {
    const testSheetId = 0;
    const headers = ['id', 'name', 'value'];

    // PART 1: Basic append with unique ID prefix (test1)
    const test1Rows = [
      ['test1-1', 'Alice-Test1', '30'],
      ['test1-2', 'Bob-Test1', '25'],
    ];

    const test1Resp = await handler(
      {
        id: sharedSpreadsheetId,
        gid: String(testSheetId),
        rows: test1Rows,
        headers,
      },
      createExtra()
    );

    // Validate complete response structure according to outputSchema
    assert.ok(test1Resp, 'Test1: Handler returned no result');

    const test1Structured = test1Resp.structuredContent?.result as Output | undefined;
    assert.ok(test1Structured, 'Test1: Response missing structuredContent.result');
    if (test1Structured?.type !== 'success') {
      assert.fail('Test1: rows-append operation failed');
    }
    assert.ok(test1Structured, 'Test1: Success result missing item');
    assert.ok(typeof test1Structured.updatedRows === 'number', 'Test1: Item missing valid updatedRows');

    const test1Content = test1Resp.content;
    assert.ok(Array.isArray(test1Content), 'Test1: Response missing content array');
    assert.ok(test1Content.length > 0, 'Test1: Content array is empty');
    const firstContentItem = test1Content[0];
    assert.ok(firstContentItem, 'Test1: First content item is undefined');
    assert.strictEqual(firstContentItem.type, 'text', 'Test1: Content item missing text type');
    if (firstContentItem.type === 'text') {
      assert.ok(typeof firstContentItem.text === 'string', 'Test1: Content item missing text field');
    }

    assert.equal(test1Structured.updatedRows, test1Rows.length, 'Test1: Should append both rows');

    // INTEGRATION TEST: Use rows-get to validate the append operation
    const rowsGetResp = await rowsGetHandler(
      {
        id: sharedSpreadsheetId,
        gid: String(testSheetId),
        range: 'A1:C10',
      },
      createExtra()
    );

    const rowsGetStructured = rowsGetResp.structuredContent?.result as RowsGetOutput | undefined;
    assert.ok(rowsGetStructured, 'Test1: rows-get should return structuredContent.result');
    if (rowsGetStructured.type !== 'success') {
      assert.fail(`Test1: rows-get validation failed: expected success type but got ${rowsGetStructured.type}`);
    }
    assert.strictEqual(rowsGetStructured.type, 'success', 'Test1: rows-get should succeed');
    assert.ok(Array.isArray(rowsGetStructured.rows), 'Test1: rows-get should return rows array');

    const allRows = rowsGetStructured.rows as string[][];
    const test1FilteredRows = allRows.filter((row: string[]) => row[0]?.startsWith('test1-'));
    assert.ok(test1FilteredRows.length >= test1Rows.length, `Test1: Should find at least ${test1Rows.length} test1-* rows`);
    assert.ok(
      test1FilteredRows.some((row: string[]) => row[1] === 'Alice-Test1'),
      'Test1: Should find Alice-Test1'
    );
    assert.ok(
      test1FilteredRows.some((row: string[]) => row[1] === 'Bob-Test1'),
      'Test1: Should find Bob-Test1'
    );

    // PART 2: Deduplication with standard headers (test2)
    const test2Id = `test2-${Date.now()}`;
    const test2Rows = [
      [`${test2Id}-1`, 'Alice-Test2', 'value-a'],
      [`${test2Id}-2`, 'Bob-Test2', 'value-b'],
    ];

    const test2Resp = await handler(
      {
        id: sharedSpreadsheetId,
        gid: String(testSheetId),
        rows: test2Rows,
        headers: headers,
        deduplicateBy: ['id'],
      },
      createExtra()
    );

    assert.ok(test2Resp, 'Test2: Handler returned no result');

    const test2Structured = test2Resp.structuredContent?.result as Output | undefined;
    assert.ok(test2Structured, 'Test2: Response missing structuredContent.result');
    if (test2Structured?.type !== 'success') {
      assert.fail('Test2: rows-append operation failed');
    }
    assert.ok(test2Structured, 'Test2: Success result missing item');
    assert.ok(typeof test2Structured.updatedRows === 'number', 'Test2: Item missing valid updatedRows');
    assert.equal(test2Structured.updatedRows, 2, 'Test2: Should add both rows with deduplication');

    // PART 3: Write headers with standard schema (test3)
    const test3Id = `test3-${Date.now()}`;
    const test3Rows = [
      [`${test3Id}-1`, 'Alice-Test3', 'value-a'],
      [`${test3Id}-2`, 'Bob-Test3', 'value-b'],
    ];

    const test3Resp = await handler(
      {
        id: sharedSpreadsheetId,
        gid: String(testSheetId),
        rows: test3Rows,
        headers: headers,
      },
      createExtra()
    );

    assert.ok(test3Resp, 'Test3: Handler returned no result');

    const test3Structured = test3Resp.structuredContent?.result as Output | undefined;
    assert.ok(test3Structured, 'Test3: Response missing structuredContent.result');
    if (test3Structured?.type !== 'success') {
      assert.fail('Test3: Rows append write headers operation failed');
    }
    assert.ok(test3Structured, 'Test3: Success result missing item');
    assert.ok(typeof test3Structured.updatedRows === 'number', 'Test3: Item missing valid updatedRows');
    assert.equal(test3Structured.updatedRows, 2, 'Test3: Should add only data rows');
  });

  it('[Shared:test4-*] rows_append should deduplicate rows based on deduplicateBy parameter', async () => {
    const testSheetId = 0;

    // Use standard headers with unique ID prefix for isolation
    const testId = `test4-${Date.now()}`;
    const headers = ['id', 'name', 'value'];
    const initialRows = [
      [`${testId}-1`, 'Alice-Test4', 'value-a'],
      [`${testId}-2`, 'Bob-Test4', 'value-b'],
    ];

    const firstWriteResp = await handler(
      {
        id: sharedSpreadsheetId,
        gid: String(testSheetId),
        rows: initialRows,
        headers: headers,
        deduplicateBy: ['id'],
      },
      createExtra()
    );

    const firstWriteStructured = firstWriteResp.structuredContent?.result as Output | undefined;

    if (firstWriteStructured?.type !== 'success') {
      assert.fail('First write failed');
    }
    assert.equal(firstWriteStructured.updatedRows, 2); // Only data rows (headers are metadata)

    // Second write: Try to add the same data plus one new row
    const duplicateRows = [
      [`${testId}-1`, 'Alice-Test4', 'value-a'], // Duplicate - should be skipped
      [`${testId}-2`, 'Bob-Test4', 'value-b'], // Duplicate - should be skipped
      [`${testId}-3`, 'Charlie-Test4', 'value-c'], // New - should be added
    ];

    const secondWriteResp = await handler(
      {
        id: sharedSpreadsheetId,
        gid: String(testSheetId),
        rows: duplicateRows,
        headers: headers,
        deduplicateBy: ['id'],
      },
      createExtra()
    );

    const secondWriteStructured = secondWriteResp.structuredContent?.result as Output | undefined;
    if (secondWriteStructured?.type !== 'success') {
      assert.fail('Second write failed');
    }

    // Should only add 1 new row (Charlie), duplicates should be skipped
    assert.equal(secondWriteStructured.updatedRows, 1, 'Should only add 1 new row (Charlie), duplicates should be skipped');

    // Expected: 2 rows should be skipped (Alice and Bob duplicates)
    assert.equal(secondWriteStructured.rowsSkipped, 2, 'Should report 2 rows were skipped (Alice and Bob duplicates)');
  });

  it('[Shared:test5-*] rows_append multiple runs reproduce workflow issue - deduplication behavior', async () => {
    const testSheetId = 0;

    // Use unique test data prefix to avoid conflicts with other tests
    const testPrefix = `test5-${Date.now()}`;

    // Use minimal data set - 3 rows is sufficient to validate batch deduplication behavior
    const headers = ['id', 'name', 'value'];
    const generateRows = (startId: number, count: number) => {
      const rows = [];
      for (let i = 0; i < count; i++) {
        rows.push([
          `${testPrefix}-item-${startId + i}`,
          `User${(startId + i) % 10}`, // This creates some duplicate names
          `Value-${startId + i}`,
        ]);
      }
      return rows;
    };

    // First run: Add 3 records
    const firstRunRows = generateRows(1, 3);

    const firstRunResp = await handler(
      {
        id: sharedSpreadsheetId,
        gid: String(testSheetId),
        rows: firstRunRows,
        headers: headers,
        deduplicateBy: ['id'],
      },
      createExtra()
    );

    const firstRunStructured = firstRunResp.structuredContent?.result as Output | undefined;
    if (firstRunStructured?.type !== 'success') {
      assert.fail('Rows append test5 first run failed');
    }
    console.log('First run result:', {
      updatedRows: firstRunStructured.updatedRows,
      rowsSkipped: firstRunStructured.rowsSkipped || 0,
    });

    // Second run: Add the SAME 3 records again (simulating workflow behavior)
    const secondRunRows = generateRows(1, 3); // Same IDs as first run

    const secondRunResp = await handler(
      {
        id: sharedSpreadsheetId,
        gid: String(testSheetId),
        rows: secondRunRows,
        headers: headers,
        deduplicateBy: ['id'],
      },
      createExtra()
    );

    const secondRunStructured = secondRunResp.structuredContent?.result as Output | undefined;
    if (secondRunStructured?.type !== 'success') {
      assert.fail('Second run failed');
    }

    console.log('Second run result (should show skips):', {
      updatedRows: secondRunStructured.updatedRows,
      rowsSkipped: secondRunStructured.rowsSkipped || 0,
    });

    // This reproduces the user's issue: should see 0 updatedRows and 3 rowsSkipped
    assert.equal(secondRunStructured.updatedRows, 0, 'Second run should add 0 rows (all duplicates)');
    assert.equal(secondRunStructured.rowsSkipped, 3, 'Second run should skip 3 rows (all duplicates)');

    // Third run: Add 2 duplicates + 1 new (mixed scenario - minimum to validate partial deduplication)
    const thirdRunRows = [
      ...generateRows(1, 2), // 2 duplicates
      ...generateRows(4, 1), // 1 new record
    ];

    const thirdRunResp = await handler(
      {
        id: sharedSpreadsheetId,
        gid: String(testSheetId),
        rows: thirdRunRows,
        headers: headers,
        deduplicateBy: ['id'],
      },
      createExtra()
    );

    const thirdRunStructured = thirdRunResp.structuredContent?.result as Output | undefined;
    if (thirdRunStructured?.type !== 'success') {
      assert.fail('Third run failed');
    }

    console.log('Third run result (mixed scenario):', {
      updatedRows: thirdRunStructured.updatedRows,
      rowsSkipped: thirdRunStructured.rowsSkipped || 0,
    });

    // Should add 1 new row and skip 2 duplicates
    assert.equal(thirdRunStructured.updatedRows, 1, 'Third run should add 1 new row');
    assert.equal(thirdRunStructured.rowsSkipped, 2, 'Third run should skip 2 duplicate rows');
  });

  it('[Shared:test9-*,test10-*] rows_append deduplication edge cases (consolidated)', async () => {
    const testSheetId = 0;
    const headers = ['id', 'name', 'value'];

    // PART 1: Complex deduplication with multiple key columns (test9)
    const test9Id = `test9-${Date.now()}`;
    const test9InitialRows = [
      [`${test9Id}-A`, 'Name1', '100'],
      [`${test9Id}-A`, 'Name2', '200'],
      [`${test9Id}-B`, 'Name1', '300'],
    ];

    const test9FirstResp = await handler(
      {
        id: sharedSpreadsheetId,
        gid: String(testSheetId),
        rows: test9InitialRows,
        headers: headers,
        deduplicateBy: ['id', 'name'],
      },
      createExtra()
    );

    const test9FirstStructured = test9FirstResp.structuredContent?.result as Output | undefined;
    if (test9FirstStructured?.type !== 'success') {
      assert.fail('Test9: Complex deduplication failed');
    }
    assert.equal(test9FirstStructured.updatedRows, 3, 'Test9: Should add 3 initial rows');
    assert.equal(test9FirstStructured.rowsSkipped, 0, 'Test9: Should skip 0 rows initially');

    const test9DuplicateRows = [
      [`${test9Id}-A`, 'Name1', '150'], // Duplicate composite key
      [`${test9Id}-A`, 'Name3', '400'], // New combination
      [`${test9Id}-B`, 'Name1', '350'], // Duplicate composite key
      [`${test9Id}-C`, 'Name1', '500'], // New combination
    ];

    const test9SecondResp = await handler(
      {
        id: sharedSpreadsheetId,
        gid: String(testSheetId),
        rows: test9DuplicateRows,
        headers: headers,
        deduplicateBy: ['id', 'name'],
      },
      createExtra()
    );

    const test9SecondStructured = test9SecondResp.structuredContent?.result as Output | undefined;
    if (test9SecondStructured?.type !== 'success') {
      assert.fail('Test9: Second write failed');
    }
    assert.equal(test9SecondStructured.updatedRows, 2, 'Test9: Should add 2 new combinations');
    assert.equal(test9SecondStructured.rowsSkipped, 2, 'Test9: Should skip 2 duplicates');

    // PART 2: Deduplication with empty/null key values (test10)
    const test10Id = `test10-${Date.now()}`;
    const test10Rows = [
      [`${test10Id}-1`, 'Alice-Test10', '100'],
      ['', 'Bob-Test10', '200'], // Empty key
      [null, 'Charlie-Test10', '300'], // Null key
      [`${test10Id}-2`, 'Dave-Test10', '400'],
    ];

    const test10Resp = await handler(
      {
        id: sharedSpreadsheetId,
        gid: String(testSheetId),
        rows: test10Rows,
        headers: headers,
        deduplicateBy: ['id'],
      },
      createExtra()
    );

    const test10Structured = test10Resp.structuredContent?.result as Output | undefined;
    if (test10Structured?.type !== 'success') {
      assert.fail('Test10: Empty keys operation failed');
    }
    assert.equal(test10Structured.updatedRows, 4, 'Test10: All rows should be added (empty/null keys are unique)');
    assert.equal(test10Structured.rowsSkipped, 0, 'Test10: Should skip 0 rows');
  });
}); // End describe block
