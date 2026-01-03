import type { Logger, LoopbackOAuthProvider } from '@mcp-z/oauth-google';
import assert from 'assert';
import crypto from 'crypto';
import fs from 'fs/promises';
import * as path from 'path';
import createTool, { type Input, type Output } from '../../../../src/mcp/tools/columns-update.ts';
import { createExtra, type TypedHandler } from '../../../lib/create-extra.ts';
import createMiddlewareContext from '../../../lib/create-middleware-context.ts';
import { createTestSpreadsheet, deleteTestSpreadsheet } from '../../../lib/spreadsheet-helpers.ts';

/**
 * OPTIMIZATION STRATEGY - columns-update.test.ts
 *
 * Strategy A: Key-Based Isolation with Unique ID Prefixes
 * =========================================================
 *
 * All tests share the default sheet (gid: 0) but use unique ID prefixes to prevent
 * data interference. The columns-update tool uses exact key matching, so different
 * prefixes ensure complete isolation.
 *
 * Key Isolation Mechanism:
 * - Tool reads ALL sheet data to build keySet (existingKeys)
 * - Uses exact Set.has() matching: "test1-1" ≠ "test2-1" ≠ "test3-1"
 * - When Test 2 updates "test2-1", it won't match Test 1's "test1-1"
 *
 * Test Data Prefixes:
 * - Test 1 [Shared:test1-*]: add-or-update behavior (IDs: test1-1, test1-2, test1-3)
 * - Test 2 [Shared:test2-*]: update-only behavior (IDs: test2-1, test2-2, test2-3)
 * - Test 3 [Shared:test3-*]: add-only behavior (IDs: test3-1, test3-2, test3-3)
 * - Test 4 [Shared:test4-*]: composite key matching (Keys: test4-A+Alice, test4-B+Alice, test4-A+Bob)
 *
 * Note: All tests use the same headers ['id', 'name', 'status'] to allow sharing sheet 0
 *
 * Benefits:
 * - Saves 4 write operations (no createTestSheet calls)
 * - Tests remain isolated through key namespacing
 * - Mimics real multi-tenant data patterns
 *
 * When Adding New Tests:
 * Use the next available prefix (test5-, test6-, etc.) and document it here.
 */

describe('columns-update tool (service-backed tests)', () => {
  // Shared instances for all tests
  let authProvider: LoopbackOAuthProvider;
  let logger: Logger;
  let accountId: string;
  let handler: TypedHandler<Input>;
  let sharedSpreadsheetId: string;

  before(async () => {
    // Get middleware for tool creation
    const middlewareContext = await createMiddlewareContext();
    authProvider = middlewareContext.authProvider;
    logger = middlewareContext.logger;
    const middleware = middlewareContext.middleware;
    accountId = middlewareContext.accountId;
    const tool = createTool();
    const wrappedTool = middleware.withToolAuth(tool);
    handler = wrappedTool.handler;
    sharedSpreadsheetId = await createTestSpreadsheet(await authProvider.getAccessToken(accountId), {
      title: `test-columns-update-${Date.now()}`,
    });
  });

  after(async () => {
    // Cleanup shared spreadsheet (automatically deletes all sheets within it)
    await deleteTestSpreadsheet(await authProvider.getAccessToken(accountId), sharedSpreadsheetId, logger);
  });

  it('[Shared:test1-*,test3-*] performs add-or-update and add-only behaviors (consolidated)', async () => {
    const tmp = path.join('.tmp', `wtsi-columns-behaviors-${crypto.randomUUID()}`);
    await fs.mkdir(tmp, { recursive: true });

    try {
      // Use default sheet (gid: 0) with test1- and test3- prefixes for isolation
      const testSheetId = 0;
      const createdFileId = sharedSpreadsheetId;
      const headers = ['id', 'name', 'status'];

      // PART 1: Test add-or-update behavior with test1- prefix
      // First operation: Add initial data
      const test1InitialRows = [
        ['test1-1', 'Alice', 'active'],
        ['test1-2', 'Bob', 'inactive'],
      ];

      const test1FirstResult = await handler(
        {
          id: createdFileId,
          gid: String(testSheetId),
          rows: test1InitialRows,
          headers: headers,
          updateBy: ['id'],
          behavior: 'add-or-update',
          valueInputOption: 'USER_ENTERED',
        },
        createExtra()
      );

      const test1FirstStructured = test1FirstResult.structuredContent?.result as Output | undefined;
      if (test1FirstStructured?.type !== 'success') {
        assert.fail('Test1 first operation failed');
      }
      assert.equal(test1FirstStructured.updatedRows, 2); // Both rows added
      assert.equal(test1FirstStructured.insertedKeys.length, 2);
      assert.equal(test1FirstStructured.rowsSkipped, 0);

      // Second operation: Update existing and add new
      const test1UpdateRows = [
        ['test1-1', 'Alice Updated', 'active'], // Update existing
        ['test1-3', 'Charlie', 'pending'], // Add new
      ];

      const test1SecondResult = await handler(
        {
          id: createdFileId,
          gid: String(testSheetId),
          rows: test1UpdateRows,
          headers: headers,
          updateBy: ['id'],
          behavior: 'add-or-update',
          valueInputOption: 'USER_ENTERED',
        },
        createExtra()
      );

      const test1SecondStructured = test1SecondResult.structuredContent?.result as Output | undefined;
      if (test1SecondStructured?.type !== 'success') {
        assert.fail('Test1 second operation failed');
      }
      assert.equal(test1SecondStructured.updatedRows, 2); // One updated, one added
      assert.equal(test1SecondStructured.insertedKeys.length, 1); // Only Charlie was new
      assert.equal(test1SecondStructured.rowsSkipped, 0);

      // PART 2: Test add-only behavior with test3- prefix
      // Add initial data
      const test3InitialRows = [
        ['test3-1', 'Alice', 'active'],
        ['test3-2', 'Bob', 'inactive'],
      ];

      const test3InitialResult = await handler(
        {
          id: createdFileId,
          gid: String(testSheetId),
          rows: test3InitialRows,
          headers: headers,
          updateBy: ['id'],
          behavior: 'add-or-update',
          valueInputOption: 'USER_ENTERED',
        },
        createExtra()
      );

      const test3InitialStructured = test3InitialResult.structuredContent?.result as Output | undefined;
      if (test3InitialStructured?.type !== 'success') {
        assert.fail('Test3 initial data write failed');
      }
      assert.equal(test3InitialStructured.updatedRows, 2, 'Test3 should add 2 initial rows');

      // Now test add-only behavior: try to update existing and add new
      const test3UpdateRows = [
        ['test3-1', 'Alice Updated', 'active'], // Should be skipped (existing row)
        ['test3-3', 'Charlie', 'pending'], // Should be added (new row)
      ];

      const test3Result = await handler(
        {
          id: createdFileId,
          gid: String(testSheetId),
          rows: test3UpdateRows,
          headers: headers,
          updateBy: ['id'],
          behavior: 'add-only',
          valueInputOption: 'USER_ENTERED',
        },
        createExtra()
      );

      const test3ResultStructured = test3Result.structuredContent?.result as Output | undefined;
      if (test3ResultStructured?.type !== 'success') {
        assert.fail('Test3 add-only operation failed');
      }
      assert.equal(test3ResultStructured.updatedRows, 1); // Only Charlie added
      assert.equal(test3ResultStructured.insertedKeys.length, 1); // Charlie was new
      assert.equal(test3ResultStructured.rowsSkipped, 1); // Alice skipped
    } finally {
      await fs.rm(tmp, { recursive: true, force: true });
    }
  });

  it('[Shared:test2-*] performs update-only behavior with real Google Sheets (service-backed)', async () => {
    const tmp = path.join('.tmp', `wtsi-columns-update-only-${crypto.randomUUID()}`);
    await fs.mkdir(tmp, { recursive: true });

    try {
      // Use default sheet (gid: 0) with test2- prefix for isolation
      const testSheetId = 0;
      const createdFileId = sharedSpreadsheetId;

      const headers = ['id', 'name', 'status'];
      const initialRows = [
        ['test2-1', 'Alice', 'active'],
        ['test2-2', 'Bob', 'inactive'],
      ];

      // Add initial data
      await handler(
        {
          id: createdFileId,
          gid: String(testSheetId),
          rows: initialRows,
          headers: headers,
          updateBy: ['id'],
          behavior: 'add-or-update',
          valueInputOption: 'USER_ENTERED',
        },
        createExtra()
      );

      // Now test update-only behavior: try to update existing and add new (with test2- prefix)
      const updateRows = [
        ['test2-1', 'Alice Updated', 'active'], // Should update existing
        ['test2-3', 'Charlie', 'pending'], // Should be skipped (new row)
      ];

      const result = await handler(
        {
          id: createdFileId,
          gid: String(testSheetId),
          rows: updateRows,
          headers: headers,
          updateBy: ['id'],
          behavior: 'update-only',
          valueInputOption: 'USER_ENTERED',
        },
        createExtra()
      );

      const resultStructured = result.structuredContent?.result as Output | undefined;
      if (resultStructured?.type !== 'success') {
        assert.fail('Update-only operation failed');
      }
      assert.equal(resultStructured.updatedRows, 1); // Only Alice updated
      assert.equal(resultStructured.insertedKeys.length, 0); // No new insertions
      assert.equal(resultStructured.rowsSkipped, 1); // Charlie skipped
    } finally {
      await fs.rm(tmp, { recursive: true, force: true });
    }
  });

  it('[Shared:test4-*] handles composite key matching correctly (service-backed)', async () => {
    const tmp = path.join('.tmp', `wtsi-columns-composite-key-${crypto.randomUUID()}`);
    await fs.mkdir(tmp, { recursive: true });

    try {
      // Use default sheet (gid: 0) with test4- prefix for isolation (composite key: id + name)
      const testSheetId = 0;
      const createdFileId = sharedSpreadsheetId;

      const headers = ['id', 'name', 'status'];
      const initialRows = [
        ['test4-A', 'Alice', 'active'],
        ['test4-B', 'Alice', 'inactive'], // Same name, different id
        ['test4-A', 'Bob', 'pending'],
      ];

      // Add initial data with composite key (id + name)
      const firstResult = await handler(
        {
          id: createdFileId,
          gid: String(testSheetId),
          rows: initialRows,
          headers: headers,
          updateBy: ['id', 'name'],
          behavior: 'add-or-update',
          valueInputOption: 'USER_ENTERED',
        },
        createExtra()
      );

      const firstStructured = firstResult.structuredContent?.result as Output | undefined;
      if (firstStructured?.type !== 'success') {
        assert.fail('First write failed');
      }
      assert.equal(firstStructured.updatedRows, 3);

      // Update with composite key matching (with test4- prefix)
      const updateRows = [
        ['test4-A', 'Alice', 'updated'], // Update test4-A+Alice
        ['test4-B', 'Bob', 'new'], // Add test4-B+Bob (new combination)
      ];

      const secondResult = await handler(
        {
          id: createdFileId,
          gid: String(testSheetId),
          rows: updateRows,
          headers: headers,
          updateBy: ['id', 'name'],
          behavior: 'add-or-update',
          valueInputOption: 'USER_ENTERED',
        },
        createExtra()
      );

      const secondStructured = secondResult.structuredContent?.result as Output | undefined;
      if (secondStructured?.type !== 'success') {
        assert.fail('Second composite key operation failed');
      }
      assert.equal(secondStructured.updatedRows, 2); // test4-A+Alice updated, test4-B+Bob added
      assert.equal(secondStructured.insertedKeys.length, 1); // Only test4-B+Bob was new
    } finally {
      await fs.rm(tmp, { recursive: true, force: true });
    }
  });

  it('prevents silent data loss with empty key values (service-backed)', async () => {
    const headers = ['id', 'name', 'status'];
    const rowsWithEmptyKeys = [
      ['', 'Alice', 'active'], // Empty key
      ['2', 'Bob', 'inactive'],
    ];

    await assert.rejects(
      async () => {
        await handler(
          {
            id: 'any-spreadsheet-id', // Won't matter since validation fails early
            gid: '0',
            rows: rowsWithEmptyKeys,
            headers: headers,
            updateBy: ['id'],
            behavior: 'add-or-update',
            valueInputOption: 'USER_ENTERED',
          },
          createExtra()
        );
      },
      (error: unknown) => {
        assert.ok(error instanceof Error && error.message.includes('Silent data loss prevented'), 'expected "Silent data loss prevented" in error');
        assert.ok(error instanceof Error && error.message.includes('empty key'), 'expected "empty key" in error');
        return true;
      }
    );
  });

  it('handles invalid spreadsheet ID gracefully', async () => {
    const invalidSpreadsheetId = 'invalid-id-that-does-not-exist';
    const headers = ['id', 'name'];
    const rows = [['1', 'test']];

    await assert.rejects(
      async () => {
        await handler(
          {
            id: invalidSpreadsheetId,
            gid: '0',
            rows: rows,
            headers: headers,
            updateBy: ['id'],
            behavior: 'add-or-update',
            valueInputOption: 'USER_ENTERED',
          },
          createExtra()
        );
      },
      (error: unknown) => {
        // Error message should indicate spreadsheet not found or permission denied
        assert.ok(error instanceof Error && (error.message.includes('not found') || error.message.includes('Spreadsheet not found') || error.message.includes('permission') || error.message.includes('PERMISSION_DENIED')), 'expected error about spreadsheet not found or permission denied');
        return true;
      }
    );
  });
});
