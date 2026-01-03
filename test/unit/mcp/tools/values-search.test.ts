import type { Logger, LoopbackOAuthProvider } from '@mcp-z/oauth-google';
import assert from 'assert';
import crypto from 'crypto';
import fs from 'fs/promises';
import type { OAuth2Client } from 'google-auth-library';
import { google } from 'googleapis';
import * as path from 'path';
import createSearchTool, { type Input, type Output } from '../../../../src/mcp/tools/values-search.ts';
import { createExtra, type TypedHandler } from '../../../lib/create-extra.ts';
import createMiddlewareContext from '../../../lib/create-middleware-context.ts';
import { createTestSpreadsheet, deleteTestSpreadsheet } from '../../../lib/spreadsheet-helpers.ts';

/**
 * OPTIMIZATION STRATEGY - search.test.ts
 *
 * Strategy: Prefix-Based Data Isolation with Single Batch Write
 * =============================================================
 *
 * All tests share the default sheet (gid: 0) and use unique data prefixes to prevent
 * search interference. All test data is written in a single batch operation in before(),
 * saving 5 write operations.
 *
 * Test Data Prefixes:
 * - Test 1 [search1-*]: Basic count test (search1-Alice, search1-Bob)
 * - Test 2 [search2-*]: Cells with a1s flag (search2-Test)
 * - Test 3 [search3-*]: Cells with values flag (search3-Alice, search3-Bob)
 * - Test 4 [search4-*]: Rows with a1s and values (search4-Alice, search4-Bob)
 * - Test 5 [search5-*]: Columns granularity (search5-Alice with Age column)
 * - Test 6 [search6-*]: Empty query test (search6-A, search6-B)
 *
 * Benefits:
 * - Saves 5 write operations (6 individual appends → 1 batch write)
 * - Tests remain isolated through prefix namespacing
 * - All tests can run in parallel without data interference
 */

describe('search tool (service-backed tests)', () => {
  // Shared test resources
  let sharedSpreadsheetId: string;
  let auth: OAuth2Client;
  let authProvider: LoopbackOAuthProvider;
  let logger: Logger;
  let accountId: string;
  let handler: TypedHandler<Input>;
  let tmpDir: string;

  before(async () => {
    tmpDir = path.join('.tmp', `search-tests-${crypto.randomUUID()}`);
    await fs.mkdir(tmpDir, { recursive: true });
    const middlewareContext = await createMiddlewareContext();
    authProvider = middlewareContext.authProvider;
    logger = middlewareContext.logger;
    auth = middlewareContext.auth;
    const middleware = middlewareContext.middleware;
    accountId = middlewareContext.accountId;
    const tool = createSearchTool();
    const wrappedTool = middleware.withToolAuth(tool);
    handler = wrappedTool.handler;
    const title = `ci-search-tests-${Date.now()}`;
    sharedSpreadsheetId = await createTestSpreadsheet(await authProvider.getAccessToken(accountId), { title });

    // OPTIMIZATION: Add all test data in a single batch write using googleapis directly
    // This saves tool overhead (schema validation, resolution, response formatting)
    // Each test uses a unique prefix to isolate its data
    try {
      const sheets = google.sheets({ version: 'v4', auth });
      await sheets.spreadsheets.values.append({
        spreadsheetId: sharedSpreadsheetId,
        range: 'Sheet1!A1',
        valueInputOption: 'RAW',
        insertDataOption: 'INSERT_ROWS',
        requestBody: {
          values: [
            // Headers
            ['Name', 'Value', 'Age', 'City'],
            // Test 1 data (count only)
            ['search1-Alice', '100', '', ''],
            ['search1-Bob', '200', '', ''],
            // Test 2 data (cells with a1s)
            ['search2-Test', '123', '', ''],
            // Test 3 data (cells with values)
            ['search3-Alice', '', '', ''],
            ['search3-Bob', '', '', ''],
            // Test 4 data (rows with a1s and values)
            ['search4-Alice', '', '30', ''],
            ['search4-Bob', '', '25', ''],
            // Test 5 data (columns granularity)
            ['search5-Alice', '', '30', 'NYC'],
            // Test 6 data (empty query)
            ['search6-A', '1', '', ''],
            ['search6-B', '2', '', ''],
          ],
        },
      });
    } catch (error) {
      // FAIL FAST: Throw to skip all tests if data setup fails
      throw new Error(`Failed to write test data in before() hook: ${error instanceof Error ? error.message : String(error)}. All tests will be skipped.`);
    }
  });

  after(async () => {
    // Cleanup resources - fail fast on errors
    const accessToken = await authProvider.getAccessToken(accountId);
    await deleteTestSpreadsheet(accessToken, sharedSpreadsheetId, logger);
    await fs.rm(tmpDir, { recursive: true, force: true });
  });

  it('[search1-*] sheets-search returns count only when no flags specified', async () => {
    // Use default sheet (gid: 0) - data already loaded in before() with search1- prefix
    const gid = 0;

    const result = await handler({ id: sharedSpreadsheetId, gid: String(gid), select: 'cells' }, createExtra());

    const branch = result.structuredContent?.result as Output | undefined;
    assert.equal(branch?.type, 'success', 'should succeed');

    if (branch?.type === 'success') {
      assert.ok(typeof branch.count === 'number', 'should have count');
      assert.ok(!branch.a1s, 'should not have a1s when not requested');
      assert.ok(!branch.values, 'should not have values when not requested');
    }
  });

  it('[search2-*] sheets-search cells granularity with a1s flag', async () => {
    // Use default sheet (gid: 0) - data already loaded in before() with search2- prefix
    const gid = 0;

    const result = await handler({ id: sharedSpreadsheetId, gid: String(gid), query: 'search2-Test', select: 'cells', a1s: true }, createExtra());

    const branch = result.structuredContent?.result as Output | undefined;
    assert.equal(branch?.type, 'success', 'should succeed');

    if (branch?.type === 'success') {
      assert.ok(branch.count > 0, 'should find matches');
      assert.ok(Array.isArray(branch.a1s), 'should have a1s array');
      if (branch.a1s) {
        assert.ok(
          branch.a1s.every((a1: string) => /^[A-Z]+\d+$/.test(a1)),
          'a1s should be in A1 notation'
        );
      }
    }
  });

  it('[search3-*] sheets-search cells granularity with values flag', async () => {
    // Use default sheet (gid: 0) - data already loaded in before() with search3- prefix
    const gid = 0;

    const result = await handler({ id: sharedSpreadsheetId, gid: String(gid), query: 'search3-Alice', select: 'cells', values: true }, createExtra());

    const branch = result.structuredContent?.result as Output | undefined;
    assert.equal(branch?.type, 'success', 'should succeed');

    if (branch?.type === 'success') {
      assert.ok(Array.isArray(branch.values), 'should have values array');
      if (branch.values) {
        assert.ok(branch.values.includes('search3-Alice'), 'should include matched value');
      }
    }
  });

  it('[search4-*] sheets-search rows granularity with a1s and values', async () => {
    // Use default sheet (gid: 0) - data already loaded in before() with search4- prefix
    const gid = 0;

    const result = await handler({ id: sharedSpreadsheetId, gid: String(gid), query: 'search4-Alice', select: 'rows', a1s: true, values: true }, createExtra());

    const branch = result.structuredContent?.result as Output | undefined;
    assert.equal(branch?.type, 'success', 'should succeed');

    if (branch?.type === 'success') {
      assert.ok(branch.count > 0, 'should find matching rows');
      assert.ok(Array.isArray(branch.a1s), 'should have a1s array');
      if (branch.a1s) {
        assert.ok(
          branch.a1s.every((a1: string) => a1.includes(':')),
          'row a1s should be ranges'
        );
      }
      assert.ok(Array.isArray(branch.values), 'should have values array');
      if (branch.values) {
        assert.ok(
          branch.values.some((row: unknown) => Array.isArray(row)),
          'values should be arrays for rows'
        );
      }
    }
  });

  it('[search5-*] sheets-search columns granularity', async () => {
    // Use default sheet (gid: 0) - data already loaded in before() with search5- prefix
    // This test searches for "Age" column header (not the data, so no prefix needed)
    const gid = 0;

    const result = await handler({ id: sharedSpreadsheetId, gid: String(gid), query: 'Age', select: 'columns', a1s: true }, createExtra());

    const branch = result.structuredContent?.result as Output | undefined;
    assert.equal(branch?.type, 'success', 'should succeed');

    if (branch?.type === 'success') {
      assert.ok(branch.count > 0, 'should find matching columns');
      assert.ok(Array.isArray(branch.a1s), 'should have a1s array');
      if (branch.a1s) {
        assert.ok(
          branch.a1s.every((a1: string) => a1.match(/^[A-Z]+:[A-Z]+$/)),
          'column a1s should be full column notation'
        );
      }
    }
  });

  it('[search6-*] sheets-search with empty query returns all cells', async () => {
    // Use default sheet (gid: 0) - data already loaded in before() with search6- prefix
    const gid = 0;

    const result = await handler({ id: sharedSpreadsheetId, gid: String(gid), select: 'cells' }, createExtra());

    const branch = result.structuredContent?.result as Output | undefined;
    assert.equal(branch?.type, 'success', 'should succeed');

    if (branch?.type === 'success') {
      assert.ok(branch.count > 0, 'should return all cells when query is empty');
    }
  });
}); // End describe block
