import '../../../lib/env-loader.ts';
import { sheets as sheetsApi } from '@googleapis/sheets';
import { mcp } from '@mcp-z/mcp-sheets';
import type { Logger } from '@mcp-z/oauth-google';
import { ProtocolError, ProtocolErrorCode } from '@mcp-z/server';
import assert from 'assert';
import crypto from 'crypto';
import fs from 'fs/promises';
import * as path from 'path';
import type { Input, Output } from '../../../../src/mcp/tools/sheet-copy.ts';
import { createExtra, type TypedHandler } from '../../../lib/create-extra.ts';
import createMiddlewareContext from '../../../lib/create-middleware-context.ts';
import { addTestTabs, deleteTestTabs, sharedSpreadsheet, type TestTab } from '../../../lib/spreadsheet-helpers.ts';

describe('sheet-copy', () => {
  /**
   * OPTIMIZATION NOTES - sheet-copy.test.ts
   *
   * API calls: 10 total (was 14)
   * - Setup: 3 (createSpreadsheet + createSheet + values.update)
   * - Test 1: 5 (handler:get + handler:3×batchUpdate + verify:get with gridData)
   * - Test 2: 1 (handler:get - fails immediately)
   * - Teardown: 1 (deleteSpreadsheet)
   *
   * Optimizations applied:
   * - Consolidated single copy + batch copy into one test (batch includes single)
   * - Combined spreadsheets.get + values.get into single get with includeGridData
   */

  // Shared test resources
  let sharedSpreadsheetId: string;
  let tab: TestTab;
  const copiedGids: string[] = [];
  let auth: Awaited<ReturnType<typeof createMiddlewareContext>>['auth'];
  let logger: Logger;
  let handler: TypedHandler<Input>;
  let tmpDir: string;
  let sourceSheetId: number;

  before(async () => {
    try {
      // Create temporary directory
      tmpDir = path.join('.tmp', `sheet-copy-tests-${crypto.randomUUID()}`);
      await fs.mkdir(tmpDir, { recursive: true });

      // Get middleware for tool creation and auth for close operations
      const middlewareContext = await createMiddlewareContext();
      auth = middlewareContext.auth;
      logger = middlewareContext.logger;
      const middleware = middlewareContext.middleware;
      const tool = mcp.toolFactories.sheetCopy();
      const wrappedTool = middleware.withToolAuth(tool);
      handler = wrappedTool.handler;

      // Create shared spreadsheet for all tests
      sharedSpreadsheetId = await sharedSpreadsheet();
      [tab] = (await addTestTabs(sharedSpreadsheetId, ['SourceTemplate'])) as [TestTab];

      sourceSheetId = tab.sheetId;

      // Add some data to the source sheet
      const sheets = sheetsApi({ version: 'v4', auth });
      await sheets.spreadsheets.values.update({
        spreadsheetId: sharedSpreadsheetId,
        range: 'SourceTemplate!A1:B2',
        valueInputOption: 'RAW',
        requestBody: {
          values: [
            ['Header1', 'Header2'],
            ['Value1', 'Value2'],
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
    if (tab) await deleteTestTabs(sharedSpreadsheetId, [tab.gid, ...copiedGids]);
    if (tmpDir) await fs.rm(tmpDir, { recursive: true, force: true });
  });

  it('sheet_copy creates multiple copies in batch with data preserved', async () => {
    // Test batch copy (which inherently tests single copy as well)
    const copies = [{ newTitle: `batch-copy-1-${Date.now()}` }, { newTitle: `batch-copy-2-${Date.now()}` }, { newTitle: `batch-copy-3-${Date.now()}` }];

    const res = await handler(
      {
        id: sharedSpreadsheetId,
        gid: String(sourceSheetId),
        copies,
      },
      createExtra()
    );

    assert.ok(res && res.structuredContent && res.content, 'missing structured result for sheet_copy');

    const branch = (res.structuredContent as { result?: unknown } | undefined)?.result as Output | undefined;
    assert.ok(branch, 'missing structured result for sheet_copy');
    assert.equal(branch.type, 'success');

    if (branch.type === 'success') {
      // Validate response structure
      assert.equal(branch.id, sharedSpreadsheetId, 'should return spreadsheet id');
      assert.equal(branch.sourceGid, String(sourceSheetId), 'should return source gid');
      assert.equal(branch.sourceTitle, 'SourceTemplate', 'should return source title');
      assert.equal(branch.itemsProcessed, 3, 'should process 3 items');
      assert.equal(branch.itemsChanged, 3, 'should create 3 sheets');
      assert.equal(branch.items.length, 3, 'should have 3 created sheets');
      copiedGids.push(...branch.items.map((item) => item.gid));

      // Verify first copy has correct title
      const firstCopy = copies[0];
      assert.ok(firstCopy, 'first copy should exist');
      assert.equal(branch.items[0]?.title, firstCopy.newTitle, 'created sheet should have correct title');

      // OPTIMIZATION: Single API call to verify sheets exist AND data was copied
      // Uses includeGridData with range for one sheet to verify data, and fields for sheet list
      const client = sheetsApi({ version: 'v4', auth });
      const info = await client.spreadsheets.get({
        spreadsheetId: sharedSpreadsheetId,
        includeGridData: true,
        ranges: [`'${firstCopy.newTitle}'!A1:B2`],
        fields: 'sheets.properties.title,sheets.data.rowData.values.formattedValue',
      });

      // Verify all sheets were created (check against handler response which lists all created sheets)
      for (const item of branch.items) {
        assert.ok(item.gid, `Sheet "${item.title}" should have a gid`);
        assert.ok(item.sheetUrl.includes(sharedSpreadsheetId), `Sheet "${item.title}" sheetUrl should contain spreadsheet id`);
      }

      // Verify data was copied (from the first sheet's grid data)
      const firstSheetData = info.data.sheets?.find((s) => s?.properties?.title === firstCopy.newTitle);
      assert.ok(firstSheetData, `Sheet "${firstCopy.newTitle}" should exist in response`);
      const gridData = firstSheetData?.data?.[0]?.rowData;
      assert.ok(gridData, 'Grid data should exist');
      assert.equal(gridData[0]?.values?.[0]?.formattedValue, 'Header1', 'First cell should be Header1');
      assert.equal(gridData[0]?.values?.[1]?.formattedValue, 'Header2', 'Second cell should be Header2');
      assert.equal(gridData[1]?.values?.[0]?.formattedValue, 'Value1', 'Third cell should be Value1');
      assert.equal(gridData[1]?.values?.[1]?.formattedValue, 'Value2', 'Fourth cell should be Value2');
    }
  });

  it('sheet_copy fails for non-existent source sheet', async () => {
    const nonExistentGid = '999999999';

    try {
      await handler(
        {
          id: sharedSpreadsheetId,
          gid: nonExistentGid,
          copies: [{ newTitle: 'should-fail' }],
        },
        createExtra()
      );
      assert.fail('Expected error for non-existent source sheet');
    } catch (error) {
      assert.ok(error instanceof Error, 'Should throw an error');
      assert.ok(error.message.includes('not found'), 'Error should mention sheet not found');
    }
  });

  // Input errors reach the client as InvalidParams, not InternalError.
  it('sheet_copy rejects titles that already exist as InvalidParams, naming each failure', async () => {
    // Both titles already exist in the spreadsheet, so Google refuses every duplicate.
    await assert.rejects(handler({ id: sharedSpreadsheetId, gid: String(sourceSheetId), copies: [{ newTitle: 'SourceTemplate' }, { newTitle: 'Sheet1' }] }, createExtra()), (error: unknown) => {
      assert.ok(error instanceof ProtocolError, 'expected a ProtocolError');
      assert.equal(error.code, ProtocolErrorCode.InvalidParams);
      assert.ok(error.message.startsWith('every copy failed'), error.message);
      assert.ok(error.message.includes('every copy failed'), error.message);
      assert.ok(error.message.includes('SourceTemplate'), error.message);
      assert.ok(error.message.includes('Sheet1'), error.message);
      return true;
    });
  });

  it('sheet_copy rejects a missing source sheet as InvalidParams', async () => {
    await assert.rejects(handler({ id: sharedSpreadsheetId, gid: '999999999', copies: [{ newTitle: 'should-fail' }] }, createExtra()), (error: unknown) => error instanceof ProtocolError && error.code === ProtocolErrorCode.InvalidParams && error.message.startsWith('Source sheet with gid'));
  });

  it('sheet_copy still succeeds with failures when only some copies fail', async () => {
    const okTitle = `partial-copy-${Date.now()}`;
    const res = await handler({ id: sharedSpreadsheetId, gid: String(sourceSheetId), copies: [{ newTitle: 'SourceTemplate' }, { newTitle: okTitle }] }, createExtra());
    const branch = (res.structuredContent as { result?: unknown } | undefined)?.result as Output | undefined;
    assert.equal(branch?.type, 'success');
    if (branch?.type === 'success') {
      copiedGids.push(...branch.items.map((item) => item.gid));
      assert.equal(branch.itemsProcessed, 2);
      assert.equal(branch.itemsChanged, 1);
      assert.equal(branch.items[0]?.title, okTitle);
      assert.equal(branch.failures?.length, 1);
      assert.equal(branch.failures?.[0]?.title, 'SourceTemplate');
    }
  });
});
