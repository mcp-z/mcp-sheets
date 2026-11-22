import '../../../lib/env-loader.ts';
import { sheets as sheetsApi } from '@googleapis/sheets';
import { mcp } from '@mcp-z/mcp-sheets';
import type { Logger } from '@mcp-z/oauth-google';
import { ProtocolError, ProtocolErrorCode } from '@mcp-z/server';
import assert from 'assert';
import crypto from 'crypto';
import fs from 'fs/promises';
import * as path from 'path';
import type { Input, Output } from '../../../../src/mcp/tools/sheet-delete.ts';
import { createExtra, type TypedHandler } from '../../../lib/create-extra.ts';
import createMiddlewareContext from '../../../lib/create-middleware-context.ts';
import { deleteTestTabs, sharedSpreadsheet } from '../../../lib/spreadsheet-helpers.ts';

/**
 * OPTIMIZATION NOTES - sheet-delete.test.ts
 *
 * API calls: 5 total (was 6)
 * - Setup: 1 (createSpreadsheet)
 * - Test 1: 3 (batchUpdate:addSheet + handler:batchUpdate:deleteSheet + verify:get)
 * - Teardown: 1 (deleteSpreadsheet)
 *
 * Optimizations applied:
 * - Removed "verify before" check - we just created the sheet, we know it exists
 */

describe('sheet-delete tool (service-backed tests)', () => {
  // Shared test resources
  let sharedSpreadsheetId: string;
  const addedGids: string[] = [];
  let auth: Awaited<ReturnType<typeof createMiddlewareContext>>['auth'];
  let logger: Logger;
  let handler: TypedHandler<Input>;
  let tmpDir: string;

  before(async () => {
    try {
      // Create temporary directory
      tmpDir = path.join('.tmp', `sheet-delete-tests-${crypto.randomUUID()}`);
      await fs.mkdir(tmpDir, { recursive: true });

      // Get middleware for tool creation
      const middlewareContext = await createMiddlewareContext();
      logger = middlewareContext.logger;
      auth = middlewareContext.auth;
      const middleware = middlewareContext.middleware;
      const tool = mcp.toolFactories.sheetDelete();
      const wrappedTool = middleware.withToolAuth(tool);
      handler = wrappedTool.handler;

      // Create shared spreadsheet for all tests
      sharedSpreadsheetId = await sharedSpreadsheet();
    } catch (error) {
      logger.error('Failed to initialize test resources:', { error });
      throw error;
    }
  });

  after(async () => {
    // Tabs the tests deleted through the tool are already gone; only a failed test leaves one.
    if (sharedSpreadsheetId) {
      const sheets = sheetsApi({ version: 'v4', auth });
      const info = await sheets.spreadsheets.get({ spreadsheetId: sharedSpreadsheetId, fields: 'sheets.properties.sheetId' });
      const present = new Set((info.data.sheets ?? []).map((s) => String(s.properties?.sheetId)));
      await deleteTestTabs(
        sharedSpreadsheetId,
        addedGids.filter((gid) => present.has(gid))
      );
    }
    if (tmpDir) await fs.rm(tmpDir, { recursive: true, force: true });
  });

  it('sheet_delete creates and deletes a sheet successfully', async () => {
    // Step 1: Create a new sheet to delete
    const sheets = sheetsApi({ version: 'v4', auth: auth });
    const addSheetResponse = await sheets.spreadsheets.batchUpdate({
      spreadsheetId: sharedSpreadsheetId,
      requestBody: {
        requests: [
          {
            addSheet: {
              properties: {
                title: `test-sheet-to-delete-${Date.now()}`,
              },
            },
          },
        ],
      },
    });

    const newSheetGid = addSheetResponse.data.replies?.[0]?.addSheet?.properties?.sheetId;
    assert.ok(newSheetGid !== undefined, 'Should have created a new sheet with a GID');
    addedGids.push(String(newSheetGid));

    // OPTIMIZATION: Skip "verify before" - we just created it, we know it exists
    // The addSheetResponse already confirmed creation was successful

    // Step 2: Delete the sheet using the tool
    const deleteResponse = await handler(
      {
        id: sharedSpreadsheetId,
        gids: [String(newSheetGid)],
      },
      createExtra()
    );

    // Validate complete response structure according to outputSchema
    assert.ok(deleteResponse, 'Handler returned no result');

    // Validate structuredContent.result matches outputSchema
    const deleteStructured = (deleteResponse.structuredContent as { result?: unknown } | undefined)?.result as Output | undefined;
    assert.ok(deleteStructured, 'Response missing structuredContent.result');
    assert.strictEqual(deleteStructured.type, 'success', 'Expected success result');

    // Validate response structure matches SheetDeleteResponseSchema
    if (deleteStructured.type === 'success') {
      assert.ok(typeof deleteStructured.id === 'string', 'Item missing valid id');
      assert.ok(typeof deleteStructured.spreadsheetUrl === 'string', 'Item missing valid spreadsheetUrl');
      assert.ok(typeof deleteStructured.operationSummary === 'string', 'Item missing valid operationSummary');
      assert.ok(typeof deleteStructured.itemsProcessed === 'number', 'Item missing valid itemsProcessed');
      assert.ok(typeof deleteStructured.itemsChanged === 'number', 'Item missing valid itemsChanged');
      assert.ok(typeof deleteStructured.completedAt === 'string', 'Item missing valid completedAt');
      assert.strictEqual(deleteStructured.recoverable, false, 'Delete operation should NOT be recoverable');

      // Validate content array matches outputSchema requirements
      const content = deleteResponse.content;
      assert.ok(Array.isArray(content), 'Response missing content array');
      assert.ok(content.length > 0, 'Content array is empty');
      const firstContent = content[0];
      assert.ok(firstContent, 'First content item missing');
      assert.strictEqual(firstContent.type, 'text', 'Content item missing text type');
      if (firstContent.type === 'text') {
        assert.ok(typeof firstContent.text === 'string', 'Content item missing text field');
      }

      // Validate deletion details
      assert.equal(deleteStructured.itemsProcessed, 1, 'Should have processed 1 sheet');
      assert.equal(deleteStructured.itemsChanged, 1, 'Should have deleted 1 sheet');
      assert.strictEqual(deleteStructured.failures, undefined, 'Should have no failures for successful deletion');
    }

    // Step 3: Verify the sheet no longer exists
    const afterDelete = await sheets.spreadsheets.get({
      spreadsheetId: sharedSpreadsheetId,
    });
    const sheetsAfterDelete = afterDelete.data.sheets || [];
    const sheetStillExists = sheetsAfterDelete.some((s) => s.properties?.sheetId === newSheetGid);
    assert.ok(!sheetStillExists, 'Sheet should NOT exist after deletion');

    // The spreadsheet should have at least the default sheet remaining
    assert.ok(sheetsAfterDelete.length >= 1, 'Spreadsheet should have at least the default sheet');
  });

  // Input errors reach the client as InvalidParams, not InternalError.
  it('sheet_delete rejects gids that do not exist as InvalidParams, naming each failure', async () => {
    // Neither gid exists in the spreadsheet, so Google refuses every delete.
    await assert.rejects(handler({ id: sharedSpreadsheetId, gids: ['999999991', '999999992'] }, createExtra()), (error: unknown) => {
      assert.ok(error instanceof ProtocolError, 'expected a ProtocolError');
      assert.equal(error.code, ProtocolErrorCode.InvalidParams);
      assert.ok(error.message.startsWith('every delete failed'), error.message);
      assert.ok(error.message.includes('every delete failed'), error.message);
      assert.ok(error.message.includes('999999991'), error.message);
      assert.ok(error.message.includes('999999992'), error.message);
      return true;
    });
  });

  it('sheet_delete still succeeds with failures when only some deletes fail', async () => {
    const sheets = sheetsApi({ version: 'v4', auth: auth });
    const added = await sheets.spreadsheets.batchUpdate({
      spreadsheetId: sharedSpreadsheetId,
      requestBody: { requests: [{ addSheet: { properties: { title: `ci-partial-delete-${Date.now()}` } } }] },
    });
    const gid = String(added.data.replies?.[0]?.addSheet?.properties?.sheetId);
    addedGids.push(gid);

    const res = await handler({ id: sharedSpreadsheetId, gids: [gid, '999999991'] }, createExtra());
    const branch = (res.structuredContent as { result?: unknown } | undefined)?.result as Output | undefined;
    assert.equal(branch?.type, 'success');
    if (branch?.type === 'success') {
      assert.equal(branch.itemsProcessed, 2);
      assert.equal(branch.itemsChanged, 1);
      assert.equal(branch.failures?.length, 1);
      assert.equal(branch.failures?.[0]?.gid, '999999991');
    }
  });
});
