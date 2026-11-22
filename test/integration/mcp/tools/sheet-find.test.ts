import '../../../lib/env-loader.ts';
import { mcp } from '@mcp-z/mcp-sheets';
import assert from 'assert';
import crypto from 'crypto';
import fs from 'fs/promises';
import * as path from 'path';
import type { Input as SheetCreateInput, Output as SheetCreateOutput } from '../../../../src/mcp/tools/sheet-create.ts';
import type { Input, Output as SheetFindOutput } from '../../../../src/mcp/tools/sheet-find.ts';
import { createExtra, type TypedHandler } from '../../../lib/create-extra.ts';
import createMiddlewareContext from '../../../lib/create-middleware-context.ts';
import { deleteTestTabs, sharedSpreadsheet } from '../../../lib/spreadsheet-helpers.ts';

describe('sheet-find', () => {
  // Shared test resources
  let sharedSpreadsheetId: string;
  const createdGids: string[] = [];
  let sheetCreateHandler: TypedHandler<SheetCreateInput>;
  let handler: TypedHandler<Input>;
  let tmpDir: string;

  before(async () => {
    try {
      // Create temporary directory
      tmpDir = path.join('.tmp', `sheet-find-tests-${crypto.randomUUID()}`);
      await fs.mkdir(tmpDir, { recursive: true });

      // Create shared context and clients
      const middlewareContext = await createMiddlewareContext();
      const middleware = middlewareContext.middleware;
      const tool = mcp.toolFactories.sheetCreate();
      const wrappedtool = middleware.withToolAuth(tool);
      sheetCreateHandler = wrappedtool.handler;
      const toolFind = mcp.toolFactories.sheetFind();
      const wrappedTool = middleware.withToolAuth(toolFind);
      handler = wrappedTool.handler;

      // Create shared spreadsheet for all tests
      sharedSpreadsheetId = await sharedSpreadsheet();
    } catch (error) {
      console.error('Failed to initialize test resources:', error);
      throw error;
    }
  });

  after(async () => {
    // Cleanup resources - fail fast on errors
    if (createdGids.length > 0) await deleteTestTabs(sharedSpreadsheetId, createdGids);
    if (tmpDir) await fs.rm(tmpDir, { recursive: true, force: true });
  });

  it('sheets-sheet-find locates a sheet by title in a spreadsheet', async () => {
    // Create a new sheet in shared spreadsheet
    const sheetTitle = `InitTab-${Date.now()}`;
    const sheetCreateResp = await sheetCreateHandler({ id: sharedSpreadsheetId, sheetTitle }, createExtra());
    const sheetCreateStructured = (sheetCreateResp.structuredContent as { result?: unknown } | undefined)?.result as SheetCreateOutput | undefined;
    assert.ok(sheetCreateStructured?.type === 'success', 'sheet create not ok');
    createdGids.push(sheetCreateStructured.gid);

    // Find the sheet by title
    const findResp = await handler({ id: sharedSpreadsheetId, sheetRef: sheetTitle }, createExtra());
    const findStructured = (findResp.structuredContent as { result?: unknown } | undefined)?.result as SheetFindOutput | undefined;
    if (findStructured?.type !== 'success') {
      assert.fail('Find sheet operation failed');
    }
    if (findStructured.type === 'success') {
      assert.equal(findStructured.id, sharedSpreadsheetId, 'sheet find id mismatch');
      assert.equal(findStructured.title, sheetTitle, 'sheet find title mismatch');
      assert.ok(typeof findStructured.gid === 'string', 'expected gid on item');
      assert.ok(typeof findStructured.sheetUrl === 'string' && findStructured.sheetUrl.includes('#gid='), 'sheet url missing gid');
    }

    // Not found path - should throw exception
    await assert.rejects(
      async () => {
        await handler({ id: sharedSpreadsheetId, sheetRef: 'NoSuchTab' }, createExtra());
      },
      (error: unknown) => {
        assert.ok(error instanceof Error && error.message.includes('Sheet not found'), 'expected "Sheet not found" error message');
        return true;
      }
    );
  });
});
