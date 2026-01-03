import type { Logger, LoopbackOAuthProvider } from '@mcp-z/oauth-google';
import assert from 'assert';
import crypto from 'crypto';
import fs from 'fs/promises';
import * as path from 'path';
import createTool, { type Input as SheetCreateInput, type Output as SheetCreateOutput } from '../../../../src/mcp/tools/sheet-create.ts';
import createSheetFindTool, { type Input, type Output as SheetFindOutput } from '../../../../src/mcp/tools/sheet-find.ts';
import { createExtra, type TypedHandler } from '../../../lib/create-extra.ts';
import createMiddlewareContext from '../../../lib/create-middleware-context.ts';
import { createTestSpreadsheet, deleteTestSpreadsheet } from '../../../lib/spreadsheet-helpers.ts';

// Shared test resources
let sharedSpreadsheetId: string;
let authProvider: LoopbackOAuthProvider;
let logger: Logger;
let accountId: string;
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
    authProvider = middlewareContext.authProvider;
    logger = middlewareContext.logger;
    const middleware = middlewareContext.middleware;
    accountId = middlewareContext.accountId;
    const tool = createTool();
    const wrappedtool = middleware.withToolAuth(tool);
    sheetCreateHandler = wrappedtool.handler;
    const toolFind = createSheetFindTool();
    const wrappedTool = middleware.withToolAuth(toolFind);
    handler = wrappedTool.handler;

    // Create shared spreadsheet for all tests
    const title = `ci-sheet-find-tests-${Date.now()}`;
    sharedSpreadsheetId = await createTestSpreadsheet(await authProvider.getAccessToken(accountId), { title });
  } catch (error) {
    console.error('Failed to initialize test resources:', error);
    throw error;
  }
});

after(async () => {
  // Cleanup resources - fail fast on errors
  const accessToken = await authProvider.getAccessToken(accountId);
  await deleteTestSpreadsheet(accessToken, sharedSpreadsheetId, logger);
  await fs.rm(tmpDir, { recursive: true, force: true });
});

it('sheets-sheet-find locates a sheet by title in a spreadsheet', async () => {
  // Create a new sheet in shared spreadsheet
  const sheetTitle = `InitTab-${Date.now()}`;
  const sheetCreateResp = await sheetCreateHandler({ id: sharedSpreadsheetId, sheetTitle }, createExtra());
  const sheetCreateStructured = sheetCreateResp.structuredContent?.result as SheetCreateOutput | undefined;
  assert.ok(sheetCreateStructured?.type === 'success', 'sheet create not ok');

  // Find the sheet by title
  const findResp = await handler({ id: sharedSpreadsheetId, sheetRef: sheetTitle }, createExtra());
  const findStructured = findResp.structuredContent?.result as SheetFindOutput | undefined;
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
