import '../../../lib/env-loader.ts';
import { auth as googleAuthPlus, sheets as sheetsApi } from '@googleapis/sheets';
import { mcp } from '@mcp-z/mcp-sheets';
import type { LoopbackOAuthProvider } from '@mcp-z/oauth-google';
import assert from 'assert';
import crypto from 'crypto';
import fs from 'fs/promises';
import * as path from 'path';
import type { Output as ColumnsGetOutput, Input } from '../../../../src/mcp/tools/columns-get.ts';
import type { Input as RowsAppendInput, Output as RowsAppendOutput } from '../../../../src/mcp/tools/rows-append.ts';
import { createExtra, type TypedHandler } from '../../../lib/create-extra.ts';
import createMiddlewareContext from '../../../lib/create-middleware-context.ts';
import { addTestTabs, deleteTestTabs, sharedSpreadsheet, type TestTab } from '../../../lib/spreadsheet-helpers.ts';

describe('columns-get', () => {
  let handler: TypedHandler<Input>;
  let rowsAppendHandler: TypedHandler<RowsAppendInput>;
  let sharedSpreadsheetId: string;
  let tab: TestTab;
  const extraGids: number[] = [];
  let auth: Awaited<ReturnType<typeof createMiddlewareContext>>['auth'];
  let authProvider: LoopbackOAuthProvider;
  let accountId: string;
  let tmpDir: string;

  before(async () => {
    try {
      // Create temporary directory
      tmpDir = path.join('.tmp', `columns-get-tests-${crypto.randomUUID()}`);
      await fs.mkdir(tmpDir, { recursive: true });

      // Create middleware and tools
      const middlewareContext = await createMiddlewareContext();
      authProvider = middlewareContext.authProvider;
      auth = middlewareContext.auth;
      const middleware = middlewareContext.middleware;
      accountId = middlewareContext.accountId;

      const tool = mcp.toolFactories.columnsGet();
      const wrappedTool = middleware.withToolAuth(tool);
      handler = wrappedTool.handler;

      const rowsAppendTool = mcp.toolFactories.rowsAppend();
      const wrappedRowsAppendTool = middleware.withToolAuth(rowsAppendTool);
      rowsAppendHandler = wrappedRowsAppendTool.handler;

      // Create shared spreadsheet
      sharedSpreadsheetId = await sharedSpreadsheet();
      [tab] = (await addTestTabs(sharedSpreadsheetId, ['columns-get'])) as [TestTab];
    } catch (error) {
      console.error('Failed to initialize test resources:', error);
      throw error;
    }
  });

  after(async () => {
    // Cleanup resources - fail fast on errors
    if (tab) await deleteTestTabs(sharedSpreadsheetId, [tab.gid, ...extraGids]);
    if (tmpDir) await fs.rm(tmpDir, { recursive: true, force: true });
  });

  it('sheets-columns-get reads columns from sheet with data', async () => {
    // Append rows with headers to create a sheet with data
    const rows = [
      ['john@example.com', 'John Doe', '555-1234'],
      ['jane@example.com', 'Jane Smith', '555-5678'],
    ];
    const headers = ['email', 'name', 'phone'];

    const appendResp = await rowsAppendHandler(
      {
        id: sharedSpreadsheetId,
        gid: tab.gid,
        rows,
        headers,
      },
      createExtra()
    );

    const appendStructured = (appendResp.structuredContent as { result?: unknown } | undefined)?.result as RowsAppendOutput | undefined;
    assert.equal(appendStructured?.type, 'success', 'rows-append expected success type');

    // Get columns
    const resp = await handler(
      {
        id: sharedSpreadsheetId,
        gid: tab.gid,
      },
      createExtra()
    );

    const structured = (resp.structuredContent as { result?: unknown } | undefined)?.result as ColumnsGetOutput | undefined;
    assert.equal(structured?.type, 'success', 'expected success type');
    if (structured?.type === 'success') {
      assert.deepEqual(structured.columns, ['email', 'name', 'phone'], 'columns mismatch');
      assert.equal(structured.isEmpty, false, 'expected isEmpty to be false');
    }
  });

  it('sheets-columns-get returns empty for empty sheet', async () => {
    // Create new empty sheet
    const tempAuth = new googleAuthPlus.OAuth2();
    const accessToken = await authProvider.getAccessToken(accountId);
    tempAuth.setCredentials({ access_token: accessToken });
    const sheets = sheetsApi({ version: 'v4', auth });

    const addSheetResp = await sheets.spreadsheets.batchUpdate({
      spreadsheetId: sharedSpreadsheetId,
      requestBody: {
        requests: [
          {
            addSheet: {
              properties: {
                title: `EmptySheet-${Date.now()}`,
              },
            },
          },
        ],
      },
    });

    const newSheetId = addSheetResp.data.replies?.[0]?.addSheet?.properties?.sheetId;
    if (newSheetId !== undefined && newSheetId !== null) extraGids.push(newSheetId);
    assert.ok(newSheetId !== undefined, 'failed to create empty sheet');

    // Get columns from empty sheet
    const resp = await handler(
      {
        id: sharedSpreadsheetId,
        gid: String(newSheetId),
      },
      createExtra()
    );

    const structured = (resp.structuredContent as { result?: unknown } | undefined)?.result as ColumnsGetOutput | undefined;
    assert.equal(structured?.type, 'success', 'expected success type');
    if (structured?.type === 'success') {
      assert.deepEqual(structured.columns, [], 'expected empty columns');
      assert.equal(structured.isEmpty, true, 'expected isEmpty to be true');
    }
  });

  it('sheets-columns-get uses direct gid lookup', async () => {
    // Uses direct gid (not title) - Sheet1 default has gid=0
    const resp = await handler(
      {
        id: sharedSpreadsheetId,
        gid: '0', // Direct gid lookup
      },
      createExtra()
    );

    const structured = (resp.structuredContent as { result?: unknown } | undefined)?.result as ColumnsGetOutput | undefined;
    assert.equal(structured?.type, 'success', 'expected success type');
    // Sheet1 should have columns from first test
    if (structured?.type === 'success') {
      assert.ok(Array.isArray(structured.columns), 'expected columns array');
    }
  });
});
