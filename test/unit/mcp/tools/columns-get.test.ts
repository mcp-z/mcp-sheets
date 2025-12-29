import type { Logger, LoopbackOAuthProvider } from '@mcp-z/oauth-google';
import assert from 'assert';
import crypto from 'crypto';
import fs from 'fs/promises';
import { OAuth2Client } from 'google-auth-library';
import { google } from 'googleapis';
import * as path from 'path';
import createColumnsGetTool, { type Output as ColumnsGetOutput, type Input } from '../../../../src/mcp/tools/columns-get.js';
import createRowsAppendTool, { type Input as RowsAppendInput, type Output as RowsAppendOutput } from '../../../../src/mcp/tools/rows-append.js';
import { createExtra, type TypedHandler } from '../../../lib/create-extra.js';
import createMiddlewareContext from '../../../lib/create-middleware-context.js';
import { createTestSpreadsheet, deleteTestSpreadsheet } from '../../../lib/spreadsheet-helpers.js';

let handler: TypedHandler<Input>;
let rowsAppendHandler: TypedHandler<RowsAppendInput>;
let sharedSpreadsheetId: string;
let auth: OAuth2Client;
let authProvider: LoopbackOAuthProvider;
let accountId: string;
let logger: Logger;
let tmpDir: string;

before(async () => {
  try {
    // Create temporary directory
    tmpDir = path.join('.tmp', `columns-get-tests-${crypto.randomUUID()}`);
    await fs.mkdir(tmpDir, { recursive: true });

    // Create middleware and tools
    const middlewareContext = await createMiddlewareContext();
    authProvider = middlewareContext.authProvider;
    logger = middlewareContext.logger;
    auth = middlewareContext.auth;
    const middleware = middlewareContext.middleware;
    accountId = middlewareContext.accountId;

    const tool = createColumnsGetTool();
    const wrappedTool = middleware.withToolAuth(tool);
    handler = wrappedTool.handler;

    const rowsAppendTool = createRowsAppendTool();
    const wrappedRowsAppendTool = middleware.withToolAuth(rowsAppendTool);
    rowsAppendHandler = wrappedRowsAppendTool.handler;

    // Create shared spreadsheet
    const title = `ci-columns-get-tests-${Date.now()}`;
    const accessToken = await authProvider.getAccessToken(accountId);
    if (!accessToken) throw new Error('Failed to get access token for initial setup');
    sharedSpreadsheetId = await createTestSpreadsheet(accessToken, { title });
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
      gid: '0',
      rows,
      headers,
    },
    createExtra()
  );

  const appendStructured = appendResp.structuredContent?.result as RowsAppendOutput | undefined;
  assert.equal(appendStructured?.type, 'success', 'rows-append expected success type');

  // Get columns
  const resp = await handler(
    {
      id: sharedSpreadsheetId,
      gid: '0',
    },
    createExtra()
  );

  const structured = resp.structuredContent?.result as ColumnsGetOutput | undefined;
  assert.equal(structured?.type, 'success', 'expected success type');
  if (structured?.type === 'success') {
    assert.deepEqual(structured.columns, ['email', 'name', 'phone'], 'columns mismatch');
    assert.equal(structured.isEmpty, false, 'expected isEmpty to be false');
  }
});

it('sheets-columns-get returns empty for empty sheet', async () => {
  // Create new empty sheet
  const tempAuth = new OAuth2Client();
  const accessToken = await authProvider.getAccessToken(accountId);
  tempAuth.setCredentials({ access_token: accessToken });
  const sheets = google.sheets({ version: 'v4', auth });

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
  assert.ok(newSheetId !== undefined, 'failed to create empty sheet');

  // Get columns from empty sheet
  const resp = await handler(
    {
      id: sharedSpreadsheetId,
      gid: String(newSheetId),
    },
    createExtra()
  );

  const structured = resp.structuredContent?.result as ColumnsGetOutput | undefined;
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

  const structured = resp.structuredContent?.result as ColumnsGetOutput | undefined;
  assert.equal(structured?.type, 'success', 'expected success type');
  // Sheet1 should have columns from first test
  if (structured?.type === 'success') {
    assert.ok(Array.isArray(structured.columns), 'expected columns array');
  }
});
