import type { Logger, LoopbackOAuthProvider } from '@mcp-z/oauth-google';
import assert from 'assert';
import crypto from 'crypto';
import fs from 'fs/promises';
import { google } from 'googleapis';
import * as path from 'path';
import createValuesCsvUpdateTool, { type Input, type Output } from '../../../../src/mcp/tools/values-csv-update.js';
import { createExtra, type TypedHandler } from '../../../lib/create-extra.js';
import createMiddlewareContext from '../../../lib/create-middleware-context.js';
import { createTestSpreadsheet, deleteTestSpreadsheet } from '../../../lib/spreadsheet-helpers.js';

let handler: TypedHandler<Input>;
let sharedSpreadsheetId: string;
let authProvider: LoopbackOAuthProvider;
let accountId: string;
let logger: Logger;
let tmpDir: string;

before(async () => {
  try {
    // Create temporary directory
    tmpDir = path.join('.tmp', `values-csv-update-tests-${crypto.randomUUID()}`);
    await fs.mkdir(tmpDir, { recursive: true });

    // Create middleware and tools
    const middlewareContext = await createMiddlewareContext();
    authProvider = middlewareContext.authProvider;
    logger = middlewareContext.logger;
    const middleware = middlewareContext.middleware;
    accountId = middlewareContext.accountId;

    const tool = createValuesCsvUpdateTool();
    const wrappedTool = middleware.withToolAuth(tool);
    handler = wrappedTool.handler;

    // Create shared spreadsheet
    const title = `ci-values-csv-update-tests-${Date.now()}`;
    sharedSpreadsheetId = await createTestSpreadsheet(await authProvider.getAccessToken(accountId), { title });
  } catch (error) {
    console.error('Failed to initialize test resources:', error);
    throw error;
  }
});

after(async () => {
  // Cleanup
  const accessToken = await authProvider.getAccessToken(accountId);
  await deleteTestSpreadsheet(accessToken, sharedSpreadsheetId, logger);
  await fs.rm(tmpDir, { recursive: true, force: true });
});

it('sheets-values-csv-update excludes source headers from data range', async () => {
  // Create test CSV with headers
  const csvPath = path.join(tmpDir, 'data-with-headers.csv');
  const csvContent = 'Header1,Header2,Header3\nvalue1a,value1b,value1c\nvalue2a,value2b,value2c';
  await fs.writeFile(csvPath, csvContent);

  const sourceUri = `file://${csvPath}`;
  const resp = await handler(
    {
      id: sharedSpreadsheetId,
      gid: '0',
      sourceUri,
      startRange: 'A1',
      valueInputOption: 'USER_ENTERED',
      sourceHasHeaders: true, // Exclude source headers from data range
    },
    createExtra()
  );

  const structured = resp.structuredContent?.result as Output | undefined;
  assert.equal(structured?.type, 'success', 'expected success type');
  if (structured?.type === 'success') {
    assert.equal(structured.updatedRows, 2, 'expected 2 data rows (source headers excluded)');
    assert.equal(structured.updatedColumns, 3, 'expected 3 columns');
  }
});

it('sheets-values-csv-update includes source first row as data', async () => {
  // Create data-only CSV (no headers)
  const csvPath = path.join(tmpDir, 'data-only.csv');
  const csvContent = 'data1a,data1b,data1c\ndata2a,data2b,data2c\ndata3a,data3b,data3c';
  await fs.writeFile(csvPath, csvContent);

  const sourceUri = `file://${csvPath}`;
  const resp = await handler(
    {
      id: sharedSpreadsheetId,
      gid: '0',
      sourceUri,
      startRange: 'B2',
      valueInputOption: 'USER_ENTERED',
      sourceHasHeaders: false, // Include all rows (first row is data)
    },
    createExtra()
  );

  const structured = resp.structuredContent?.result as Output | undefined;
  assert.equal(structured?.type, 'success', 'expected success type');
  if (structured?.type === 'success') {
    assert.equal(structured.updatedRows, 3, 'expected 3 rows (all rows included)');
    assert.equal(structured.updatedColumns, 3, 'expected 3 columns');
  }
});

it('sheets-values-csv-update writes to custom range', async () => {
  // Create small CSV
  const csvPath = path.join(tmpDir, 'small.csv');
  const csvContent = 'A,B\nC,D';
  await fs.writeFile(csvPath, csvContent);

  const sourceUri = `file://${csvPath}`;
  const resp = await handler(
    {
      id: sharedSpreadsheetId,
      gid: '0',
      sourceUri,
      startRange: 'D5', // Start at D5
      valueInputOption: 'USER_ENTERED',
      sourceHasHeaders: true, // Exclude source headers
    },
    createExtra()
  );

  const structured = resp.structuredContent?.result as Output | undefined;
  assert.equal(structured?.type, 'success', 'expected success type');
  if (structured?.type === 'success') {
    assert.ok(structured.updatedRange?.includes('D5'), 'expected range to start at D5');
  }
});

it('sheets-values-csv-update handles empty CSV', async () => {
  // Create empty CSV
  const csvPath = path.join(tmpDir, 'empty.csv');
  await fs.writeFile(csvPath, '');

  const sourceUri = `file://${csvPath}`;

  // Should throw error for empty CSV
  await assert.rejects(
    async () => {
      await handler(
        {
          id: sharedSpreadsheetId,
          gid: '0',
          sourceUri,
          startRange: 'A1',
          valueInputOption: 'USER_ENTERED',
          sourceHasHeaders: true,
        },
        createExtra()
      );
    },
    (error: unknown) => {
      assert.ok(error instanceof Error && error.message.includes('empty'), 'expected error about empty CSV');
      return true;
    }
  );
});

it('sheets-values-csv-update uses RAW value input option', async () => {
  // Create CSV with formula-like text
  const csvPath = path.join(tmpDir, 'with-formulas.csv');
  const csvContent = 'Formula\n=SUM(A1:A10)\n=TODAY()';
  await fs.writeFile(csvPath, csvContent);

  const sourceUri = `file://${csvPath}`;
  const resp = await handler(
    {
      id: sharedSpreadsheetId,
      gid: '0',
      sourceUri,
      startRange: 'A1',
      valueInputOption: 'RAW', // Don't parse as formulas
      sourceHasHeaders: true,
    },
    createExtra()
  );

  const structured = resp.structuredContent?.result as Output | undefined;
  assert.equal(structured?.type, 'success', 'expected success type');
  // Values should be written as text, not executed as formulas
});

it('sheets-values-csv-update initializes sheet with headers from CSV', async () => {
  // Create CSV with headers that we want in the sheet
  const csvPath = path.join(tmpDir, 'with-headers-to-keep.csv');
  const csvContent = 'id,name,email\n1,Alice,alice@example.com\n2,Bob,bob@example.com';
  await fs.writeFile(csvPath, csvContent);

  const sourceUri = `file://${csvPath}`;
  const resp = await handler(
    {
      id: sharedSpreadsheetId,
      gid: '0',
      sourceUri,
      startRange: 'A10', // Use a different range to avoid conflicts
      valueInputOption: 'USER_ENTERED',
      sourceHasHeaders: false, // Treat headers as data to write them
    },
    createExtra()
  );

  const structured = resp.structuredContent?.result as Output | undefined;
  assert.equal(structured?.type, 'success', 'expected success type');
  if (structured?.type === 'success') {
    assert.equal(structured.updatedRows, 3, 'expected 3 rows (headers + 2 data rows)');
    assert.equal(structured.updatedColumns, 3, 'expected 3 columns');
  }

  // Verify actual cell content using direct API
  const auth = new google.auth.OAuth2();
  auth.setCredentials({ access_token: await authProvider.getAccessToken(accountId) });
  const sheets = google.sheets({ version: 'v4', auth });
  const valuesResponse = await sheets.spreadsheets.values.get({
    spreadsheetId: sharedSpreadsheetId,
    range: 'Sheet1!A10:C12',
  });

  const values = valuesResponse.data.values || [];
  assert.equal(values.length, 3, 'expected 3 rows in actual data');
  assert.deepStrictEqual(values[0], ['id', 'name', 'email'], 'first row should be headers');
  assert.deepStrictEqual(values[1], ['1', 'Alice', 'alice@example.com'], 'second row should be first data row');
  assert.deepStrictEqual(values[2], ['2', 'Bob', 'bob@example.com'], 'third row should be second data row');
});
