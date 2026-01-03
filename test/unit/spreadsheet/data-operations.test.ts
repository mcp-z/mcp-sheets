import type { Logger } from '@mcp-z/oauth-google';
import assert from 'assert';
import type { OAuth2Client } from 'google-auth-library';
import type { drive_v3, sheets_v4 } from 'googleapis';
import { google } from 'googleapis';
import { appendRows } from '../../../src/spreadsheet/data-operations.ts';
import { ensureTabAndHeaders } from '../../../src/spreadsheet/sheet-operations.ts';
import createMiddlewareContext from '../../lib/create-middleware-context.ts';

let auth: OAuth2Client;
let logger: Logger;
let sheets: sheets_v4.Sheets;
let drive: drive_v3.Drive;

before(async () => {
  const middlewareContext = await createMiddlewareContext();
  auth = middlewareContext.auth;
  logger = middlewareContext.logger;
  sheets = google.sheets({ version: 'v4', auth: auth });
  drive = google.drive({ version: 'v3', auth: auth });
});

it('integration: ensureTabAndHeaders + appendRows', async () => {
  // integration test that creates a temporary spreadsheet, exercises the helpers, then deletes the spreadsheet.

  // Create a temporary spreadsheet for this test
  const title = `test-sheets-writer-${Date.now()}`;
  // SheetsHttpClient exposes a createSpreadsheet helper (not a nested `spreadsheets` namespace)
  const created = await sheets.spreadsheets.create({ requestBody: { properties: { title } } });
  const spreadsheetId = created.data.spreadsheetId as string;
  assert.ok(spreadsheetId, 'failed to create test spreadsheet');

  try {
    const sheetTitle = 'Sheet1';

    // 1) ensure headers (no requiredHeader provided -> default header will be used)
    const { header: header1, keySet: ks1 } = await ensureTabAndHeaders(sheets, { spreadsheetId, sheetTitle, logger: logger });
    assert.ok(Array.isArray(header1));
    assert.ok(ks1 instanceof Set);
    // initially there should be no keys
    assert.strictEqual(ks1.size, 0);

    // 2) append two rows
    const rows = [
      ['mid1', 'gmail', 't1'],
      ['mid2', 'gmail', 't2'],
    ];
    const appendRes = await appendRows(sheets, { spreadsheetId, sheetTitle, rows, keyColumns: ['id', 'provider'], logger: logger });
    // expect 2 rows inserted
    assert.strictEqual(appendRes.updatedRows, 2);
    assert.deepStrictEqual(appendRes.inserted.length, 2);

    // 3) re-run ensureTabAndHeaders to build keySet from existing rows
    const { header: header2, keySet: ks2 } = await ensureTabAndHeaders(sheets, { spreadsheetId, sheetTitle, keyColumns: ['id', 'provider'], logger: logger });
    assert.ok(Array.isArray(header2));
    // keySet should now contain two keys
    assert.strictEqual(ks2.size >= 2, true);
    assert.ok(ks2.has('gmail\\mid1'), 'expected key gmail\\mid1');
    assert.ok(ks2.has('gmail\\mid2'), 'expected key gmail\\mid2');

    // 4) appendRows with keySet should skip duplicates
    const moreRows = [
      ['mid1', 'gmail', 't1-dup'],
      ['mid3', 'gmail', 't3'],
    ];
    const appendRes2 = await appendRows(sheets, { spreadsheetId, sheetTitle, rows: moreRows, keySet: ks2, keyColumns: ['id', 'provider'], logger: logger });
    // Only mid3 should be inserted
    assert.strictEqual(appendRes2.updatedRows, 1);
    assert.deepStrictEqual(appendRes2.inserted, ['gmail\\mid3']);
  } finally {
    // Clean up: delete the temporary spreadsheet and fail loudly if close cannot be completed
    try {
      await drive.files.delete({ fileId: spreadsheetId as string });
    } catch (e) {
      assert.fail(`failed to delete test spreadsheet: ${e instanceof Error ? e.message : String(e)}`);
    }
  }
});
