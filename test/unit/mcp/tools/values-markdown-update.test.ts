import { sheets as sheetsApi } from '@googleapis/sheets';
import { mcp } from '@mcp-z/mcp-sheets';
import type { Logger, LoopbackOAuthProvider } from '@mcp-z/oauth-google';
import assert from 'assert';
import crypto from 'crypto';
import fs from 'fs/promises';
import * as path from 'path';
import type { Input, Output } from '../../../../src/mcp/tools/values-markdown-update.ts';
import { createExtra, type TypedHandler } from '../../../lib/create-extra.ts';
import createMiddlewareContext from '../../../lib/create-middleware-context.ts';
import { createTestSpreadsheet, deleteTestSpreadsheet } from '../../../lib/spreadsheet-helpers.ts';

/**
 * RANGE ALLOCATION MAP - values-markdown-update.test.ts
 *
 * All tests use shared sheet (gid: 0) with non-overlapping cells.
 *
 * Allocated Cells:
 * - I2 = Test 1: two links + bold, read back via textFormatRuns
 * - I3, "A1:B2" = Test 2: one valid cell + one multi-cell "cell" (rejected) in the same batch
 */

describe('values-markdown-update tool (service-backed tests)', () => {
  // Shared test resources
  let sharedSpreadsheetId: string;
  let auth: Awaited<ReturnType<typeof createMiddlewareContext>>['auth'];
  let authProvider: LoopbackOAuthProvider;
  let logger: Logger;
  let accountId: string;
  let handler: TypedHandler<Input>;
  let tmpDir: string;

  before(async () => {
    try {
      // Create temporary directory
      tmpDir = path.join('.tmp', `values-markdown-update-tests-${crypto.randomUUID()}`);
      await fs.mkdir(tmpDir, { recursive: true });

      // Get middleware for tool creation
      const middlewareContext = await createMiddlewareContext();
      authProvider = middlewareContext.authProvider;
      logger = middlewareContext.logger;
      auth = middlewareContext.auth;
      const middleware = middlewareContext.middleware;
      accountId = middlewareContext.accountId;
      const tool = mcp.toolFactories.valuesMarkdownUpdate();
      const wrappedTool = middleware.withToolAuth(tool);
      handler = wrappedTool.handler;

      // Create shared spreadsheet for all tests (tests use default sheet 0 with non-overlapping cells)
      const title = `ci-values-markdown-update-tests-${Date.now()}`;
      sharedSpreadsheetId = await createTestSpreadsheet(await authProvider.getAccessToken(accountId), { title });
    } catch (error) {
      logger.error('Failed to initialize test resources:', { error });
      throw error;
    }
  });

  after(async () => {
    // Cleanup resources - fail fast on errors
    const accessToken = await authProvider.getAccessToken(accountId);
    await deleteTestSpreadsheet(accessToken, sharedSpreadsheetId, logger);
    await fs.rm(tmpDir, { recursive: true, force: true });
  });

  it('[I2] writes two clickable links plus bold text and reads back matching textFormatRuns', async () => {
    const testSheetId = 0; // Use default sheet
    const cell = 'I2';
    const linkAUri = 'https://a.example.com/one';
    const linkBUri = 'https://b.example.com/two';

    const response = await handler(
      {
        id: sharedSpreadsheetId,
        gid: String(testSheetId),
        requests: [
          {
            cell,
            text: `See **[Alpha](${linkAUri})** and [Beta](${linkBUri}) for details`,
          },
        ],
      },
      createExtra()
    );

    assert.ok(response, 'Handler returned no result');
    const structured = (response.structuredContent as { result?: unknown } | undefined)?.result as Output | undefined;
    assert.ok(structured, 'Response missing structuredContent.result');
    assert.strictEqual(structured?.type, 'success', 'Expected success result');

    if (structured?.type !== 'success') return;
    assert.strictEqual(structured.successCount, 1, 'Should write 1 cell');
    assert.strictEqual(structured.failedCells, undefined, 'Should have no failed cells');
    assert.strictEqual(structured.cells.length, 1, 'Should report 1 written cell');

    const written = structured.cells[0];
    assert.ok(written, 'Missing written cell entry');
    assert.strictEqual(written.cell, cell);
    assert.strictEqual(written.text, 'See Alpha and Beta for details', 'Plain text should have markdown syntax stripped');
    assert.strictEqual(written.linkCount, 2, 'Should report 2 links');

    // Read back what actually landed via the Sheets API — truth over a hopeful "done".
    const sheets = sheetsApi({ version: 'v4', auth });
    const sheetData = await sheets.spreadsheets.get({
      spreadsheetId: sharedSpreadsheetId,
      ranges: [`'${structured.sheetTitle}'!${cell}`],
      fields: 'sheets.data.rowData.values(formattedValue,textFormatRuns)',
    });

    const rowData = sheetData.data.sheets?.[0]?.data?.[0]?.rowData;
    const cellData = rowData?.[0]?.values?.[0];
    assert.ok(cellData, 'Cell data should exist');
    assert.strictEqual(cellData.formattedValue, 'See Alpha and Beta for details');

    const runs = cellData.textFormatRuns ?? [];
    const linkUris = runs.map((run) => run.format?.link?.uri).filter((uri): uri is string => Boolean(uri));
    assert.strictEqual(linkUris.length, 2, 'Should have 2 runs carrying a link');
    assert.ok(linkUris.includes(linkAUri), `Expected ${linkAUri} among written links`);
    assert.ok(linkUris.includes(linkBUri), `Expected ${linkBUri} among written links`);

    // The "Alpha" link run should also carry bold (it was wrapped in **...**).
    const boldLinkRun = runs.find((run) => run.format?.link?.uri === linkAUri);
    assert.strictEqual(boldLinkRun?.format?.bold, true, 'Alpha link run should also be bold');
  });

  it('[I3, A1:B2] reports a multi-cell "cell" reference as a failed cell without failing the whole batch', async () => {
    const testSheetId = 0; // Use default sheet
    const validCell = 'I3';

    const response = await handler(
      {
        id: sharedSpreadsheetId,
        gid: String(testSheetId),
        requests: [
          { cell: validCell, text: 'plain text, no markdown' },
          { cell: 'A1:B2', text: 'this is a range, not a single cell' }, // rejected per-request
        ],
      },
      createExtra()
    );

    const structured = (response.structuredContent as { result?: unknown } | undefined)?.result as Output | undefined;
    assert.ok(structured, 'Response missing structuredContent.result');
    assert.strictEqual(structured?.type, 'success', 'Expected success result even with a partial failure');

    if (structured?.type !== 'success') return;
    assert.strictEqual(structured.successCount, 1, 'Should write the 1 valid cell');
    assert.strictEqual(structured.cells.length, 1);
    assert.strictEqual(structured.cells[0]?.cell, validCell);

    assert.ok(Array.isArray(structured.failedCells), 'Should have failedCells array');
    assert.strictEqual(structured.failedCells?.length, 1, 'Should have 1 failed cell');
    assert.strictEqual(structured.failedCells?.[0]?.cell, 'A1:B2');
  });
});
