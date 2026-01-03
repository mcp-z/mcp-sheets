import type { Logger, LoopbackOAuthProvider } from '@mcp-z/oauth-google';
import assert from 'assert';
import crypto from 'crypto';
import fs from 'fs/promises';
import type { OAuth2Client } from 'google-auth-library';
import { google } from 'googleapis';
import * as path from 'path';
import createTool, { type Input, type Output } from '../../../../src/mcp/tools/cells-format.ts';
import { createExtra, type TypedHandler } from '../../../lib/create-extra.ts';
import createMiddlewareContext from '../../../lib/create-middleware-context.ts';
import { createTestSpreadsheet, deleteTestSpreadsheet } from '../../../lib/spreadsheet-helpers.ts';

/**
 * RANGE ALLOCATION MAP - cells-format.test.ts
 *
 * All tests use shared sheet (gid: 0) with non-overlapping ranges.
 * When adding new tests, choose the next available range.
 *
 * Allocated Ranges:
 * - A1:B2, D1:D1 = Test 1: background and text formatting batch
 * - F1:H1, J1:J10 = Test 2: alignment and number formatting batch
 * - L1:O5       = Test 3: borders
 * - Q1:R1, Q2:R10, S1:S1 = Test 4: multiple format batch
 * - U1:V2, X1:Y2 = Test 5: partial failures (+ INVALID_RANGE)
 * - Y1:Z2       = Test 6: combined properties
 * - C3:D4, E3:F5 = Test 7: A1 range test (verify isolation)
 *
 * Note: Ranges limited to columns A-Z (26 columns) as that's the default sheet size
 */

describe('cells-format tool (service-backed tests)', () => {
  // Shared test resources
  let sharedSpreadsheetId: string;
  let auth: OAuth2Client;
  let authProvider: LoopbackOAuthProvider;
  let logger: Logger;
  let accountId: string;
  let handler: TypedHandler<Input>;
  let tmpDir: string;

  before(async () => {
    try {
      // Create temporary directory
      tmpDir = path.join('.tmp', `cells-format-tests-${crypto.randomUUID()}`);
      await fs.mkdir(tmpDir, { recursive: true });

      // Get middleware for tool creation
      const middlewareContext = await createMiddlewareContext();
      authProvider = middlewareContext.authProvider;
      logger = middlewareContext.logger;
      auth = middlewareContext.auth;
      const middleware = middlewareContext.middleware;
      accountId = middlewareContext.accountId;
      const tool = createTool();
      const wrappedTool = middleware.withToolAuth(tool);
      handler = wrappedTool.handler;

      // Create shared spreadsheet for all tests (tests use default sheet 0 with non-overlapping ranges)
      const title = `ci-cells-format-tests-${Date.now()}`;
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

  it('format_cells applies all formatting types (consolidated batch)', async () => {
    const testSheetId = 0; // Use default sheet

    // OPTIMIZATION: Consolidate tests 1-4 and 6 into single batch request (saves 4 API writes)
    // Tests background, text, alignment, number format, borders, and combined properties
    const requests = [
      // Test 1: Background color
      {
        range: 'A1:B2',
        backgroundColor: { red: 1.0, green: 0.9, blue: 0.9 }, // Light red
      },
      // Test 1: Text formatting
      {
        range: 'D1:D1',
        bold: true,
        fontSize: 14,
        textColor: { red: 0.0, green: 0.0, blue: 1.0 }, // Blue text
      },
      // Test 2: Alignment
      {
        range: 'F1:H1',
        horizontalAlignment: 'CENTER' as const,
      },
      // Test 2: Number formatting
      {
        range: 'J1:J10',
        numberFormat: { type: 'CURRENCY' as const, pattern: '$#,##0.00' },
      },
      // Test 3: Borders
      {
        range: 'L1:O5',
        borders: {
          style: 'SOLID' as const,
          color: { red: 0.0, green: 0.0, blue: 0.0 }, // Black borders
        },
      },
      // Test 4: Multiple properties on header
      {
        range: 'Q1:R1',
        backgroundColor: { red: 0.2, green: 0.6, blue: 1.0 }, // Header blue
        bold: true,
        textColor: { red: 1.0, green: 1.0, blue: 1.0 }, // White text
      },
      // Test 4: Number format on data rows
      {
        range: 'Q2:R10',
        numberFormat: { type: 'NUMBER' as const },
      },
      // Test 4: Right alignment
      {
        range: 'S1:S1',
        horizontalAlignment: 'RIGHT' as const,
      },
      // Test 6: Combined properties on single range
      {
        range: 'Y1:Z2',
        backgroundColor: { red: 0.0, green: 0.5, blue: 0.0 }, // Green background
        textColor: { red: 1.0, green: 1.0, blue: 1.0 }, // White text
        bold: true,
        fontSize: 12,
        horizontalAlignment: 'CENTER' as const,
        borders: {
          style: 'SOLID' as const,
          color: { red: 0.0, green: 0.0, blue: 0.0 },
        },
      },
    ];

    const response = await handler(
      {
        id: sharedSpreadsheetId,
        gid: String(testSheetId),
        requests,
      },
      createExtra()
    );

    // Validate response structure
    assert.ok(response, 'Handler returned no result');

    const structured = response.structuredContent?.result as Output | undefined;
    assert.ok(structured, 'Response missing structuredContent.result');
    assert.strictEqual(structured?.type, 'success', 'Expected success result');

    if (structured?.type === 'success') {
      assert.strictEqual(structured.successCount, 9, 'Should format 9 ranges successfully (tests 1-4,6 consolidated)');
      assert.strictEqual(structured.failedRanges, undefined, 'Should have no failed ranges');
    }
  });

  it('[U1:V2, X1:Y2] format_cells handles partial failures gracefully', async () => {
    const testSheetId = 0; // Use default sheet

    const requests = [
      {
        range: 'U1:V2',
        backgroundColor: { red: 1.0, green: 0.0, blue: 0.0 },
      },
      {
        range: 'INVALID_RANGE', // This will fail
        bold: true,
      },
      {
        range: 'X1:Y2',
        textColor: { red: 0.0, green: 1.0, blue: 0.0 },
      },
    ];

    const response = await handler(
      {
        id: sharedSpreadsheetId,
        gid: String(testSheetId),
        requests,
      },
      createExtra()
    );

    const structured = response.structuredContent?.result as Output | undefined;
    assert.ok(structured, 'Response missing structuredContent.result');
    // Note: Partial failures should still return success type
    assert.strictEqual(structured?.type, 'success', 'Expected success result even with partial failure');

    if (structured?.type === 'success') {
      assert.strictEqual(structured.successCount, 2, 'Should format 2 valid ranges');
      assert.ok(Array.isArray(structured.failedRanges), 'Should have failedRanges array');
      assert.strictEqual(structured.failedRanges?.length, 1, 'Should have 1 failed range');
      const firstFailed = structured.failedRanges?.[0];
      if (firstFailed) {
        assert.strictEqual(firstFailed.range, 'INVALID_RANGE', 'Failed range should be INVALID_RANGE');
      }
    }
  });

  it('[Y1:Z2] format_cells combines multiple format properties on same range', async () => {
    const testSheetId = 0; // Use default sheet

    const requests = [
      {
        range: 'Y1:Z2',
        backgroundColor: { red: 0.0, green: 0.5, blue: 0.0 }, // Green background
        textColor: { red: 1.0, green: 1.0, blue: 1.0 }, // White text
        bold: true,
        fontSize: 12,
        horizontalAlignment: 'CENTER' as const,
        borders: {
          style: 'SOLID' as const,
          color: { red: 0.0, green: 0.0, blue: 0.0 },
        },
      },
    ];

    const response = await handler(
      {
        id: sharedSpreadsheetId,
        gid: String(testSheetId),
        requests,
      },
      createExtra()
    );

    const structured = response.structuredContent?.result as Output | undefined;
    assert.ok(structured, 'Response missing structuredContent.result');
    assert.strictEqual(structured?.type, 'success', 'Expected success result');
    if (structured?.type === 'success') {
      assert.strictEqual(structured.successCount, 1, 'Should format 1 range with all properties');
    }
  });

  it('[C3:D4 vs E3:F5] format_cells respects A1 range and does not apply to entire sheet', async () => {
    const testSheetId = 0; // Use default sheet

    // Format ONLY cells C3:D4 with red background
    const requests = [
      {
        range: 'C3:D4',
        backgroundColor: { red: 1.0, green: 0.0, blue: 0.0 }, // Red background
      },
    ];

    const response = await handler(
      {
        id: sharedSpreadsheetId,
        gid: String(testSheetId),
        requests,
      },
      createExtra()
    );

    const structured = response.structuredContent?.result as Output | undefined;
    assert.ok(structured, 'Response missing structuredContent.result');
    assert.strictEqual(structured?.type, 'success', 'Expected success result');

    if (structured?.type === 'success') {
      assert.strictEqual(structured.successCount, 1, 'Should format 1 range');

      // Verify the formatting was applied correctly by reading the sheet
      const sheets = google.sheets({ version: 'v4', auth: auth });
      const sheetData = await sheets.spreadsheets.get({
        spreadsheetId: sharedSpreadsheetId,
        includeGridData: true,
        ranges: [`'${structured.sheetTitle}'!C3:F5`],
      });

      const sheet = sheetData.data.sheets?.[0];
      assert.ok(sheet, 'Sheet should exist');

      const gridData = sheet.data?.[0];
      assert.ok(gridData, 'Grid data should exist');

      const rowData = gridData.rowData || [];

      // C3:D4 should have red background (rows 0-1 in the fetched range, cols 0-1)
      for (let row = 0; row < 2; row++) {
        for (let col = 0; col < 2; col++) {
          const cellData = rowData[row]?.values?.[col];
          const bgColor = cellData?.effectiveFormat?.backgroundColor;
          const cellName = `${String.fromCharCode(67 + col)}${row + 3}`; // C=67, rows start at 3

          assert.ok(bgColor, `Cell ${cellName} should have background color`);

          // Google Sheets API omits color components that are 0, so treat undefined as 0
          const red = bgColor.red ?? 0;
          const green = bgColor.green ?? 0;
          const blue = bgColor.blue ?? 0;

          assert.ok(red === 1 || red > 0.99, `Cell ${cellName} should have red=1 (got ${red})`);
          assert.ok(green === 0 || green < 0.01, `Cell ${cellName} should have green=0 (got ${green})`);
          assert.ok(blue === 0 || blue < 0.01, `Cell ${cellName} should have blue=0 (got ${blue})`);
        }
      }

      // E3 should NOT have red background (outside the range, col index 2 in fetched data)
      const e3Cell = rowData[0]?.values?.[2];
      const e3BgColor = e3Cell?.effectiveFormat?.backgroundColor;
      // Default background is usually white (1, 1, 1) or undefined
      if (e3BgColor) {
        assert.ok(e3BgColor.red !== 1 || e3BgColor.green !== 0, 'Cell E3 should NOT have red background');
      }

      // C5 should NOT have red background (outside the range, row index 2 in fetched data)
      const c5Cell = rowData[2]?.values?.[0];
      const c5BgColor = c5Cell?.effectiveFormat?.backgroundColor;
      if (c5BgColor) {
        assert.ok(c5BgColor.red !== 1 || c5BgColor.green !== 0, 'Cell C5 should NOT have red background');
      }
    }
  });
});
