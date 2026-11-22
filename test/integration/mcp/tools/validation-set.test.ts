import '../../../lib/env-loader.ts';
import { sheets as sheetsApi } from '@googleapis/sheets';
import { mcp } from '@mcp-z/mcp-sheets';
import type { Logger } from '@mcp-z/oauth-google';
import { ProtocolError, ProtocolErrorCode } from '@mcp-z/server';
import assert from 'assert';
import crypto from 'crypto';
import fs from 'fs/promises';
import * as path from 'path';
import type { Input, Output } from '../../../../src/mcp/tools/validation-set.ts';
import { createExtra, type TypedHandler } from '../../../lib/create-extra.ts';
import createMiddlewareContext from '../../../lib/create-middleware-context.ts';
import { addTestTabs, deleteTestTabs, sharedSpreadsheet, type TestTab } from '../../../lib/spreadsheet-helpers.ts';

/**
 * RANGE ALLOCATION MAP - validation-set.test.ts
 *
 * All tests use this file's own tab with non-overlapping column ranges.
 * When adding new tests, choose the next available column from this list.
 *
 * OPTIMIZATION: Consolidated 11 separate tests into 3 tests to save 8 API writes
 *
 * Allocated Ranges:
 * - A1:A10 thru Q1:Q10  = Test 1 (consolidated): All validation types batch (ONE_OF_LIST,
 *                         ONE_OF_RANGE, NUMBER, TEXT, DATE, CUSTOM_FORMULA, input messages,
 *                         non-strict) - 15 validation rules in ONE API write (saves 8 writes)
 * - O1:O10, P1:P10 = Test 2: Partial failures test (+ INVALID_RANGE)
 * - R1:R3   = Test 3: A1 range validation test (verifies range isolation)
 *
 * Next available: S1:S10
 */

describe('validation-set tool (service-backed tests)', () => {
  // Shared test resources
  let sharedSpreadsheetId: string;
  let tab: TestTab;
  let auth: Awaited<ReturnType<typeof createMiddlewareContext>>['auth'];
  let logger: Logger;
  let handler: TypedHandler<Input>;
  let tmpDir: string;

  before(async () => {
    try {
      // Create temporary directory
      tmpDir = path.join('.tmp', `validation-set-tests-${crypto.randomUUID()}`);
      await fs.mkdir(tmpDir, { recursive: true });

      // Get middleware for tool creation
      const middlewareContext = await createMiddlewareContext();
      logger = middlewareContext.logger;
      auth = middlewareContext.auth;
      const middleware = middlewareContext.middleware;
      const tool = mcp.toolFactories.validationSet();
      const wrappedTool = middleware.withToolAuth(tool);
      handler = wrappedTool.handler;

      // Create shared spreadsheet for all tests (tests use the file's own tab with non-overlapping ranges)
      sharedSpreadsheetId = await sharedSpreadsheet();
      [tab] = (await addTestTabs(sharedSpreadsheetId, ['validation-set'])) as [TestTab];
    } catch (error) {
      logger.error('Failed to initialize test resources:', { error });
      throw error;
    }
  });

  after(async () => {
    // Cleanup resources - fail fast on errors
    if (tab) await deleteTestTabs(sharedSpreadsheetId, [tab.gid]);
    if (tmpDir) await fs.rm(tmpDir, { recursive: true, force: true });
  });

  it('validation_set applies all validation types (consolidated batch)', async () => {
    const testSheetId = tab.sheetId;

    // OPTIMIZATION: Consolidate tests 1-8 and 10 into single batch request (saves 8 API writes)
    // Tests ONE_OF_LIST, ONE_OF_RANGE, NUMBER, TEXT, DATE, CUSTOM_FORMULA validations,
    // input messages, and non-strict validations
    const requests = [
      // Test 1: ONE_OF_LIST dropdown
      {
        range: 'A1:A10',
        rule: {
          conditionType: 'ONE_OF_LIST' as const,
          values: ['Red', 'Green', 'Blue'],
          showDropdown: true,
          strict: true,
        },
      },
      // Test 2: ONE_OF_RANGE dropdown
      {
        range: 'B1:B10',
        rule: {
          conditionType: 'ONE_OF_RANGE' as const,
          sourceRange: 'D1:D5',
          showDropdown: true,
          strict: true,
        },
      },
      // Test 3: NUMBER validations
      {
        range: 'C1:C10',
        rule: {
          conditionType: 'NUMBER_GREATER' as const,
          values: [0],
          strict: true,
        },
      },
      {
        range: 'D1:D10',
        rule: {
          conditionType: 'NUMBER_BETWEEN' as const,
          values: [1, 100],
          strict: true,
        },
      },
      // Test 4: TEXT validations
      {
        range: 'E1:E10',
        rule: {
          conditionType: 'TEXT_CONTAINS' as const,
          values: ['@'],
          strict: true,
        },
      },
      {
        range: 'F1:F10',
        rule: {
          conditionType: 'TEXT_IS_EMAIL' as const,
          strict: true,
        },
      },
      {
        range: 'G1:G10',
        rule: {
          conditionType: 'TEXT_IS_URL' as const,
          strict: true,
        },
      },
      // Test 5: DATE validations
      {
        range: 'H1:H10',
        rule: {
          conditionType: 'DATE_AFTER' as const,
          values: ['2024-01-01'],
          strict: true,
        },
      },
      {
        range: 'I1:I10',
        rule: {
          conditionType: 'DATE_BETWEEN' as const,
          values: ['2024-01-01', '2024-12-31'],
          strict: true,
        },
      },
      // Test 6: CUSTOM_FORMULA validation
      {
        range: 'J1:J10',
        rule: {
          conditionType: 'CUSTOM_FORMULA' as const,
          formula: '=A1>0',
          strict: true,
        },
      },
      // Test 7: Multiple validation types in batch
      {
        range: 'K1:K10',
        rule: {
          conditionType: 'ONE_OF_LIST' as const,
          values: ['Option1', 'Option2', 'Option3'],
          showDropdown: true,
          strict: true,
        },
      },
      {
        range: 'L1:L10',
        rule: {
          conditionType: 'NUMBER_GREATER' as const,
          values: [0],
          strict: true,
        },
      },
      {
        range: 'M1:M10',
        rule: {
          conditionType: 'TEXT_IS_EMAIL' as const,
          strict: true,
        },
      },
      // Test 8: Input message validation
      {
        range: 'N1:N10',
        rule: {
          conditionType: 'ONE_OF_LIST' as const,
          values: ['Yes', 'No'],
          showDropdown: true,
          strict: true,
        },
        inputMessage: 'Please select Yes or No',
      },
      // Test 10: Non-strict validation (warning mode)
      {
        range: 'Q1:Q10',
        rule: {
          conditionType: 'ONE_OF_LIST' as const,
          values: ['Valid', 'Options'],
          showDropdown: true,
          strict: false, // Warning mode - allows invalid entries
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

    const structured = (response.structuredContent as { result?: unknown } | undefined)?.result as Output | undefined;
    assert.ok(structured, 'Response missing structuredContent.result');
    if (structured?.type !== 'success') {
      assert.fail('Validation set consolidated batch operation failed');
    }
    assert.strictEqual(structured.successCount, 15, 'Should set 15 validation rules successfully (tests 1-8,10 consolidated)');
    assert.strictEqual(structured.failedRanges, undefined, 'Should have no failed ranges');
  });

  it('[O1:O10, P1:P10] validation_set handles partial failures gracefully', async () => {
    const testSheetId = tab.sheetId;

    const requests = [
      {
        range: 'O1:O10',
        rule: {
          conditionType: 'ONE_OF_LIST' as const,
          values: ['Valid1', 'Valid2'],
          showDropdown: true,
          strict: true,
        },
      },
      {
        range: 'INVALID_RANGE',
        rule: {
          conditionType: 'NUMBER_GREATER' as const,
          values: [0],
          strict: true,
        },
      },
      {
        range: 'P1:P10',
        rule: {
          conditionType: 'TEXT_IS_EMAIL' as const,
          strict: true,
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

    const structured = (response.structuredContent as { result?: unknown } | undefined)?.result as Output | undefined;
    assert.ok(structured, 'Response missing structuredContent.result');
    assert.strictEqual(structured.type, 'success', 'Expected success result even with partial failure');

    if (structured.type === 'success') {
      assert.strictEqual(structured.successCount, 2, 'Should set 2 valid validation rules');
      assert.ok(Array.isArray(structured.failedRanges), 'Should have failedRanges array');
      assert.strictEqual(structured.failedRanges.length, 1, 'Should have 1 failed range');
      const firstFailed = structured.failedRanges[0];
      if (firstFailed) {
        assert.strictEqual(firstFailed.range, 'INVALID_RANGE', 'Failed range should be INVALID_RANGE');
      }
    }
  });

  it('[R1:R3] validation_set respects A1 range and does not apply to entire sheet', async () => {
    const testSheetId = tab.sheetId;

    // Set validation ONLY on R1:R3
    const requests = [
      {
        range: 'R1:R3',
        rule: {
          conditionType: 'ONE_OF_LIST' as const,
          values: ['Red', 'Green', 'Blue'],
          showDropdown: true,
          strict: true,
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

    const structured = (response.structuredContent as { result?: unknown } | undefined)?.result as Output | undefined;
    assert.ok(structured, 'Response missing structuredContent.result');
    assert.strictEqual(structured.type, 'success', 'Expected success result');

    if (structured.type === 'success') {
      assert.strictEqual(structured.successCount, 1, 'Should set 1 validation rule');

      // Verify the validation was applied correctly by reading the sheet
      const sheets = sheetsApi({ version: 'v4', auth: auth });
      const sheetData = await sheets.spreadsheets.get({
        spreadsheetId: sharedSpreadsheetId,
        includeGridData: true,
        ranges: [`'${structured.sheetTitle}'!R1:S4`],
      });

      const sheet = sheetData.data.sheets?.[0];
      assert.ok(sheet, 'Sheet should exist');

      const gridData = sheet.data?.[0];
      assert.ok(gridData, 'Grid data should exist');

      const rowData = gridData.rowData || [];

      // R1:R3 should have validation
      for (let row = 0; row < 3; row++) {
        const cellData = rowData[row]?.values?.[0];
        assert.ok(cellData?.dataValidation, `Cell R${row + 1} should have validation`);
      }

      // R4 should NOT have validation (outside the range)
      const r4Cell = rowData[3]?.values?.[0];
      assert.strictEqual(r4Cell?.dataValidation, undefined, 'Cell R4 should NOT have validation');

      // S1 should NOT have validation (different column)
      const s1Cell = rowData[0]?.values?.[1];
      assert.strictEqual(s1Cell?.dataValidation, undefined, 'Cell S1 should NOT have validation');
    }
  });

  it('validation_set rejects when every range fails, naming each failure', async () => {
    const rule = { conditionType: 'NUMBER_GREATER' as const, values: [0], strict: true };
    await assert.rejects(
      handler(
        {
          id: sharedSpreadsheetId,
          gid: tab.gid,
          requests: [
            { range: 'INVALID_RANGE', rule },
            { range: 'NOT_A_RANGE', rule },
          ],
        },
        createExtra()
      ),
      (error: Error) => {
        assert.ok(error.message.includes('every range failed'), error.message);
        assert.ok(error.message.includes('INVALID_RANGE'), error.message);
        assert.ok(error.message.includes('NOT_A_RANGE'), error.message);
        return true;
      }
    );
  });

  // Input errors reach the client as InvalidParams, not InternalError.
  for (const input of ['Sheet1!A1:B2', 'a1:b2', '$A$1', 'A2:B', 'B', '5']) {
    it(`validation-set rejects "${input}" as InvalidParams when every range fails`, async () => {
      await assert.rejects(
        handler({ id: sharedSpreadsheetId, gid: '0', requests: [{ range: input, rule: { conditionType: 'NUMBER_GREATER' as const, values: [0], strict: true } }] }, createExtra()),
        (error: unknown) => error instanceof ProtocolError && error.code === ProtocolErrorCode.InvalidParams && error.message.startsWith('every ')
      );
    });
  }

  it('validation-set rejects a missing sheet as InvalidParams', async () => {
    await assert.rejects(
      handler({ id: sharedSpreadsheetId, gid: '987654321', requests: [{ range: 'A1', rule: { conditionType: 'NUMBER_GREATER' as const, values: [0], strict: true } }] }, createExtra()),
      (error: unknown) => error instanceof ProtocolError && error.code === ProtocolErrorCode.InvalidParams && error.message.includes('Sheet not found')
    );
  });
});
