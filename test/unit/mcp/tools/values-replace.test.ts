import type { Logger, LoopbackOAuthProvider } from '@mcp-z/oauth-google';
import assert from 'assert';
import type { OAuth2Client } from 'google-auth-library';
import { google } from 'googleapis';
import createValuesReplaceTool, { type Input, type Output } from '../../../../src/mcp/tools/values-replace.js';
import { createExtra, type TypedHandler } from '../../../lib/create-extra.js';
import createMiddlewareContext from '../../../lib/create-middleware-context.js';
import { createTestSpreadsheet, deleteTestSpreadsheet } from '../../../lib/spreadsheet-helpers.js';

/**
 * OPTIMIZATION STRATEGY - values-replace.test.ts
 *
 * Single Spreadsheet, Single Batch Write, Row-Based Isolation
 * ===========================================================
 *
 * All test data is written in a single batchUpdate call in before().
 * Each test operates on SEPARATE ROWS to prevent interference:
 *
 * Row Layout:
 * - Row 1:  [repl1-*] Basic replacement
 * - Row 2:  [repl2-*] Case sensitivity
 * - Row 3:  [repl3-*] Entire cell match
 * - Row 4:  [repl4-*] Regex patterns
 * - Row 5:  [repl5-*] Formula inclusion
 * - Row 6:  [repl6-*] Cross-sheet scope (also in Sheet2)
 * - Row 7:  [repl7-*] Range scope
 * - Row 8:  [repl8-*] No matches (edge case)
 * - Row 9:  [repl9-*] Empty replacement (deletion)
 *
 * API calls:
 * - Setup: 3 (createSpreadsheet + batchUpdate addSheet + values.batchUpdate)
 * - Tests: 11 (9 replaces + 2 read-back verifications in tests 1 and 5)
 * - Teardown: 1 (deleteSpreadsheet)
 *
 * Total: ~15 API calls
 */

describe('values-replace tool (service-backed tests)', () => {
  let sharedSpreadsheetId: string;
  let sharedGid: string;
  let auth: OAuth2Client;
  let authProvider: LoopbackOAuthProvider;
  let logger: Logger;
  let accountId: string;
  let valuesReplaceHandler: TypedHandler<Input>;

  before(async () => {
    const middlewareContext = await createMiddlewareContext();
    authProvider = middlewareContext.authProvider;
    logger = middlewareContext.logger;
    auth = middlewareContext.auth;
    const middleware = middlewareContext.middleware;
    accountId = middlewareContext.accountId;

    const valuesReplaceTool = createValuesReplaceTool();
    const wrappedTool = middleware.withToolAuth(valuesReplaceTool);
    valuesReplaceHandler = wrappedTool.handler;

    const title = `ci-values-replace-tests-${Date.now()}`;
    const accessToken = await authProvider.getAccessToken(accountId);
    sharedSpreadsheetId = await createTestSpreadsheet(accessToken, { title });
    sharedGid = '0';

    const sheets = google.sheets({ version: 'v4', auth });

    // Create second sheet for allSheets scope testing
    const _addSheetResponse = await sheets.spreadsheets.batchUpdate({
      spreadsheetId: sharedSpreadsheetId,
      requestBody: {
        requests: [{ addSheet: { properties: { title: 'Sheet2' } } }],
      },
    });

    // Write all test data in single batch - each row isolated
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: sharedSpreadsheetId,
      requestBody: {
        valueInputOption: 'USER_ENTERED',
        data: [
          {
            range: 'Sheet1!A1:C9',
            values: [
              // Row 1: Basic replacement
              ['repl1-OldName', 'repl1-OldName-copy', 'other1'],
              // Row 2: Case sensitivity
              ['repl2-UPPER', 'repl2-upper', 'repl2-Upper'],
              // Row 3: Entire cell match
              ['repl3-exact', 'repl3-exact-partial', 'repl3-notexact'],
              // Row 4: Regex patterns
              ['repl4-user_001', 'repl4-user_002', 'repl4-admin_001'],
              // Row 5: Formula inclusion
              ['repl5-value', '=CONCAT("repl5-","formula")', 'repl5-other'],
              // Row 6: Cross-sheet (also appears in Sheet2)
              ['repl6-cross', 'sheet1-only', 'more6'],
              // Row 7: Range scope test
              ['repl7-inA', 'repl7-inB', 'repl7-inC'],
              // Row 8: No matches
              ['repl8-unique', 'no-match-here', 'repl8-another'],
              // Row 9: Empty replacement (deletion)
              ['prefix-repl9-DELETE-suffix', 'repl9-keep', 'other9'],
            ],
          },
          {
            range: 'Sheet2!A1:B1',
            values: [
              // Sheet2 row for allSheets test
              ['repl6-cross', 'sheet2-only'],
            ],
          },
        ],
      },
    });
  });

  after(async () => {
    const accessToken = await authProvider.getAccessToken(accountId);
    await deleteTestSpreadsheet(accessToken, sharedSpreadsheetId, logger);
  });

  // ─────────────────────────────────────────────────────────────────
  // BASIC REPLACEMENT
  // ─────────────────────────────────────────────────────────────────

  it('[repl1] basic find and replace', async () => {
    const result = await valuesReplaceHandler(
      {
        id: sharedSpreadsheetId,
        find: 'repl1-OldName',
        replacement: 'repl1-NewName',
        gid: sharedGid,
      },
      createExtra()
    );

    const branch = result.structuredContent?.result as Output | undefined;
    assert.equal(branch?.type, 'success', 'replace should succeed');

    if (branch?.type === 'success') {
      assert.equal(branch.occurrencesChanged, 2, 'should replace 2 occurrences');
      assert.equal(branch.valuesChanged, 2, 'should change 2 values');
    }

    // READ-BACK VERIFICATION: Don't just trust the count, verify actual replacement
    const sheets = google.sheets({ version: 'v4', auth });
    const readResult = await sheets.spreadsheets.values.get({
      spreadsheetId: sharedSpreadsheetId,
      range: 'Sheet1!A1:C1',
    });
    const row = readResult.data.values?.[0];
    assert.ok(row, 'should have data in row 1');
    assert.equal(row[0], 'repl1-NewName', 'cell A1 should contain replaced text');
    assert.equal(row[1], 'repl1-NewName-copy', 'cell B1 should contain replaced text');
    assert.equal(row[2], 'other1', 'cell C1 should be unchanged (no match)');
  });

  // ─────────────────────────────────────────────────────────────────
  // MATCH OPTIONS
  // ─────────────────────────────────────────────────────────────────

  it('[repl2] matchCase: true only replaces exact case', async () => {
    const result = await valuesReplaceHandler(
      {
        id: sharedSpreadsheetId,
        find: 'repl2-UPPER',
        replacement: 'repl2-REPLACED',
        gid: sharedGid,
        matchCase: true,
      },
      createExtra()
    );

    const branch = result.structuredContent?.result as Output | undefined;
    assert.equal(branch?.type, 'success');

    if (branch?.type === 'success') {
      // Should only match 'repl2-UPPER', not 'repl2-upper' or 'repl2-Upper'
      assert.equal(branch.occurrencesChanged, 1, 'should replace 1 case-matched occurrence');
    }
  });

  it('[repl3] matchEntireCell: true only replaces full cell match', async () => {
    const result = await valuesReplaceHandler(
      {
        id: sharedSpreadsheetId,
        find: 'repl3-exact',
        replacement: 'repl3-matched',
        gid: sharedGid,
        matchEntireCell: true,
      },
      createExtra()
    );

    const branch = result.structuredContent?.result as Output | undefined;
    assert.equal(branch?.type, 'success');

    if (branch?.type === 'success') {
      // Should only match cell with exactly 'repl3-exact'
      assert.equal(branch.occurrencesChanged, 1, 'should replace 1 entire-cell occurrence');
    }
  });

  it('[repl4] searchByRegex: pattern replacement with backrefs', async () => {
    const result = await valuesReplaceHandler(
      {
        id: sharedSpreadsheetId,
        find: 'repl4-(user|admin)_(\\d+)',
        replacement: 'repl4-$1_id_$2',
        gid: sharedGid,
        searchByRegex: true,
      },
      createExtra()
    );

    const branch = result.structuredContent?.result as Output | undefined;
    assert.equal(branch?.type, 'success');

    if (branch?.type === 'success') {
      // Should match user_001, user_002, admin_001
      assert.equal(branch.occurrencesChanged, 3, 'should replace 3 regex matches');
    }
  });

  it('[repl5] includeFormulas: true modifies formula text', async () => {
    const result = await valuesReplaceHandler(
      {
        id: sharedSpreadsheetId,
        find: 'repl5-',
        replacement: 'repl5replaced-',
        gid: sharedGid,
        includeFormulas: true,
      },
      createExtra()
    );

    const branch = result.structuredContent?.result as Output | undefined;
    assert.equal(branch?.type, 'success');

    if (branch?.type === 'success') {
      // Exact counts: 2 plain cells + 1 formula
      assert.equal(branch.valuesChanged, 2, 'should change exactly 2 values');
      assert.equal(branch.formulasChanged, 1, 'should change 1 formula');
    }

    // READ-BACK VERIFICATION: Verify formula text was actually modified
    const sheets = google.sheets({ version: 'v4', auth });
    const readResult = await sheets.spreadsheets.values.get({
      spreadsheetId: sharedSpreadsheetId,
      range: 'Sheet1!A5:C5',
      valueRenderOption: 'FORMULA', // Get formula text, not calculated value
    });
    const row = readResult.data.values?.[0];
    assert.ok(row, 'should have data in row 5');
    assert.equal(row[0], 'repl5replaced-value', 'cell A5 should have replaced text');
    assert.equal(row[1], '=CONCAT("repl5replaced-","formula")', 'cell B5 formula should have replaced text');
    assert.equal(row[2], 'repl5replaced-other', 'cell C5 should have replaced text');
  });

  // ─────────────────────────────────────────────────────────────────
  // SCOPE TESTS
  // ─────────────────────────────────────────────────────────────────

  it('[repl6] no gid = allSheets replaces across all sheets', async () => {
    const result = await valuesReplaceHandler(
      {
        id: sharedSpreadsheetId,
        find: 'repl6-cross',
        replacement: 'repl6-found',
        // no gid = all sheets
      },
      createExtra()
    );

    const branch = result.structuredContent?.result as Output | undefined;
    assert.equal(branch?.type, 'success');

    if (branch?.type === 'success') {
      // Should find in both Sheet1 and Sheet2
      assert.equal(branch.sheetsChanged, 2, 'should change 2 sheets');
      assert.equal(branch.occurrencesChanged, 2, 'should replace 2 occurrences');
    }
  });

  it('[repl7] gid + range only replaces within bounds', async () => {
    const result = await valuesReplaceHandler(
      {
        id: sharedSpreadsheetId,
        find: 'repl7-in',
        replacement: 'repl7-OUT',
        gid: sharedGid,
        range: 'A7:B7',
      },
      createExtra()
    );

    const branch = result.structuredContent?.result as Output | undefined;
    assert.equal(branch?.type, 'success');

    if (branch?.type === 'success') {
      // Columns A and B have 'repl7-in*', Column C also has it but outside range
      assert.equal(branch.occurrencesChanged, 2, 'should replace 2 occurrences in range');
    }
  });

  // ─────────────────────────────────────────────────────────────────
  // EDGE CASES
  // ─────────────────────────────────────────────────────────────────

  it('[repl8] returns zero counts when no matches found', async () => {
    const result = await valuesReplaceHandler(
      {
        id: sharedSpreadsheetId,
        find: 'nonexistent-text-xyz-12345',
        replacement: 'replacement',
        gid: sharedGid,
      },
      createExtra()
    );

    const branch = result.structuredContent?.result as Output | undefined;
    assert.equal(branch?.type, 'success');

    if (branch?.type === 'success') {
      assert.equal(branch.occurrencesChanged, 0, 'should have 0 occurrences');
      assert.equal(branch.valuesChanged, 0, 'should change 0 values');
    }
  });

  it('[repl9] empty replacement deletes matched text', async () => {
    const result = await valuesReplaceHandler(
      {
        id: sharedSpreadsheetId,
        find: '-repl9-DELETE-',
        replacement: '',
        gid: sharedGid,
      },
      createExtra()
    );

    const branch = result.structuredContent?.result as Output | undefined;
    assert.equal(branch?.type, 'success');

    if (branch?.type === 'success') {
      assert.equal(branch.occurrencesChanged, 1, 'should delete 1 occurrence');
      // Cell now contains 'prefixsuffix'
    }
  });
});
