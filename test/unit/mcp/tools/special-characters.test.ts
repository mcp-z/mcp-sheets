import type { Logger, LoopbackOAuthProvider } from '@mcp-z/oauth-google';
import assert from 'assert';
import type { OAuth2Client } from 'google-auth-library';
import { google } from 'googleapis';
import createValuesReplaceTool, { type Input as ReplaceInput, type Output as ReplaceOutput } from '../../../../src/mcp/tools/values-replace.js';
import createSearchTool, { type Input as SearchInput, type Output as SearchOutput } from '../../../../src/mcp/tools/values-search.js';
import { createExtra, type TypedHandler } from '../../../lib/create-extra.js';
import createMiddlewareContext from '../../../lib/create-middleware-context.js';
import { createTestSpreadsheet, deleteTestSpreadsheet } from '../../../lib/spreadsheet-helpers.js';

/**
 * SPECIAL CHARACTERS TEST SUITE
 * =============================
 *
 * Tests to verify that special characters (colons, brackets, regex metacharacters)
 * are handled correctly in both search and replace operations.
 *
 * Background: An agent reported issues searching for strings like "HEADLINE:TITLE"
 * and "HEADLINE:PROFILE". This test suite verifies:
 *
 * 1. Colons work correctly in literal (non-regex) searches
 * 2. Colons work correctly in literal (non-regex) replacements
 * 3. RE2 regex metacharacters are properly escaped when searchByRegex is enabled
 * 4. Various special characters don't interfere with literal searches
 *
 * Test Data Layout (all in Row 1-12):
 * - Rows 1-2:  [colon-*] Tests with colons (e.g., "HEADLINE:TITLE")
 * - Rows 3-4:  [bracket-*] Tests with brackets (e.g., "[TAG]value")
 * - Rows 5-6:  [dot-*] Tests with dots/periods (e.g., "file.name.ext")
 * - Rows 7-8:  [regex-*] Tests with multiple regex metacharacters
 * - Rows 9-10: [pipe-*] Tests with pipe characters
 * - Rows 11-12: [mixed-*] Tests with mixed special characters
 */

describe('special characters handling (service-backed tests)', () => {
  let sharedSpreadsheetId: string;
  let sharedGid: string;
  let auth: OAuth2Client;
  let authProvider: LoopbackOAuthProvider;
  let logger: Logger;
  let accountId: string;
  let valuesReplaceHandler: TypedHandler<ReplaceInput>;
  let valuesSearchHandler: TypedHandler<SearchInput>;

  before(async () => {
    const middlewareContext = await createMiddlewareContext();
    authProvider = middlewareContext.authProvider;
    logger = middlewareContext.logger;
    auth = middlewareContext.auth;
    const middleware = middlewareContext.middleware;
    accountId = middlewareContext.accountId;

    const valuesReplaceTool = createValuesReplaceTool();
    const wrappedReplaceTool = middleware.withToolAuth(valuesReplaceTool);
    valuesReplaceHandler = wrappedReplaceTool.handler;

    const searchTool = createSearchTool();
    const wrappedSearchTool = middleware.withToolAuth(searchTool);
    valuesSearchHandler = wrappedSearchTool.handler;

    const title = `ci-special-chars-tests-${Date.now()}`;
    const accessToken = await authProvider.getAccessToken(accountId);
    sharedSpreadsheetId = await createTestSpreadsheet(accessToken, { title });
    sharedGid = '0';

    const sheets = google.sheets({ version: 'v4', auth });

    // Write all test data with special characters
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: sharedSpreadsheetId,
      requestBody: {
        valueInputOption: 'RAW', // RAW to prevent formula interpretation
        data: [
          {
            range: 'Sheet1!A1:C12',
            values: [
              // Rows 1-2: Colon tests (the original issue case)
              ['colon1-HEADLINE:TITLE', 'colon1-HEADLINE:PROFILE', 'colon1-OTHER'],
              ['colon2-scope:value', 'colon2-namespace:item:sub', 'colon2-plain'],
              // Rows 3-4: Bracket tests
              ['bracket1-[TAG]value', 'bracket1-(GROUP)item', 'bracket1-{SET}data'],
              ['bracket2-array[0]', 'bracket2-func(arg)', 'bracket2-obj{key}'],
              // Rows 5-6: Dot tests
              ['dot1-file.name.ext', 'dot1-192.168.1.1', 'dot1-no-dots'],
              ['dot2-a.b.c.d', 'dot2-version.1.0', 'dot2-plain'],
              // Rows 7-8: Multiple regex metacharacters
              ['regex1-$100.00+tax', 'regex1-a*b?c', 'regex1-x|y|z'],
              ['regex2-start^end$', 'regex2-back\\slash', 'regex2-plain'],
              // Rows 9-10: Pipe character tests
              ['pipe1-a|b|c', 'pipe1-left|right', 'pipe1-no-pipe'],
              ['pipe2-option1|option2', 'pipe2-cmd|filter', 'pipe2-plain'],
              // Rows 11-12: Mixed special characters
              ['mixed1-[TAG]:value.ext', 'mixed1-(a|b):$100', 'mixed1-plain'],
              ['mixed2-{key}:item[0]', 'mixed2-path/to/file.txt', 'mixed2-no-specials'],
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
  // COLON TESTS (Original reported issue)
  // ─────────────────────────────────────────────────────────────────

  describe('colon character (:)', () => {
    it('[colon-search] values-search finds text containing colons', async () => {
      const result = await valuesSearchHandler(
        {
          id: sharedSpreadsheetId,
          gid: sharedGid,
          query: 'HEADLINE:TITLE',
          select: 'cells',
          values: true,
          a1s: true,
        },
        createExtra()
      );

      const branch = result.structuredContent?.result as SearchOutput | undefined;
      assert.equal(branch?.type, 'success', 'search should succeed');

      if (branch?.type === 'success') {
        assert.equal(branch.count, 1, 'should find exactly 1 match for HEADLINE:TITLE');
        assert.ok(branch.values?.includes('colon1-HEADLINE:TITLE'), 'should include the matching cell value');
      }
    });

    it('[colon-search-multiple] values-search finds multiple colons', async () => {
      const result = await valuesSearchHandler(
        {
          id: sharedSpreadsheetId,
          gid: sharedGid,
          query: 'namespace:item:sub',
          select: 'cells',
          values: true,
        },
        createExtra()
      );

      const branch = result.structuredContent?.result as SearchOutput | undefined;
      assert.equal(branch?.type, 'success', 'search should succeed');

      if (branch?.type === 'success') {
        assert.equal(branch.count, 1, 'should find exactly 1 match for multiple colons');
      }
    });

    it('[colon-replace] values-replace works with colon patterns (literal)', async () => {
      const result = await valuesReplaceHandler(
        {
          id: sharedSpreadsheetId,
          gid: sharedGid,
          find: 'HEADLINE:PROFILE',
          replacement: 'basics',
        },
        createExtra()
      );

      const branch = result.structuredContent?.result as ReplaceOutput | undefined;
      assert.equal(branch?.type, 'success', 'replace should succeed');

      if (branch?.type === 'success') {
        assert.equal(branch.occurrencesChanged, 1, 'should replace 1 occurrence of HEADLINE:PROFILE');
      }

      // Verify the replacement actually happened
      const sheets = google.sheets({ version: 'v4', auth });
      const readResult = await sheets.spreadsheets.values.get({
        spreadsheetId: sharedSpreadsheetId,
        range: 'Sheet1!B1',
      });
      assert.equal(readResult.data.values?.[0]?.[0], 'colon1-basics', 'cell should contain replaced text');
    });

    it('[colon-replace-partial] values-replace handles partial colon matches', async () => {
      const result = await valuesReplaceHandler(
        {
          id: sharedSpreadsheetId,
          gid: sharedGid,
          find: 'scope:value',
          replacement: 'scope:REPLACED',
        },
        createExtra()
      );

      const branch = result.structuredContent?.result as ReplaceOutput | undefined;
      assert.equal(branch?.type, 'success', 'replace should succeed');

      if (branch?.type === 'success') {
        assert.equal(branch.occurrencesChanged, 1, 'should replace 1 occurrence');
      }
    });
  });

  // ─────────────────────────────────────────────────────────────────
  // BRACKET TESTS
  // ─────────────────────────────────────────────────────────────────

  describe('bracket characters ([], (), {})', () => {
    it('[bracket-search-square] values-search finds text with square brackets', async () => {
      const result = await valuesSearchHandler(
        {
          id: sharedSpreadsheetId,
          gid: sharedGid,
          query: '[TAG]',
          select: 'cells',
          values: true,
        },
        createExtra()
      );

      const branch = result.structuredContent?.result as SearchOutput | undefined;
      assert.equal(branch?.type, 'success', 'search should succeed');

      if (branch?.type === 'success') {
        assert.ok(branch.count >= 1, 'should find at least 1 match with square brackets');
      }
    });

    it('[bracket-replace-square] values-replace works with square brackets (literal)', async () => {
      const result = await valuesReplaceHandler(
        {
          id: sharedSpreadsheetId,
          gid: sharedGid,
          find: 'array[0]',
          replacement: 'array[REPLACED]',
        },
        createExtra()
      );

      const branch = result.structuredContent?.result as ReplaceOutput | undefined;
      assert.equal(branch?.type, 'success', 'replace should succeed');

      if (branch?.type === 'success') {
        assert.equal(branch.occurrencesChanged, 1, 'should replace 1 occurrence with brackets');
      }
    });

    it('[bracket-replace-regex-escaped] values-replace with regex requires escaped brackets', async () => {
      // When using regex mode, brackets have special meaning and must be escaped
      const result = await valuesReplaceHandler(
        {
          id: sharedSpreadsheetId,
          gid: sharedGid,
          find: '\\[TAG\\]value', // Escaped brackets for regex
          replacement: '[NEWTAG]value',
          searchByRegex: true,
        },
        createExtra()
      );

      const branch = result.structuredContent?.result as ReplaceOutput | undefined;
      assert.equal(branch?.type, 'success', 'replace should succeed');

      if (branch?.type === 'success') {
        assert.equal(branch.occurrencesChanged, 1, 'should replace 1 occurrence with escaped brackets in regex');
      }
    });
  });

  // ─────────────────────────────────────────────────────────────────
  // DOT/PERIOD TESTS
  // ─────────────────────────────────────────────────────────────────

  describe('dot/period character (.)', () => {
    it('[dot-search] values-search finds text with dots (literal)', async () => {
      const result = await valuesSearchHandler(
        {
          id: sharedSpreadsheetId,
          gid: sharedGid,
          query: '192.168.1.1',
          select: 'cells',
          values: true,
        },
        createExtra()
      );

      const branch = result.structuredContent?.result as SearchOutput | undefined;
      assert.equal(branch?.type, 'success', 'search should succeed');

      if (branch?.type === 'success') {
        assert.equal(branch.count, 1, 'should find exactly 1 IP address match');
      }
    });

    it('[dot-replace-literal] values-replace with dots works in literal mode', async () => {
      const result = await valuesReplaceHandler(
        {
          id: sharedSpreadsheetId,
          gid: sharedGid,
          find: 'version.1.0',
          replacement: 'version.2.0',
        },
        createExtra()
      );

      const branch = result.structuredContent?.result as ReplaceOutput | undefined;
      assert.equal(branch?.type, 'success', 'replace should succeed');

      if (branch?.type === 'success') {
        assert.equal(branch.occurrencesChanged, 1, 'should replace 1 occurrence with dots');
      }
    });

    it('[dot-replace-regex-unescaped] dots in regex mode match any character', async () => {
      // In regex mode, unescaped dots match any character
      // "a.b.c.d" with pattern "a.b" would match "aXb" too if present
      const result = await valuesReplaceHandler(
        {
          id: sharedSpreadsheetId,
          gid: sharedGid,
          find: 'a.b.c.d', // Unescaped - dots match any char
          replacement: 'MATCHED',
          searchByRegex: true,
        },
        createExtra()
      );

      const branch = result.structuredContent?.result as ReplaceOutput | undefined;
      assert.equal(branch?.type, 'success', 'replace should succeed');

      if (branch?.type === 'success') {
        // Should match "a.b.c.d" because dots match the actual dots
        assert.equal(branch.occurrencesChanged, 1, 'regex with dots should match');
      }
    });
  });

  // ─────────────────────────────────────────────────────────────────
  // PIPE CHARACTER TESTS
  // ─────────────────────────────────────────────────────────────────

  describe('pipe character (|)', () => {
    it('[pipe-search] values-search finds text with pipes (literal)', async () => {
      const result = await valuesSearchHandler(
        {
          id: sharedSpreadsheetId,
          gid: sharedGid,
          query: 'a|b|c',
          select: 'cells',
          values: true,
        },
        createExtra()
      );

      const branch = result.structuredContent?.result as SearchOutput | undefined;
      assert.equal(branch?.type, 'success', 'search should succeed');

      if (branch?.type === 'success') {
        assert.equal(branch.count, 1, 'should find exactly 1 match with pipes');
      }
    });

    it('[pipe-replace-literal] values-replace with pipes works in literal mode', async () => {
      const result = await valuesReplaceHandler(
        {
          id: sharedSpreadsheetId,
          gid: sharedGid,
          find: 'left|right',
          replacement: 'LEFT|RIGHT',
        },
        createExtra()
      );

      const branch = result.structuredContent?.result as ReplaceOutput | undefined;
      assert.equal(branch?.type, 'success', 'replace should succeed');

      if (branch?.type === 'success') {
        assert.equal(branch.occurrencesChanged, 1, 'should replace 1 occurrence with pipes');
      }
    });

    it('[pipe-replace-regex-escaped] pipes in regex mode need escaping', async () => {
      // In regex mode, pipe means alternation
      // To match literal pipe, it must be escaped
      const result = await valuesReplaceHandler(
        {
          id: sharedSpreadsheetId,
          gid: sharedGid,
          find: 'option1\\|option2', // Escaped pipe
          replacement: 'OPTION_MATCHED',
          searchByRegex: true,
        },
        createExtra()
      );

      const branch = result.structuredContent?.result as ReplaceOutput | undefined;
      assert.equal(branch?.type, 'success', 'replace should succeed');

      if (branch?.type === 'success') {
        assert.equal(branch.occurrencesChanged, 1, 'should match literal pipe when escaped');
      }
    });
  });

  // ─────────────────────────────────────────────────────────────────
  // MIXED SPECIAL CHARACTERS TESTS
  // ─────────────────────────────────────────────────────────────────

  describe('mixed special characters', () => {
    it('[mixed-search] values-search finds complex patterns', async () => {
      const result = await valuesSearchHandler(
        {
          id: sharedSpreadsheetId,
          gid: sharedGid,
          query: '[TAG]:value.ext',
          select: 'cells',
          values: true,
        },
        createExtra()
      );

      const branch = result.structuredContent?.result as SearchOutput | undefined;
      assert.equal(branch?.type, 'success', 'search should succeed');

      if (branch?.type === 'success') {
        assert.equal(branch.count, 1, 'should find complex pattern with brackets, colon, and dot');
      }
    });

    it('[mixed-replace] values-replace handles complex patterns (literal)', async () => {
      const result = await valuesReplaceHandler(
        {
          id: sharedSpreadsheetId,
          gid: sharedGid,
          find: '{key}:item[0]',
          replacement: '{newkey}:newitem[0]',
        },
        createExtra()
      );

      const branch = result.structuredContent?.result as ReplaceOutput | undefined;
      assert.equal(branch?.type, 'success', 'replace should succeed');

      if (branch?.type === 'success') {
        assert.equal(branch.occurrencesChanged, 1, 'should replace complex pattern');
      }
    });

    it('[mixed-replace-dollar] values-replace handles dollar sign (literal)', async () => {
      const result = await valuesReplaceHandler(
        {
          id: sharedSpreadsheetId,
          gid: sharedGid,
          find: '$100.00+tax',
          replacement: '$200.00+tax',
        },
        createExtra()
      );

      const branch = result.structuredContent?.result as ReplaceOutput | undefined;
      assert.equal(branch?.type, 'success', 'replace should succeed');

      if (branch?.type === 'success') {
        assert.equal(branch.occurrencesChanged, 1, 'should replace pattern with dollar sign');
      }
    });
  });

  // ─────────────────────────────────────────────────────────────────
  // CASE SENSITIVITY TESTS
  // ─────────────────────────────────────────────────────────────────

  describe('case sensitivity (matchCase option)', () => {
    it('[case-insensitive-default] values-search is case-insensitive by default', async () => {
      // Search for lowercase should find uppercase HEADLINE:TITLE
      const result = await valuesSearchHandler(
        {
          id: sharedSpreadsheetId,
          gid: sharedGid,
          query: 'headline:title', // lowercase
          select: 'cells',
          values: true,
        },
        createExtra()
      );

      const branch = result.structuredContent?.result as SearchOutput | undefined;
      assert.equal(branch?.type, 'success', 'search should succeed');

      if (branch?.type === 'success') {
        assert.ok(branch.count >= 1, 'should find match case-insensitively');
      }
    });

    it('[case-sensitive-enabled] values-search with matchCase=true is case-sensitive', async () => {
      // Search for lowercase with matchCase=true should NOT find uppercase
      const result = await valuesSearchHandler(
        {
          id: sharedSpreadsheetId,
          gid: sharedGid,
          query: 'headline:title', // lowercase - won't match HEADLINE:TITLE
          select: 'cells',
          values: true,
          matchCase: true,
        },
        createExtra()
      );

      const branch = result.structuredContent?.result as SearchOutput | undefined;
      assert.equal(branch?.type, 'success', 'search should succeed');

      if (branch?.type === 'success') {
        // Should NOT find the uppercase version
        assert.equal(branch.count, 0, 'should not find match when case differs and matchCase=true');
      }
    });

    it('[case-sensitive-exact] values-search with matchCase=true finds exact case', async () => {
      // Search for exact case should find match
      const result = await valuesSearchHandler(
        {
          id: sharedSpreadsheetId,
          gid: sharedGid,
          query: 'HEADLINE:TITLE', // exact case
          select: 'cells',
          values: true,
          matchCase: true,
        },
        createExtra()
      );

      const branch = result.structuredContent?.result as SearchOutput | undefined;
      assert.equal(branch?.type, 'success', 'search should succeed');

      if (branch?.type === 'success') {
        assert.equal(branch.count, 1, 'should find exact case match');
      }
    });
  });

  // ─────────────────────────────────────────────────────────────────
  // REGEX VS LITERAL MODE COMPARISON
  // ─────────────────────────────────────────────────────────────────

  describe('regex vs literal mode behavior', () => {
    it('[regex-literal-comparison] same pattern behaves differently in regex vs literal', async () => {
      // "a*b" in literal mode matches the string "a*b"
      // "a*b" in regex mode means "zero or more 'a' followed by 'b'"

      // First, do a literal search to confirm the test data exists
      const searchResult = await valuesSearchHandler(
        {
          id: sharedSpreadsheetId,
          gid: sharedGid,
          query: 'a*b?c',
          select: 'cells',
          values: true,
        },
        createExtra()
      );

      const searchBranch = searchResult.structuredContent?.result as SearchOutput | undefined;
      assert.equal(searchBranch?.type, 'success', 'search should succeed');
      if (searchBranch?.type === 'success') {
        assert.equal(searchBranch.count, 1, 'should find the literal string a*b?c');
      }

      // Now verify regex mode with escaped metacharacters works
      const regexResult = await valuesReplaceHandler(
        {
          id: sharedSpreadsheetId,
          gid: sharedGid,
          find: 'a\\*b\\?c', // Escaped for regex literal match
          replacement: 'REGEX_MATCHED',
          searchByRegex: true,
        },
        createExtra()
      );

      const regexBranch = regexResult.structuredContent?.result as ReplaceOutput | undefined;
      assert.equal(regexBranch?.type, 'success', 'regex replace should succeed');
      if (regexBranch?.type === 'success') {
        assert.equal(regexBranch.occurrencesChanged, 1, 'should match with escaped regex metacharacters');
      }
    });
  });
});
