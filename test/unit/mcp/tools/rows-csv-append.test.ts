import type { Logger, LoopbackOAuthProvider } from '@mcp-z/oauth-google';
import assert from 'assert';
import crypto from 'crypto';
import fs from 'fs/promises';
import * as path from 'path';
import createRowsCsvAppendTool, { type Input, type Output } from '../../../../src/mcp/tools/rows-csv-append.ts';
import { createExtra, type TypedHandler } from '../../../lib/create-extra.ts';
import createMiddlewareContext from '../../../lib/create-middleware-context.ts';
import { createTestSheet, createTestSpreadsheet, deleteTestSpreadsheet } from '../../../lib/spreadsheet-helpers.ts';

let handler: TypedHandler<Input>;
let sharedSpreadsheetId: string;
let authProvider: LoopbackOAuthProvider;
let accountId: string;
let logger: Logger;
let tmpDir: string;
let deduplicationTestGid: string; // Separate sheet for deduplication test

before(async () => {
  try {
    // Create temporary directory
    tmpDir = path.join('.tmp', `rows-csv-append-tests-${crypto.randomUUID()}`);
    await fs.mkdir(tmpDir, { recursive: true });

    // Create middleware and tools
    const middlewareContext = await createMiddlewareContext();
    authProvider = middlewareContext.authProvider;
    logger = middlewareContext.logger;
    const middleware = middlewareContext.middleware;
    accountId = middlewareContext.accountId;

    const tool = createRowsCsvAppendTool();
    const wrappedTool = middleware.withToolAuth(tool);
    handler = wrappedTool.handler;

    // Create shared spreadsheet
    const accessToken = await authProvider.getAccessToken(accountId);
    const title = `ci-rows-csv-append-tests-${Date.now()}`;
    sharedSpreadsheetId = await createTestSpreadsheet(accessToken, { title });

    // Create separate sheet for deduplication test (for test isolation)
    const dedupSheetId = await createTestSheet(accessToken, sharedSpreadsheetId, { title: 'deduplication-test' });
    deduplicationTestGid = String(dedupSheetId);
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

it('sheets-rows-csv-append imports CSV with header mapping (string names)', async () => {
  // Create test CSV
  const csvPath = path.join(tmpDir, 'contacts.csv');
  const csvContent = 'Email Address,Full Name,Phone Number\njohn@example.com,John Doe,555-1234\njane@example.com,Jane Smith,555-5678';
  await fs.writeFile(csvPath, csvContent);

  const sourceUri = `file://${csvPath}`;
  const resp = await handler(
    {
      id: sharedSpreadsheetId,
      gid: '0',
      sourceUri,
      sourceHasHeaders: true,
      headerMap: [
        { source: 'Email Address', target: 'email' },
        { source: 'Full Name', target: 'name' },
        { source: 'Phone Number', target: 'phone' },
      ],
    },
    createExtra()
  );

  const structured = resp.structuredContent?.result as Output | undefined;
  if (structured?.type !== 'success') {
    assert.fail('Operation failed');
  }
  assert.equal(structured.updatedRows, 2, 'expected 2 rows updated');
  assert.equal(structured.rowsSkipped, 0, 'expected 0 rows skipped');
});

it('sheets-rows-csv-append supports numeric indices', async () => {
  // Create test CSV without specific header names
  const csvPath = path.join(tmpDir, 'data-indices.csv');
  const csvContent = 'col0,col1,col2\nval0a,val1a,val2a\nval0b,val1b,val2b';
  await fs.writeFile(csvPath, csvContent);

  const sourceUri = `file://${csvPath}`;
  const resp = await handler(
    {
      id: sharedSpreadsheetId,
      gid: '0',
      sourceUri,
      sourceHasHeaders: true,
      headerMap: [
        { source: 0, target: 'field_0' }, // CSV col 0 -> sheet "field_0"
        { source: 1, target: 'field_1' }, // CSV col 1 -> sheet "field_1"
        { source: 2, target: 'field_2' }, // CSV col 2 -> sheet "field_2"
      ],
    },
    createExtra()
  );

  const structured = resp.structuredContent?.result as Output | undefined;
  if (structured?.type !== 'success') {
    assert.fail('Operation failed');
  }
  assert.equal(structured.updatedRows, 2, 'expected 2 rows updated');
});

it('sheets-rows-csv-append deduplicates by column name', async () => {
  // Create CSV with duplicate email
  const csvPath = path.join(tmpDir, 'contacts-dup.csv');
  const csvContent = 'Email,Name\njohn@example.com,John Doe\njane@example.com,Jane Smith\njohn@example.com,John Duplicate';
  await fs.writeFile(csvPath, csvContent);

  // First import - using isolated sheet to avoid test pollution
  const sourceUri = `file://${csvPath}`;
  const resp1 = await handler(
    {
      id: sharedSpreadsheetId,
      gid: deduplicationTestGid,
      sourceUri,
      sourceHasHeaders: true,
      headerMap: [
        { source: 'Email', target: 'email' },
        { source: 'Name', target: 'name' },
      ],
      deduplicateBy: ['email'],
    },
    createExtra()
  );

  const structured1 = resp1.structuredContent?.result as Output | undefined;
  if (structured1?.type !== 'success') {
    assert.fail('First import failed');
  }
  assert.equal(structured1.updatedRows, 2, 'expected 2 unique rows in first import');
  assert.equal(structured1.rowsSkipped, 1, 'expected 1 duplicate skipped in first import');

  // Second import (should skip all as duplicates)
  const resp2 = await handler(
    {
      id: sharedSpreadsheetId,
      gid: deduplicationTestGid,
      sourceUri,
      sourceHasHeaders: true,
      headerMap: [
        { source: 'Email', target: 'email' },
        { source: 'Name', target: 'name' },
      ],
      deduplicateBy: ['email'],
    },
    createExtra()
  );

  const structured2 = resp2.structuredContent?.result as Output | undefined;
  if (structured2?.type !== 'success') {
    assert.fail('Second import failed');
  }
  assert.equal(structured2.updatedRows, 0, 'expected 0 rows in second import');
  assert.equal(structured2.rowsSkipped, 3, 'expected all 3 rows skipped in second import');
});

it('sheets-rows-csv-append supports data-only mode (sourceHasHeaders=false)', async () => {
  // Create data-only CSV (no headers)
  const csvPath = path.join(tmpDir, 'raw-data.csv');
  const csvContent = 'data0a,data0b,data0c\ndata1a,data1b,data1c';
  await fs.writeFile(csvPath, csvContent);

  const sourceUri = `file://${csvPath}`;
  const resp = await handler(
    {
      id: sharedSpreadsheetId,
      gid: '0',
      sourceUri,
      sourceHasHeaders: false,
      headerMap: [
        { source: 0, target: 0 }, // Must use numeric indices
        { source: 1, target: 1 },
        { source: 2, target: 2 },
      ],
    },
    createExtra()
  );

  const structured = resp.structuredContent?.result as Output | undefined;
  if (structured?.type !== 'success') {
    assert.fail('Operation failed');
  }
  assert.equal(structured.updatedRows, 2, 'expected 2 rows updated');
});

it('sheets-rows-csv-append rejects string refs when sourceHasHeaders=false', async () => {
  // Create data-only CSV
  const csvPath = path.join(tmpDir, 'invalid-refs.csv');
  const csvContent = 'data0,data1\ndata2,data3';
  await fs.writeFile(csvPath, csvContent);

  const sourceUri = `file://${csvPath}`;

  // Should throw validation error
  await assert.rejects(
    async () => {
      await handler(
        {
          id: sharedSpreadsheetId,
          gid: '0',
          sourceUri,
          sourceHasHeaders: false,
          headerMap: [
            { source: 'Email', target: 0 }, // Invalid: string source when sourceHasHeaders=false
          ],
        },
        createExtra()
      );
    },
    (error: unknown) => {
      assert.ok(error instanceof Error && error.message.includes('sourceHasHeaders=false requires numeric indices'), 'expected validation error message');
      return true;
    }
  );
});

it('sheets-rows-csv-append omits unmapped columns', async () => {
  // Create CSV with extra column
  const csvPath = path.join(tmpDir, 'with-extra.csv');
  const csvContent = 'Email,Name,Internal,Phone\njohn@example.com,John,secret123,555-1234';
  await fs.writeFile(csvPath, csvContent);

  const sourceUri = `file://${csvPath}`;
  const resp = await handler(
    {
      id: sharedSpreadsheetId,
      gid: '0',
      sourceUri,
      sourceHasHeaders: true,
      headerMap: [
        { source: 'Email', target: 'email' },
        { source: 'Name', target: 'name' },
        { source: 'Phone', target: 'phone' },
        // "Internal" column omitted - should be ignored
      ],
    },
    createExtra()
  );

  const structured = resp.structuredContent?.result as Output | undefined;
  if (structured?.type !== 'success') {
    assert.fail('Operation failed');
  }
  assert.equal(structured.updatedRows, 1, 'expected 1 row updated');
});

// Edge Case Tests

it('sheets-rows-csv-append supports mixed type headerMap (strings and numbers)', async () => {
  // Create CSV
  const csvPath = path.join(tmpDir, 'mixed-types.csv');
  const csvContent = 'Email,Name,Phone\ntest@example.com,Test User,555-9999';
  await fs.writeFile(csvPath, csvContent);

  const sourceUri = `file://${csvPath}`;
  const resp = await handler(
    {
      id: sharedSpreadsheetId,
      gid: '0',
      sourceUri,
      sourceHasHeaders: true,
      headerMap: [
        { source: 'Email', target: 'email' }, // string → string
        { source: 1, target: 'name' }, // number → string (CSV col 1)
        { source: 'Phone', target: 2 }, // string → number (sheet col 2)
      ],
    },
    createExtra()
  );

  const structured = resp.structuredContent?.result as Output | undefined;
  if (structured?.type !== 'success') {
    assert.fail('Operation failed');
  }
  assert.equal(structured.updatedRows, 1, 'expected 1 row updated');
});

it('sheets-rows-csv-append supports composite key deduplication', async () => {
  // Create CSV with composite key (provider + id)
  const csvPath = path.join(tmpDir, 'composite-key.csv');
  const csvContent = 'Provider,ID,Data\nGmail,msg123,Data1\nOutlook,msg123,Data2\nGmail,msg123,Duplicate';
  await fs.writeFile(csvPath, csvContent);

  const sourceUri = `file://${csvPath}`;
  const resp = await handler(
    {
      id: sharedSpreadsheetId,
      gid: '0',
      sourceUri,
      sourceHasHeaders: true,
      headerMap: [
        { source: 'Provider', target: 'provider' },
        { source: 'ID', target: 'id' },
        { source: 'Data', target: 'data' },
      ],
      deduplicateBy: ['provider', 'id'], // Composite key
    },
    createExtra()
  );

  const structured = resp.structuredContent?.result as Output | undefined;
  if (structured?.type !== 'success') {
    assert.fail('Operation failed');
  }
  assert.equal(structured?.updatedRows, 2, 'expected 2 unique rows (Gmail+msg123, Outlook+msg123)');
  assert.equal(structured?.rowsSkipped, 1, 'expected 1 duplicate (Gmail+msg123)');
});

it('sheets-rows-csv-append handles large CSV (streaming verification)', async () => {
  // Create CSV with 1000+ rows
  const csvPath = path.join(tmpDir, 'large.csv');
  const rows = ['Email,Name,Index'];
  for (let i = 0; i < 1500; i++) {
    rows.push(`user${i}@example.com,User ${i},${i}`);
  }
  await fs.writeFile(csvPath, rows.join('\n'));

  const sourceUri = `file://${csvPath}`;
  const resp = await handler(
    {
      id: sharedSpreadsheetId,
      gid: '0',
      sourceUri,
      sourceHasHeaders: true,
      headerMap: [
        { source: 'Email', target: 'email' },
        { source: 'Name', target: 'name' },
        { source: 'Index', target: 'index' },
      ],
    },
    createExtra()
  );

  const structured = resp.structuredContent?.result as Output | undefined;
  if (structured?.type !== 'success') {
    assert.fail('Operation failed');
  }
  assert.equal(structured?.updatedRows, 1500, 'expected 1500 rows updated');
});

it('sheets-rows-csv-append handles header not found error', async () => {
  // Create CSV
  const csvPath = path.join(tmpDir, 'wrong-headers.csv');
  const csvContent = 'Email,Name\ntest@example.com,Test';
  await fs.writeFile(csvPath, csvContent);

  const sourceUri = `file://${csvPath}`;

  // Should throw error
  await assert.rejects(
    async () => {
      await handler(
        {
          id: sharedSpreadsheetId,
          gid: '0',
          sourceUri,
          sourceHasHeaders: true,
          headerMap: [
            { source: 'Email', target: 'email' },
            { source: 'NonExistent', target: 'name' }, // Header doesn't exist
          ],
        },
        createExtra()
      );
    },
    (error: unknown) => {
      assert.ok(error instanceof Error && error.message.includes('NonExistent'), 'expected error about missing header');
      return true;
    }
  );
});

it('sheets-rows-csv-append handles column index out of bounds', async () => {
  // Create CSV with only 2 columns
  const csvPath = path.join(tmpDir, 'out-of-bounds.csv');
  const csvContent = 'Col1,Col2\nval1,val2';
  await fs.writeFile(csvPath, csvContent);

  const sourceUri = `file://${csvPath}`;
  const resp = await handler(
    {
      id: sharedSpreadsheetId,
      gid: '0',
      sourceUri,
      sourceHasHeaders: true,
      headerMap: [
        { source: 0, target: 'col1' },
        { source: 1, target: 'col2' },
        { source: 5, target: 'col3' }, // Index out of bounds
      ],
    },
    createExtra()
  );

  // Should handle gracefully (fill with null)
  const structured = resp.structuredContent?.result as Output | undefined;
  if (structured?.type !== 'success') {
    assert.fail('Operation failed');
  }
});

it('sheets-rows-csv-append supports file:// URIs', async () => {
  const csvContent = 'Name,Age,City\nAlice,30,NYC\nBob,25,SF';
  const csvPath = path.join(tmpDir, 'file-uri-test.csv');
  await fs.writeFile(csvPath, csvContent);

  const sourceUri = `file://${csvPath}`;

  const resp = await handler(
    {
      id: sharedSpreadsheetId,
      gid: '0',
      sourceUri,
      sourceHasHeaders: true,
      headerMap: [
        { source: 'Name', target: 'name' },
        { source: 'Age', target: 'age' },
        { source: 'City', target: 'city' },
      ],
    },
    createExtra()
  );

  const structured = resp.structuredContent?.result as Output | undefined;
  if (structured?.type !== 'success') {
    assert.fail('Operation failed');
  }
  assert.equal(structured.updatedRows, 2, 'expected 2 rows updated');
});
