import '../../../lib/env-loader.ts';
import { mcp } from '@mcp-z/mcp-sheets';
import { ProtocolError, ProtocolErrorCode } from '@mcp-z/server';
import assert from 'assert';
import crypto from 'crypto';
import fs from 'fs/promises';
import { safeRmSync } from 'fs-remove-compat';
import * as path from 'path';
import type { Input, Output } from '../../../../src/mcp/tools/csv-get-columns.ts';
import { createExtra, type TypedHandler } from '../../../lib/create-extra.ts';

let handler: TypedHandler<Input>;
let tmpDir: string;

before(async () => {
  try {
    // Create temporary directory
    tmpDir = path.join('.tmp', `csv-get-columns-tests-${crypto.randomUUID()}`);
    await fs.mkdir(tmpDir, { recursive: true });

    handler = mcp.toolFactories.csvGetColumns().handler;
  } catch (error) {
    console.error('Failed to initialize test resources:', error);
    throw error;
  }
});

after(async () => {
  // Cleanup
  if (tmpDir) safeRmSync(tmpDir, { recursive: true, force: true });
});

it('sheets-csv-get-columns reads columns from CSV with data', async () => {
  // Create test CSV with column names
  const csvPath = path.join(tmpDir, 'test-data.csv');
  const csvContent = 'Email,Name,Phone\njohn@example.com,John Doe,555-1234\njane@example.com,Jane Smith,555-5678';
  await fs.writeFile(csvPath, csvContent);

  const sourceUri = `file://${csvPath}`;
  const resp = await handler({ sourceUri }, createExtra());

  const structured = (resp.structuredContent as { result?: unknown } | undefined)?.result as Output | undefined;
  assert.equal(structured?.type, 'success', 'expected success type');
  if (structured?.type === 'success') {
    assert.deepEqual(structured.columns, ['Email', 'Name', 'Phone'], 'columns mismatch');
    assert.equal(structured.isEmpty, false, 'expected isEmpty to be false');
  }
});

it('sheets-csv-get-columns returns empty for empty CSV', async () => {
  // Create empty CSV
  const csvPath = path.join(tmpDir, 'empty.csv');
  await fs.writeFile(csvPath, '');

  const sourceUri = `file://${csvPath}`;
  const resp = await handler({ sourceUri }, createExtra());

  const structured = (resp.structuredContent as { result?: unknown } | undefined)?.result as Output | undefined;
  assert.equal(structured?.type, 'success', 'expected success type');
  if (structured?.type === 'success') {
    assert.deepEqual(structured.columns, [], 'expected empty columns');
    assert.equal(structured.isEmpty, true, 'expected isEmpty to be true');
  }
});

it('sheets-csv-get-columns reads only first row (not all data)', async () => {
  // Create CSV with many rows
  const csvPath = path.join(tmpDir, 'large.csv');
  const rows = ['Header1,Header2,Header3'];
  for (let i = 0; i < 300000; i++) {
    rows.push(`value${i}a,value${i}b,value${i}c`);
  }
  rows.push('"unterminated');
  await fs.writeFile(csvPath, rows.join('\n'));

  const sourceUri = `file://${csvPath}`;
  const resp = await handler({ sourceUri }, createExtra());

  const structured = (resp.structuredContent as { result?: unknown } | undefined)?.result as Output | undefined;
  assert.equal(structured?.type, 'success', 'expected success type');
  if (structured?.type === 'success') {
    assert.deepEqual(structured.columns, ['Header1', 'Header2', 'Header3'], 'columns mismatch');
    assert.equal(structured.isEmpty, false, 'expected isEmpty to be false');
  }
  safeRmSync(csvPath, { force: true });
  await assert.rejects(fs.stat(csvPath), { code: 'ENOENT' });
});

it('sheets-csv-get-columns handles CSV with only columns (no data rows)', async () => {
  // Create CSV with only column row
  const csvPath = path.join(tmpDir, 'columns-only.csv');
  await fs.writeFile(csvPath, 'Col1,Col2,Col3');

  const sourceUri = `file://${csvPath}`;
  const resp = await handler({ sourceUri }, createExtra());

  const structured = (resp.structuredContent as { result?: unknown } | undefined)?.result as Output | undefined;
  assert.equal(structured?.type, 'success', 'expected success type');
  if (structured?.type === 'success') {
    assert.deepEqual(structured.columns, ['Col1', 'Col2', 'Col3'], 'columns mismatch');
    assert.equal(structured.isEmpty, false, 'expected isEmpty to be false');
  }
});

it('sheets-csv-get-columns preserves a malformed first-record failure and releases its file', async () => {
  const csvPath = path.join(tmpDir, 'malformed.csv');
  await fs.writeFile(csvPath, '"unterminated');
  await assert.rejects(handler({ sourceUri: `file://${csvPath}` }, createExtra()), (error: unknown) => {
    assert.ok(error instanceof ProtocolError);
    assert.equal(error.code, ProtocolErrorCode.InternalError);
    assert.match(error.message, /Error getting CSV columns:.*Quote Not Closed/);
    return true;
  });
  safeRmSync(csvPath, { force: true });
  await assert.rejects(fs.stat(csvPath), { code: 'ENOENT' });
});

it('sheets-csv-get-columns reports a missing source through ProtocolError', async () => {
  const csvPath = path.join(tmpDir, 'missing.csv');
  await assert.rejects(handler({ sourceUri: `file://${csvPath}` }, createExtra()), (error: unknown) => {
    assert.ok(error instanceof ProtocolError);
    assert.equal(error.code, ProtocolErrorCode.InternalError);
    assert.match(error.message, /Error getting CSV columns:.*ENOENT/);
    assert.ok(error.message.includes('missing.csv'));
    return true;
  });
});
