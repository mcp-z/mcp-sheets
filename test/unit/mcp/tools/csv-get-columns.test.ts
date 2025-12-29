import assert from 'assert';
import crypto from 'crypto';
import fs from 'fs/promises';
import * as path from 'path';
import createCsvGetColumnsTool, { type Input, type Output } from '../../../../src/mcp/tools/csv-get-columns.js';
import { createExtra, type TypedHandler } from '../../../lib/create-extra.js';
import createMiddlewareContext from '../../../lib/create-middleware-context.js';

let handler: TypedHandler<Input>;
let tmpDir: string;

before(async () => {
  try {
    // Create temporary directory
    tmpDir = path.join('.tmp', `csv-get-columns-tests-${crypto.randomUUID()}`);
    await fs.mkdir(tmpDir, { recursive: true });

    // Create tool with middleware
    const middlewareContext = await createMiddlewareContext();
    const middleware = middlewareContext.middleware;
    const tool = createCsvGetColumnsTool();
    const wrappedTool = middleware.withToolAuth(tool);
    handler = wrappedTool.handler;
  } catch (error) {
    console.error('Failed to initialize test resources:', error);
    throw error;
  }
});

after(async () => {
  // Cleanup
  await fs.rm(tmpDir, { recursive: true, force: true });
});

it('sheets-csv-get-columns reads columns from CSV with data', async () => {
  // Create test CSV with column names
  const csvPath = path.join(tmpDir, 'test-data.csv');
  const csvContent = 'Email,Name,Phone\njohn@example.com,John Doe,555-1234\njane@example.com,Jane Smith,555-5678';
  await fs.writeFile(csvPath, csvContent);

  const sourceUri = `file://${csvPath}`;
  const resp = await handler({ sourceUri }, createExtra());

  const structured = resp.structuredContent?.result as Output | undefined;
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

  const structured = resp.structuredContent?.result as Output | undefined;
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
  for (let i = 0; i < 1000; i++) {
    rows.push(`value${i}a,value${i}b,value${i}c`);
  }
  await fs.writeFile(csvPath, rows.join('\n'));

  const sourceUri = `file://${csvPath}`;
  const resp = await handler({ sourceUri }, createExtra());

  const structured = resp.structuredContent?.result as Output | undefined;
  assert.equal(structured?.type, 'success', 'expected success type');
  if (structured?.type === 'success') {
    assert.deepEqual(structured.columns, ['Header1', 'Header2', 'Header3'], 'columns mismatch');
    assert.equal(structured.isEmpty, false, 'expected isEmpty to be false');
  }
});

it('sheets-csv-get-columns handles CSV with only columns (no data rows)', async () => {
  // Create CSV with only column row
  const csvPath = path.join(tmpDir, 'columns-only.csv');
  await fs.writeFile(csvPath, 'Col1,Col2,Col3');

  const sourceUri = `file://${csvPath}`;
  const resp = await handler({ sourceUri }, createExtra());

  const structured = resp.structuredContent?.result as Output | undefined;
  assert.equal(structured?.type, 'success', 'expected success type');
  if (structured?.type === 'success') {
    assert.deepEqual(structured.columns, ['Col1', 'Col2', 'Col3'], 'columns mismatch');
    assert.equal(structured.isEmpty, false, 'expected isEmpty to be false');
  }
});
