import '../../lib/env-loader.ts';
import assert from 'assert';
import { snapshotHeaderAndKeys, snapshotHeaderKeysAndPositions } from '../../../src/spreadsheet/data-operations.ts';

type GetParams = { range: string };

/** Plain parameter object returning fixed data: `keyRows` are the sheet's data rows (row 2 onward), `rowCount` the grid size. */
function makeSheets(keyRows: string[][], opts: { rowCount?: number; failKeyRead?: unknown; omitTrailingEmpty?: boolean } = {}) {
  const rowCount = opts.rowCount ?? keyRows.length + 1;
  const ranges: string[] = [];
  const sheets = {
    spreadsheets: {
      get: async () => ({ data: { sheets: [{ properties: { title: 'S', gridProperties: { rowCount } } }] } }),
      values: {
        get: async ({ range }: GetParams) => {
          ranges.push(range);
          if (range === 'S!1:1') return { data: { values: [['id', 'name']] } };
          if (opts.failKeyRead) throw opts.failKeyRead;
          const [, start, end] = /^S!A(\d+):A(\d+)$/.exec(range) as RegExpExecArray;
          // fixed slice of the prepared rows; trailing empties are simply absent
          const values = keyRows.slice(Number(start) - 2, Number(end) - 1);
          // Sheets omits trailing empty rows of the requested range
          if (opts.omitTrailingEmpty) while (values.length > 0 && values[values.length - 1].length === 0) values.pop();
          return { data: { values } };
        },
      },
    },
  };
  return { sheets: sheets as never, ranges };
}

describe('snapshotHeaderKeysAndPositions', () => {
  it('rejects naming a duplicate key and both its row numbers', async () => {
    const { sheets } = makeSheets([['a'], ['b'], ['a']]);
    await assert.rejects(snapshotHeaderKeysAndPositions(sheets, 'id1', 'S', ['id']), (err: Error) => {
      assert.ok(err.message.includes('"a" (rows 2, 4)'), err.message);
      return true;
    });
  });

  it('rejects a key-column read error and keeps the original as cause', async () => {
    const original = Object.assign(new Error('The caller does not have permission'), { code: 403 });
    const { sheets } = makeSheets([['a']], { failKeyRead: original });
    await assert.rejects(snapshotHeaderKeysAndPositions(sheets, 'id1', 'S', ['id']), (err: Error & { cause?: unknown }) => {
      assert.strictEqual(err.cause, original);
      return true;
    });
  });

  it('reads 1000 + n unique keys across chunks and ends without error', async () => {
    const keyRows = Array.from({ length: 1005 }, (_, i) => [`k${i}`]);
    const { sheets } = makeSheets(keyRows);
    const result = await snapshotHeaderKeysAndPositions(sheets, 'id1', 'S', ['id']);
    assert.strictEqual(result.keySet.size, 1005);
    assert.strictEqual(result.keyToRowMap.get('k1004'), 1006);
  });

  it('stops at the grid edge when the last chunk is exactly full', async () => {
    const keyRows = Array.from({ length: 1000 }, (_, i) => [`k${i}`]);
    const { sheets, ranges } = makeSheets(keyRows); // rowCount 1001
    const result = await snapshotHeaderKeysAndPositions(sheets, 'id1', 'S', ['id']);
    assert.strictEqual(result.keySet.size, 1000);
    assert.deepStrictEqual(
      ranges.filter((r) => r !== 'S!1:1'),
      ['S!A2:A1001']
    );
  });

  it('keeps scanning past blank rows at the end of a chunk and finds a duplicate below them', async () => {
    const keyRows: string[][] = Array.from({ length: 9 }, (_, i) => [`k${i}`]); // rows 2-10; k1 is on row 3
    keyRows[1] = ['dup'];
    keyRows[1498] = ['dup']; // row 1500, after blank rows 11..1500 are omitted from the first chunk
    for (let i = 9; i < 1498; i++) keyRows[i] = [];
    const { sheets } = makeSheets(keyRows, { rowCount: 2000, omitTrailingEmpty: true });
    await assert.rejects(snapshotHeaderKeysAndPositions(sheets, 'id1', 'S', ['id']), (err: Error) => {
      assert.ok(err.message.includes('"dup" (rows 3, 1500)'), err.message);
      return true;
    });
  });
});

describe('snapshotHeaderAndKeys', () => {
  it('rejects a key-column read error and keeps the original as cause', async () => {
    const original = Object.assign(new Error('The caller does not have permission'), { code: 403 });
    const { sheets } = makeSheets([['a']], { failKeyRead: original });
    await assert.rejects(snapshotHeaderAndKeys(sheets, 'id1', 'S', ['id']), (err: Error & { cause?: unknown }) => {
      assert.ok(err.message.startsWith('Failed to read key columns from S!A2:A2: '), err.message);
      assert.strictEqual(err.cause, original);
      return true;
    });
  });

  it('finds keys below a blank gap that spans a whole chunk', async () => {
    const keyRows: string[][] = Array.from({ length: 1999 }, () => []);
    keyRows[0] = ['top'];
    keyRows[1500] = ['below']; // row 1502, after blank rows 3..1001 omitted from the first chunk
    const { sheets } = makeSheets(keyRows, { rowCount: 2000, omitTrailingEmpty: true });
    const result = await snapshotHeaderAndKeys(sheets, 'id1', 'S', ['id']);
    assert.deepStrictEqual([...result.keySet].sort(), ['below', 'top']);
  });

  it('makes no extra read when the last chunk is exactly full at the grid edge', async () => {
    const keyRows = Array.from({ length: 1000 }, (_, i) => [`k${i}`]);
    const { sheets, ranges } = makeSheets(keyRows); // rowCount 1001
    const result = await snapshotHeaderAndKeys(sheets, 'id1', 'S', ['id']);
    assert.strictEqual(result.keySet.size, 1000);
    assert.deepStrictEqual(
      ranges.filter((r) => r !== 'S!1:1'),
      ['S!A2:A1001']
    );
  });

  it('reads 1000 + n unique keys across chunks', async () => {
    const keyRows = Array.from({ length: 1005 }, (_, i) => [`k${i}`]);
    const { sheets } = makeSheets(keyRows);
    const result = await snapshotHeaderAndKeys(sheets, 'id1', 'S', ['id']);
    assert.strictEqual(result.keySet.size, 1005);
    assert.ok(result.keySet.has('k1004'));
  });

  it('throws when the sheet row count is missing', async () => {
    const { sheets } = makeSheets([['a']]);
    (sheets as never as { spreadsheets: { get: unknown } }).spreadsheets.get = async () => ({ data: { sheets: [] } });
    await assert.rejects(snapshotHeaderAndKeys(sheets, 'id1', 'S', ['id']), /Could not read the row count of sheet "S"/);
  });
});
