import '../../lib/env-loader.ts';
import assert from 'assert';
import { ensureTabAndHeaders } from '../../../src/spreadsheet/sheet-operations.ts';

const logger = {} as never;

/** Plain fake: header is ['id', 'name']; `keyRows` are data rows from row 2, `rowCount` the grid size. */
function makeSheets(keyRows: string[][], opts: { rowCount?: number; failKeyRead?: unknown } = {}) {
  const rowCount = opts.rowCount ?? keyRows.length + 1;
  const ranges: string[] = [];
  const sheets = {
    spreadsheets: {
      get: async () => ({ data: { sheets: [{ properties: { title: 'S', gridProperties: { rowCount } } }] } }),
      values: {
        get: async ({ range }: { range: string }) => {
          ranges.push(range);
          if (range === 'S!1:1') return { data: { values: [['id', 'name']] } };
          if (opts.failKeyRead) throw opts.failKeyRead;
          const [, start, end] = /^S!A(\d+):A(\d+)$/.exec(range) as RegExpExecArray;
          const values = keyRows.slice(Number(start) - 2, Number(end) - 1);
          // Sheets omits trailing empty rows of the requested range
          while (values.length > 0 && values[values.length - 1].length === 0) values.pop();
          return { data: { values } };
        },
      },
    },
  };
  return { sheets: sheets as never, ranges };
}

const call = (sheets: never) => ensureTabAndHeaders(sheets, { spreadsheetId: 'id1', sheetTitle: 'S', requiredHeader: ['id', 'name'], keyColumns: ['id'], logger });

describe('ensureTabAndHeaders key scan', () => {
  it('finds a key below a full blank chunk', async () => {
    const keyRows: string[][] = Array.from({ length: 2000 }, () => []);
    keyRows[0] = ['top'];
    keyRows[1500] = ['below']; // row 1502
    const { sheets } = makeSheets(keyRows, { rowCount: 2001 });
    const { keySet } = await call(sheets);
    assert.deepStrictEqual([...(keySet as Set<string>)].sort(), ['below', 'top']);
  });

  it('reports a key read error and keeps the original as cause', async () => {
    const original = Object.assign(new Error('The caller does not have permission'), { code: 403 });
    const { sheets } = makeSheets([['a']], { failKeyRead: original });
    await assert.rejects(call(sheets), (err: Error & { cause?: unknown }) => {
      assert.ok(err.message.startsWith('Failed to read key columns from S!A2:A2: '), err.message);
      assert.strictEqual(err.cause, original);
      return true;
    });
  });

  it('reads nothing past the grid row count', async () => {
    const keyRows = Array.from({ length: 1000 }, (_, i) => [`k${i}`]);
    const { sheets, ranges } = makeSheets(keyRows); // rowCount 1001
    const { keySet } = await call(sheets);
    assert.strictEqual((keySet as Set<string>).size, 1000);
    assert.deepStrictEqual(
      ranges.filter((r) => r !== 'S!1:1'),
      ['S!A2:A1001']
    );
  });
});
