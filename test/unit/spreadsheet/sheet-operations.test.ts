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

describe('ensureTabAndHeaders header read errors', () => {
  /** Records every write; reading S!1:1 fails with `readError`, and `tabs` are the tab titles the spreadsheet has. */
  function makeWriteRecorder(readError: unknown, tabs: string[]) {
    const writes: string[] = [];
    const sheets = {
      spreadsheets: {
        get: async () => ({ data: { sheets: tabs.map((title, i) => ({ properties: { title, sheetId: i } })) } }),
        batchUpdate: async () => {
          writes.push('batchUpdate');
          return { data: {} };
        },
        values: {
          get: async () => {
            throw readError;
          },
          append: async () => {
            writes.push('append');
            return { data: {} };
          },
          update: async () => {
            writes.push('update');
            return { data: {} };
          },
        },
      },
    };
    return { sheets: sheets as never, writes };
  }

  it('rejects with the read error and writes nothing when the tab exists', async () => {
    const unavailable = Object.assign(new Error('The service is currently unavailable'), { code: 503 });
    const { sheets, writes } = makeWriteRecorder(unavailable, ['S']);
    await assert.rejects(call(sheets), (err: Error) => err === unavailable);
    assert.deepStrictEqual(writes, []);
  });

  it('rejects with the read error when the sheetRef names an existing tab', async () => {
    const unavailable = Object.assign(new Error('The service is currently unavailable'), { code: 503 });
    const { sheets, writes } = makeWriteRecorder(unavailable, ['S']);
    await assert.rejects(ensureTabAndHeaders(sheets, { spreadsheetId: 'id1', sheetRef: 'S', logger }), (err: Error) => err === unavailable);
    assert.deepStrictEqual(writes, []);
  });

  it('creates the tab and writes the header when the sheetRef names no tab', async () => {
    const calls: string[] = [];
    const sheets = {
      spreadsheets: {
        get: async () => ({ data: { sheets: [{ properties: { title: 'Other', sheetId: 0 } }] } }),
        batchUpdate: async () => {
          calls.push('addSheet');
          return { data: {} };
        },
        values: {
          get: async () => ({ data: {} }),
          append: async () => {
            calls.push('append');
            return { data: {} };
          },
        },
      },
    };
    const { header } = await ensureTabAndHeaders(sheets as never, { spreadsheetId: 'id1', sheetRef: 'New', requiredHeader: ['id', 'name'], logger });
    assert.deepStrictEqual(header, ['id', 'name']);
    assert.deepStrictEqual(calls, ['addSheet', 'append']);
  });

  it('rejects with the addSheet error when creating the tab fails', async () => {
    const denied = Object.assign(new Error('The caller does not have permission'), { code: 403 });
    const sheets = {
      spreadsheets: {
        get: async () => ({ data: { sheets: [] } }),
        batchUpdate: async () => {
          throw denied;
        },
        values: { get: async () => ({ data: {} }), append: async () => ({ data: {} }) },
      },
    };
    await assert.rejects(ensureTabAndHeaders(sheets as never, { spreadsheetId: 'id1', sheetRef: 'New', logger }), (err: Error) => err === denied);
  });
});
