import '../../lib/env-loader.ts';
import assert from 'assert';
import { performBatchUpdates, snapshotHeaderAndKeys, snapshotHeaderKeysAndPositions, upsertByKey } from '../../../src/spreadsheet/data-operations.ts';

type GetParams = { range: string };

/** Plain parameter object returning fixed data: `keyRows` are the sheet's data rows (row 2 onward), `rowCount` the grid size. */
function makeSheets(keyRows: string[][], opts: { rowCount?: number; failKeyRead?: unknown; failHeaderRead?: unknown; omitTrailingEmpty?: boolean } = {}) {
  const rowCount = opts.rowCount ?? keyRows.length + 1;
  const ranges: string[] = [];
  const appends: unknown[] = [];
  const sheets = {
    spreadsheets: {
      get: async () => ({ data: { sheets: [{ properties: { title: 'S', gridProperties: { rowCount } } }] } }),
      values: {
        append: async (params: unknown) => {
          appends.push(params);
          return { data: {} };
        },
        get: async ({ range }: GetParams) => {
          ranges.push(range);
          if (range === 'S!1:1' && opts.failHeaderRead) throw opts.failHeaderRead;
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
  return { sheets: sheets as never, ranges, appends };
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

describe('header read errors', () => {
  const quota = Object.assign(new Error('Quota exceeded'), { code: 429 });

  it('snapshotHeaderKeysAndPositions rejects instead of returning an empty key set', async () => {
    const { sheets } = makeSheets([['a']], { failHeaderRead: quota });
    await assert.rejects(snapshotHeaderKeysAndPositions(sheets, 'id1', 'S', ['id']), (err: Error) => err === quota);
  });

  it('snapshotHeaderAndKeys rejects instead of returning an empty key set', async () => {
    const { sheets } = makeSheets([['a']], { failHeaderRead: quota });
    await assert.rejects(snapshotHeaderAndKeys(sheets, 'id1', 'S', ['id']), (err: Error) => err === quota);
  });

  it('upsertByKey rejects without appending a header row', async () => {
    const { sheets, appends } = makeSheets([['a']], { failHeaderRead: quota });
    await assert.rejects(
      upsertByKey(sheets, {
        spreadsheetId: 'id1',
        sheetTitle: 'S',
        rows: [['b', 'B']],
        canonicalHeaders: ['id', 'name'],
        options: { keyStrategy: { keyColumns: ['id'], useProviderIdLogic: false }, allowUpdates: true },
        logger: {} as never,
      }),
      (err: Error) => err === quota
    );
    assert.strictEqual(appends.length, 0);
  });
});

/** Recording-only client: keeps every write request as sent and returns an empty response; `get` lists a gid 0 tab "First" and the target tab "Target" (gid 123). */
function makeRecordingSheets() {
  const writes: Array<{ api: string; body: unknown }> = [];
  const record = (api: string) => async (params: unknown) => {
    writes.push({ api, body: params });
    return { data: {} };
  };
  const sheets = {
    spreadsheets: {
      get: async () => ({ data: { sheets: [{ properties: { sheetId: 0, title: 'First' } }, { properties: { sheetId: 123, title: 'Target' } }] } }),
      batchUpdate: record('spreadsheets.batchUpdate'),
      values: { batchUpdate: record('values.batchUpdate'), update: record('values.update') },
    },
  };
  return { sheets: sheets as never, writes };
}

/** Tab each sent write addresses: the sheetId of an updateCells range, or the tab title of an A1 range. */
function addressedTabs(writes: Array<{ api: string; body: unknown }>): Array<number | string> {
  const tabs: Array<number | string> = [];
  for (const { api, body } of writes) {
    const params = body as { requestBody?: { requests?: Array<{ updateCells?: { range: { sheetId: number } } }>; data?: Array<{ range: string }> }; range?: string };
    if (api === 'spreadsheets.batchUpdate') for (const r of params.requestBody?.requests ?? []) tabs.push(r.updateCells?.range.sheetId ?? -1);
    else if (api === 'values.batchUpdate') for (const d of params.requestBody?.data ?? []) tabs.push(/^'?(.+?)'?!/.exec(d.range)?.[1] ?? '');
    else tabs.push(/^'?(.+?)'?!/.exec(params.range ?? '')?.[1] ?? '');
  }
  return tabs;
}

describe('performBatchUpdates', () => {
  it('sends every write to the target tab and none to the gid 0 tab', async () => {
    const { sheets, writes } = makeRecordingSheets();
    await performBatchUpdates(sheets, 'id1', 'Target', [{ row: ['a', 'new'], existingRowIndex: 2 }], ['id', 'name'], 50, 'USER_ENTERED');
    const tabs = addressedTabs(writes);
    assert.ok(tabs.length > 0, 'a write was sent');
    assert.ok(
      tabs.every((t) => t === 'Target' || t === 123),
      `writes addressed ${JSON.stringify(tabs)}`
    );
  });

  it('passes valueInputOption to the batch write', async () => {
    const { sheets, writes } = makeRecordingSheets();
    await performBatchUpdates(sheets, 'id1', 'Target', [{ row: ['a', 'new'], existingRowIndex: 2 }], ['id', 'name'], 50, 'RAW');
    const options = writes.map((w) => (w.body as { requestBody?: { valueInputOption?: string } }).requestBody?.valueInputOption);
    assert.deepStrictEqual(options, ['RAW']);
  });
});

