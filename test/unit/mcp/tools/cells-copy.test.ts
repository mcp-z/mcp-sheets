import { ProtocolError, ProtocolErrorCode } from '@mcp-z/server';
import assert from 'assert';
import { toolFactories } from '../../../../src/mcp/index.ts';
import { type CellsCopySheets, copyCells, type Input } from '../../../../src/mcp/tools/cells-copy.ts';
import { calculatePasteRange, cellBlockToA1 } from '../../../../src/spreadsheet/range-operations.ts';

type BatchParams = Parameters<CellsCopySheets['spreadsheets']['batchUpdate']>[0];

function createFake(grids: Record<number, { rows: number; columns: number; title: string }> = { 0: { rows: 100, columns: 26, title: 'Main' }, 7: { rows: 20, columns: 5, title: 'Small' } }) {
  const gets: unknown[] = [];
  const updates: BatchParams[] = [];
  const sheets: CellsCopySheets = {
    spreadsheets: {
      async get(params) {
        gets.push(params);
        return {
          data: {
            properties: { title: 'Book' },
            spreadsheetUrl: 'https://docs.google.com/spreadsheets/d/SID/edit',
            sheets: Object.entries(grids).map(([sheetId, g]) => ({ properties: { sheetId: Number(sheetId), title: g.title, gridProperties: { rowCount: g.rows, columnCount: g.columns } } })),
          },
        };
      },
      async batchUpdate(params) {
        updates.push(params);
        return {};
      },
    },
  };
  return { sheets, gets, updates };
}

const base: Input = { id: 'SID', sourceGid: '0', sourceRange: 'A1:C10', destinationGid: '0', destinationRange: 'E1', pasteType: 'PASTE_NORMAL', pasteOrientation: 'NORMAL' };

async function rejects(input: Partial<Input>, grids?: Parameters<typeof createFake>[0]) {
  const fake = createFake(grids);
  await assert.rejects(
    () => copyCells(fake.sheets, { ...base, ...input }),
    (error: unknown) => error instanceof ProtocolError && error.code === ProtocolErrorCode.InvalidParams
  );
  assert.equal(fake.updates.length, 0, 'no write on rejection');
  return fake;
}

describe('cells-copy tool', () => {
  const tool = toolFactories.cellsCopy();

  it('registers with strict gids and default paste modes', () => {
    assert.equal(tool.name, 'cells-copy');
    const parsed = tool.config.inputSchema.parse({ id: 'SID', sourceGid: 0, sourceRange: 'A1', destinationGid: '5', destinationRange: 'B2' });
    assert.equal(parsed.sourceGid, '0');
    assert.equal(parsed.destinationGid, '5');
    assert.equal(parsed.pasteType, 'PASTE_NORMAL');
    assert.equal(parsed.pasteOrientation, 'NORMAL');
    for (const gid of [null, undefined, true, -1, 1.5, 'abc', '']) {
      assert.equal(tool.config.inputSchema.safeParse({ ...base, sourceGid: gid }).success, false);
      assert.equal(tool.config.inputSchema.safeParse({ ...base, destinationGid: gid }).success, false);
    }
  });

  it('converts to the JSON Schema that tools/list advertises, accepting integer and digit-string gids', () => {
    // Same call as @modelcontextprotocol/server standardSchemaToJsonSchema(schema, 'input') used by tools/list
    // (node_modules/@modelcontextprotocol/server/dist/mcp-DIH4cS6P.mjs:1685 -> src-Cqbh3MYc.mjs:5377-5381)
    const json = tool.config.inputSchema['~standard'].jsonSchema.input({ target: 'draft-2020-12' }) as { required?: string[]; properties: Record<string, { anyOf?: { type?: string }[] }> };
    assert.ok(json.required?.includes('sourceGid') && json.required.includes('destinationGid'));
    for (const key of ['sourceGid', 'destinationGid']) {
      const types = json.properties[key]?.anyOf?.map((branch) => branch.type).sort();
      assert.deepEqual(types, ['integer', 'string']);
    }
  });

  it('rejects gids above int32', () => {
    for (const gid of [2147483648, '2147483648']) {
      assert.equal(tool.config.inputSchema.safeParse({ ...base, sourceGid: gid }).success, false);
    }
  });

  it('fails instead of skipping the grid check when the grid size is missing', async () => {
    for (const gridProperties of [undefined, { rowCount: 100 }, { columnCount: 26 }]) {
      const fake = createFake();
      const get = fake.sheets.spreadsheets.get;
      fake.sheets.spreadsheets.get = async (params) => {
        const r = await get(params);
        for (const sheet of r.data.sheets ?? []) sheet.properties = { ...sheet.properties, gridProperties };
        return r;
      };
      await assert.rejects(() => copyCells(fake.sheets, base), /Could not read the grid size of sheet "Main"/);
      assert.equal(fake.updates.length, 0);
    }
  });

  describe('range validation', () => {
    it('anchors a single destination cell at the source size', async () => {
      const { sheets } = createFake();
      assert.equal((await copyCells(sheets, base)).pastedRange, 'E1:G10');
    });

    it('accepts an exact multiple destination and tiles it', async () => {
      const { sheets } = createFake();
      assert.equal((await copyCells(sheets, { ...base, destinationRange: 'E1:J20' })).pastedRange, 'E1:J20');
    });

    it('rejects a non-multiple destination', async () => {
      await rejects({ destinationRange: 'E1:H10' });
      await rejects({ destinationRange: 'E1:G15' });
    });

    it('rejects a paste that exceeds the destination grid', async () => {
      await rejects({ destinationGid: '7', destinationRange: 'A15' }); // rows 15-24 > 20
      await rejects({ destinationGid: '7', destinationRange: 'D1' }); // columns D-F > 5
    });

    it('rejects a source outside its grid', async () => {
      await rejects({ sourceGid: '7', sourceRange: 'A1:C25' });
    });

    it('rejects infinite, reversed, prefixed and invalid ranges before any network call', async () => {
      for (const sourceRange of ['A:B', '1:2', 'C10:A1', 'Main!A1:B2', 'not a range']) {
        const fake = await rejects({ sourceRange });
        assert.equal(fake.gets.length, 0, `${sourceRange} fails before any read`);
      }
      for (const destinationRange of ['A:B', '1:2', 'E10:E1', 'Main!E1']) {
        const fake = await rejects({ destinationRange });
        assert.equal(fake.gets.length, 0, `${destinationRange} fails before any read`);
      }
    });

    it('rejects unknown sheet ids', async () => {
      await rejects({ sourceGid: '999' });
      await rejects({ destinationGid: '999' });
    });

    it('computes the TRANSPOSE footprint', async () => {
      assert.equal(cellBlockToA1(calculatePasteRange('A1:C10', 'E1', 'TRANSPOSE')), 'E1:N3');
      const { sheets } = createFake();
      assert.equal((await copyCells(sheets, { ...base, pasteOrientation: 'TRANSPOSE' })).pastedRange, 'E1:N3');
      // 10x3 transposed is 3x10: a 6x20 destination is a multiple, a 10x3 one is not
      assert.equal(cellBlockToA1(calculatePasteRange('A1:C10', 'E1:X6', 'TRANSPOSE')), 'E1:X6');
      assert.throws(() => calculatePasteRange('A1:C10', 'E1:G10', 'TRANSPOSE'));
    });

    it('checks the grid with the transposed footprint', async () => {
      await rejects({ destinationGid: '7', destinationRange: 'A1', pasteOrientation: 'TRANSPOSE' }); // 3x10 needs 10 columns, grid has 5
    });
  });

  describe('CopyPasteRequest', () => {
    it('sends PASTE_NORMAL / NORMAL by default with zero-based half-open GridRanges', async () => {
      const { sheets, updates } = createFake();
      const parsed = tool.config.inputSchema.parse({ id: 'SID', sourceGid: 0, sourceRange: 'A1:C10', destinationGid: 7, destinationRange: 'B2' });
      const result = await copyCells(sheets, parsed);
      assert.equal(updates.length, 1);
      assert.equal(updates[0]?.spreadsheetId, 'SID');
      assert.deepEqual(updates[0]?.requestBody.requests, [
        {
          copyPaste: {
            source: { sheetId: 0, startRowIndex: 0, endRowIndex: 10, startColumnIndex: 0, endColumnIndex: 3 },
            destination: { sheetId: 7, startRowIndex: 1, endRowIndex: 2, startColumnIndex: 1, endColumnIndex: 2 },
            pasteType: 'PASTE_NORMAL',
            pasteOrientation: 'NORMAL',
          },
        },
      ]);
      assert.equal(result.pastedRange, 'B2:D11');
      assert.equal(result.destinationSheetTitle, 'Small');
      assert.equal(result.sheetUrl, 'https://docs.google.com/spreadsheets/d/SID/edit#gid=7');
    });

    it('sends each of the 7 paste types explicitly', async () => {
      for (const pasteType of ['PASTE_NORMAL', 'PASTE_VALUES', 'PASTE_FORMAT', 'PASTE_NO_BORDERS', 'PASTE_FORMULA', 'PASTE_DATA_VALIDATION', 'PASTE_CONDITIONAL_FORMATTING'] as const) {
        const { sheets, updates } = createFake();
        const result = await copyCells(sheets, { ...base, pasteType });
        assert.equal(updates[0]?.requestBody.requests?.[0]?.copyPaste?.pasteType, pasteType);
        assert.equal(result.pasteType, pasteType);
      }
      assert.equal(tool.config.inputSchema.safeParse({ ...base, pasteType: 'PASTE_ALL' }).success, false);
    });

    it('sends TRANSPOSE orientation', async () => {
      const { sheets, updates } = createFake();
      await copyCells(sheets, { ...base, pasteOrientation: 'TRANSPOSE' });
      assert.equal(updates[0]?.requestBody.requests?.[0]?.copyPaste?.pasteOrientation, 'TRANSPOSE');
    });
  });
});
