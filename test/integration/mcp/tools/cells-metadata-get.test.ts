// Network: real Google Sheets, owned by the Live provider tests workflow and local configured runs.
import '../../../lib/env-loader.ts';
import type { sheets_v4 } from '@googleapis/sheets';
import { mcp } from '@mcp-z/mcp-sheets';
import { ProtocolError, ProtocolErrorCode } from '@mcp-z/server';
import assert from 'assert';
import type { Input } from '../../../../src/mcp/tools/cells-metadata-get.ts';
import { createExtra, type TypedHandler } from '../../../lib/create-extra.ts';
import { createMetadataFixture, METADATA_SHEET_ID, METADATA_SHEET_TITLE } from '../../../lib/metadata-fixture.ts';

function selectedCells(metadata: sheets_v4.Schema$Spreadsheet) {
  const cells = new Map<string, sheets_v4.Schema$CellData>();
  for (const sheet of metadata.sheets ?? []) {
    for (const data of sheet.data ?? []) {
      for (const [row, rowData] of (data.rowData ?? []).entries()) {
        for (const [column, cell] of (rowData.values ?? []).entries()) {
          cells.set(`${(data.startRow ?? 0) + row},${(data.startColumn ?? 0) + column}`, cell);
        }
      }
    }
  }
  return cells;
}

describe('cells-metadata-get (Google integration)', () => {
  let fixture: Awaited<ReturnType<typeof createMetadataFixture>>;
  let handler: TypedHandler<Input>;
  const tool = mcp.toolFactories.cellsMetadataGet();

  before(async () => {
    fixture = await createMetadataFixture();
    handler = fixture.middleware.withToolAuth(tool).handler;
  });
  after(async () => {
    if (fixture) await fixture.close();
  });

  it('resolves a quoted title by gid and preserves sparse rows and columns in batched reads', async () => {
    const response = await handler({ id: fixture.id, gid: String(METADATA_SHEET_ID), ranges: ['C6:F9', 'J20'] }, createExtra());
    const result = tool.config.outputSchema.parse(response.structuredContent).result;
    if (result.type !== 'success') assert.fail('Expected authenticated success');
    assert.deepEqual(result.requestedRanges, ["'Investor''s ! Notes'!C6:F9", "'Investor''s ! Notes'!J20"]);
    const metadata = result.metadata as sheets_v4.Schema$Spreadsheet;
    assert.equal(metadata.sheets?.length, 1);
    assert.equal(metadata.sheets?.[0]?.properties?.sheetId, METADATA_SHEET_ID);
    assert.equal(metadata.sheets?.[0]?.properties?.title, METADATA_SHEET_TITLE);
    const cells = selectedCells(metadata);
    assert.equal(cells.get('7,4')?.note, 'Selected note');
    assert.deepEqual(cells.get('6,3')?.dataValidation, fixture.validation);
    assert.equal(cells.get('19,9')?.note, 'Outside note');
    assert.equal(metadata.sheets?.[0]?.data?.[0]?.startRow, 5);
    assert.equal(metadata.sheets?.[0]?.data?.[0]?.startColumn, 2);
    for (const cell of cells.values()) {
      assert.equal(cell.userEnteredValue, undefined);
      assert.equal(cell.effectiveValue, undefined);
      assert.equal(cell.formattedValue, undefined);
      assert.equal(cell.userEnteredFormat, undefined);
      assert.equal(cell.effectiveFormat, undefined);
    }
    assert.equal(metadata.namedRanges, undefined);
    assert.equal(metadata.sheets?.[0]?.charts, undefined);
    const text = response.content.find((content) => content.type === 'text');
    assert.ok(text && text.type === 'text');
    assert.deepEqual(JSON.parse(text.text), result);
    console.info('Cell metadata fixture response bytes:', Buffer.byteLength(text.text));
  });

  it('limits notes and validation to the selected cells', async () => {
    const response = await handler({ id: fixture.id, gid: String(METADATA_SHEET_ID), ranges: ['E8'] }, createExtra());
    const result = tool.config.outputSchema.parse(response.structuredContent).result;
    if (result.type !== 'success') assert.fail('Expected authenticated success');
    const cells = selectedCells(result.metadata as sheets_v4.Schema$Spreadsheet);
    assert.deepEqual([...cells.entries()], [['7,4', { note: 'Selected note' }]]);
  });

  it('reads a regular pivot definition at its anchor cell', async () => {
    const response = await handler({ id: fixture.id, gid: String(METADATA_SHEET_ID), ranges: ['M2'] }, createExtra());
    const result = tool.config.outputSchema.parse(response.structuredContent).result;
    if (result.type !== 'success') assert.fail('Expected authenticated success');
    const pivot = selectedCells(result.metadata as sheets_v4.Schema$Spreadsheet).get('1,12')?.pivotTable;
    assert.equal(pivot?.source?.sheetId, METADATA_SHEET_ID);
    assert.equal(pivot?.values?.[0]?.summarizeFunction, 'SUM');
  });

  it('returns no fabricated metadata for an empty range', async () => {
    const response = await handler({ id: fixture.id, gid: String(METADATA_SHEET_ID), ranges: ['Z90:Z91'] }, createExtra());
    const result = tool.config.outputSchema.parse(response.structuredContent).result;
    if (result.type !== 'success') assert.fail('Expected authenticated success');
    for (const cell of selectedCells(result.metadata as sheets_v4.Schema$Spreadsheet).values()) assert.deepEqual(cell, {});
  });

  it('preserves InvalidParams for missing gids and invalid ranges and wraps missing workbook errors', async () => {
    const invalidParams = (error: unknown) => error instanceof ProtocolError && error.code === ProtocolErrorCode.InvalidParams;
    await assert.rejects(() => handler({ id: fixture.id, gid: '999999', ranges: ['A1'] }, createExtra()), invalidParams);
    await assert.rejects(() => handler({ id: fixture.id, gid: String(METADATA_SHEET_ID), ranges: ['A:A'] }, createExtra()), invalidParams);
    await assert.rejects(
      () => handler({ id: 'nonexistent-metadata-workbook', gid: '0', ranges: ['A1'] }, createExtra()),
      (error: unknown) => error instanceof ProtocolError && error.code === ProtocolErrorCode.InternalError
    );
  });

  it('rejects ranges on a real OBJECT sheet', async () => {
    const response = await fixture.sheets.spreadsheets.batchUpdate({
      spreadsheetId: fixture.id,
      requestBody: {
        requests: [
          {
            addChart: {
              chart: {
                spec: {
                  title: 'Object chart',
                  basicChart: {
                    chartType: 'COLUMN',
                    domains: [{ domain: { sourceRange: { sources: [{ sheetId: METADATA_SHEET_ID, startRowIndex: 0, endRowIndex: 4, startColumnIndex: 0, endColumnIndex: 1 }] } } }],
                    series: [{ series: { sourceRange: { sources: [{ sheetId: METADATA_SHEET_ID, startRowIndex: 0, endRowIndex: 4, startColumnIndex: 1, endColumnIndex: 2 }] } } }],
                  },
                },
                position: { newSheet: true },
              },
            },
          },
        ],
      },
    });
    const gid = response.data.replies?.[0]?.addChart?.chart?.position?.sheetId;
    assert.equal(typeof gid, 'number');
    await assert.rejects(
      () => handler({ id: fixture.id, gid: String(gid), ranges: ['A1'] }, createExtra()),
      (error: unknown) => error instanceof ProtocolError && error.code === ProtocolErrorCode.InvalidParams && /OBJECT/.test(error.message)
    );
  });
});
