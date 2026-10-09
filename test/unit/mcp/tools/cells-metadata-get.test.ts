import { ProtocolError, ProtocolErrorCode } from '@mcp-z/server';
import assert from 'assert';
import { z } from 'zod';
import { toolFactories } from '../../../../src/mcp/index.ts';
import { createExtra } from '../../../lib/create-extra.ts';

describe('cells-metadata-get tool', () => {
  const tool = toolFactories.cellsMetadataGet();
  const input = { id: 'spreadsheet-id', gid: '0', ranges: ['A1'] };

  it('registers a read operation with required sheet-local ranges and coerces gid zero', () => {
    assert.equal(tool.name, 'cells-metadata-get');
    assert.equal(tool.config.annotations.readOnlyHint, true);
    assert.deepEqual(tool.config.inputSchema.parse({ ...input, gid: 0, ranges: [' D9 '] }), { ...input, ranges: ['D9'] });
    assert.equal(tool.config.inputSchema.safeParse({ id: input.id, gid: input.gid }).success, false);
    assert.equal(tool.config.inputSchema.safeParse({ id: input.id, ranges: input.ranges }).success, false);
    for (const gid of [undefined, null, {}, [], true, '', NaN, Infinity]) {
      assert.equal(tool.config.inputSchema.safeParse({ ...input, gid }).success, false);
    }
  });

  it('advertises gid as a required string-or-number input in the MCP JSON schema', () => {
    const schema = z.toJSONSchema(tool.config.inputSchema, { io: 'input' });
    assert.deepEqual(schema.required, ['id', 'gid', 'ranges']);
    const gidSchema = schema.properties?.gid;
    assert.ok(gidSchema && typeof gidSchema === 'object');
    assert.deepEqual(
      gidSchema.anyOf?.map((branch) => (typeof branch === 'object' ? branch.type : undefined)),
      ['string', 'number']
    );
    assert.equal(gidSchema.anyOf?.[0]?.minLength, 1);
  });

  it('accepts exactly 50 ranges and rejects 51', () => {
    assert.equal(tool.config.inputSchema.safeParse({ ...input, ranges: Array(50).fill('A1') }).success, true);
    assert.equal(tool.config.inputSchema.safeParse({ ...input, ranges: Array(51).fill('A1') }).success, false);
  });

  it('accepts exactly 250000 aggregate cells and rejects 250001 across several ranges', () => {
    const ranges = ['A1:J10000', 'K1:T10000', 'U1:Y10000'];
    assert.equal(tool.config.inputSchema.safeParse({ ...input, ranges }).success, true);
    assert.equal(tool.config.inputSchema.safeParse({ ...input, ranges: [...ranges, 'Z1'] }).success, false);
    assert.equal(tool.config.inputSchema.safeParse({ ...input, ranges: ['A1:A125000', 'A1:A125000'] }).success, true);
    assert.equal(tool.config.inputSchema.safeParse({ ...input, ranges: ['A1:A125001', 'A1:A125000'] }).success, false);
  });

  it('rejects empty, malformed, reversed, unbounded and sheet-qualified selections', () => {
    for (const ranges of [[], [''], ['  '], ['A0'], ['A1:'], ['A:A'], ['1:2'], ['A5:B2'], ['C1:A5'], ['C5:A2'], ['Sheet1!A1'], ["'Sheet 1'!A1"], ['a1'], ['A10000001']]) {
      assert.equal(tool.config.inputSchema.safeParse({ ...input, ranges }).success, false, JSON.stringify(ranges));
    }
  });

  it('returns InvalidParams before authentication when a direct call has invalid ranges', async () => {
    await assert.rejects(
      () => tool.handler({ ...input, ranges: ['A:A'] }, createExtra()),
      (error: unknown) => error instanceof ProtocolError && error.code === ProtocolErrorCode.InvalidParams
    );
  });

  it('preserves sparse offsets, empty positions and selected nested definitions', () => {
    const result = {
      type: 'success',
      id: input.id,
      gid: input.gid,
      fieldMask: 'provider-selection',
      requestedRanges: ["'Sheet 1'!D7:F9"],
      metadata: {
        spreadsheetId: input.id,
        sheets: [
          {
            properties: { sheetId: 0, title: 'Sheet 1' },
            data: [{ startRow: 6, startColumn: 3, rowData: [{}, { values: [{}, { note: 'Note', dataValidation: { condition: { type: 'ONE_OF_LIST', values: [{ userEnteredValue: 'Yes' }] } } }] }] }, { rowData: [{ values: [{ pivotTable: { source: { sheetId: 0 } }, dataSourceTable: { dataSourceId: 'source-id' } }] }] }],
          },
        ],
      },
    };
    assert.deepEqual(tool.config.outputSchema.parse({ result }), { result });
    for (const metadata of [null, [], {}, { spreadsheetId: input.id, sheets: [{ data: [{ startRow: -1 }] }] }, { spreadsheetId: input.id, sheets: [{ data: [{ rowData: [{ values: [{ note: 3 }] }] }] }] }]) {
      assert.equal(tool.config.outputSchema.safeParse({ result: { ...result, metadata } }).success, false);
    }
  });
});
