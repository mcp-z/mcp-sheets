import assert from 'assert';
import { toolFactories } from '../../../../src/mcp/index.ts';

describe('spreadsheet-metadata-get tool', () => {
  const tool = toolFactories.spreadsheetMetadataGet();

  it('registers an id-only read operation', () => {
    assert.equal(tool.name, 'spreadsheet-metadata-get');
    assert.equal(tool.config.annotations.readOnlyHint, true);
    assert.deepEqual(tool.config.inputSchema.parse({ id: 'spreadsheet-id' }), { id: 'spreadsheet-id' });
    assert.equal(tool.config.inputSchema.safeParse({ id: 'spreadsheet-id', ranges: ['A1'] }).success, false);
    assert.equal(tool.config.inputSchema.safeParse({ id: '' }).success, false);
  });

  it('validates a spreadsheet envelope while retaining nested provider definitions', () => {
    const metadata = {
      spreadsheetId: 'spreadsheet-id',
      properties: { title: 'Example', locale: 'en_CA', timeZone: 'America/Vancouver' },
      namedRanges: [{ name: 'Example', range: { sheetId: 7, startRowIndex: 0 } }],
      sheets: [{ properties: { sheetId: 7, title: 'Candidate', hidden: true, gridProperties: { frozenRowCount: 2 } }, charts: [{ chartId: 12, spec: { title: 'Example chart' } }] }],
    };
    const result = { type: 'success', id: 'spreadsheet-id', metadata, fieldMask: 'provider-selection' };
    assert.deepEqual(tool.config.outputSchema.parse({ result }), { result });
    for (const invalidMetadata of [null, [], 'metadata', {}, { spreadsheetId: 'id', properties: [] }, { spreadsheetId: 'id', sheets: [false] }]) {
      assert.equal(tool.config.outputSchema.safeParse({ result: { ...result, metadata: invalidMetadata } }).success, false);
    }
  });
});
