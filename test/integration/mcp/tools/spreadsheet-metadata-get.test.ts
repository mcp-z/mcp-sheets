// Network: real Google Sheets, owned by the Live provider tests workflow and local configured runs.
import '../../../lib/env-loader.ts';
import type { sheets_v4 } from '@googleapis/sheets';
import { mcp } from '@mcp-z/mcp-sheets';
import { ProtocolError, ProtocolErrorCode } from '@mcp-z/server';
import assert from 'assert';
import type { Input } from '../../../../src/mcp/tools/spreadsheet-metadata-get.ts';
import { createExtra, type TypedHandler } from '../../../lib/create-extra.ts';
import { createMetadataFixture, METADATA_SHEET_ID, METADATA_SHEET_TITLE } from '../../../lib/metadata-fixture.ts';

describe('spreadsheet-metadata-get (Google integration)', () => {
  let fixture: Awaited<ReturnType<typeof createMetadataFixture>>;
  let handler: TypedHandler<Input>;
  const tool = mcp.toolFactories.spreadsheetMetadataGet();

  before(async () => {
    fixture = await createMetadataFixture();
    handler = fixture.middleware.withToolAuth(tool).handler;
  });
  after(async () => {
    if (fixture) await fixture.close();
  });

  it('reads both tabs, full properties and independently seeded structures without grid data', async () => {
    const response = await handler({ id: fixture.id }, createExtra());
    const result = tool.config.outputSchema.parse(response.structuredContent).result;
    assert.equal(result.type, 'success');
    if (result.type !== 'success') assert.fail('Expected authenticated success');
    const metadata = result.metadata as sheets_v4.Schema$Spreadsheet;
    assert.equal(metadata.spreadsheetId, fixture.id);
    assert.equal(metadata.properties?.locale, 'en_CA');
    assert.equal(metadata.properties?.timeZone, 'America/Vancouver');
    assert.equal(metadata.properties?.autoRecalc, 'ON_CHANGE');
    assert.equal(metadata.sheets?.length, 2);
    assert.equal(metadata.sheets?.find((sheet) => sheet.properties?.sheetId === 0)?.properties?.hidden, true);
    const sheet = metadata.sheets?.find((sheet) => sheet.properties?.sheetId === METADATA_SHEET_ID);
    assert.equal(sheet?.properties?.title, METADATA_SHEET_TITLE);
    assert.equal(sheet?.properties?.gridProperties?.frozenRowCount, 2);
    assert.equal(sheet?.properties?.gridProperties?.frozenColumnCount, 1);
    assert.equal(sheet?.charts?.[0]?.spec?.title, 'Metadata chart');
    assert.equal(sheet?.protectedRanges?.[0]?.description, 'Metadata protection');
    assert.equal(sheet?.conditionalFormats?.[0]?.booleanRule?.condition?.type, 'NUMBER_GREATER');
    assert.equal(metadata.namedRanges?.[0]?.name, 'MetadataSource');
    for (const tab of metadata.sheets ?? []) assert.equal(tab.data, undefined);
    const text = response.content.find((content) => content.type === 'text');
    assert.ok(text && text.type === 'text');
    assert.deepEqual(JSON.parse(text.text), result);
    console.info('Workbook metadata fixture response bytes:', Buffer.byteLength(text.text));
  });

  it('preserves omitted collections on an empty tab', async () => {
    const response = await handler({ id: fixture.id }, createExtra());
    const result = tool.config.outputSchema.parse(response.structuredContent).result;
    if (result.type !== 'success') assert.fail('Expected authenticated success');
    const metadata = result.metadata as sheets_v4.Schema$Spreadsheet;
    const empty = metadata.sheets?.find((sheet) => sheet.properties?.sheetId === 0);
    assert.ok(empty);
    assert.equal(empty.charts, undefined);
    assert.equal(empty.protectedRanges, undefined);
    assert.equal(empty.conditionalFormats, undefined);
  });

  it('reports provider errors for a missing workbook', async () => {
    await assert.rejects(
      () => handler({ id: 'nonexistent-metadata-workbook' }, createExtra()),
      (error: unknown) => error instanceof ProtocolError && error.code === ProtocolErrorCode.InternalError
    );
  });
});
