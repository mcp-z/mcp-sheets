// Network: built-package metadata invocations use Google Sheets with configured credentials.
import '../lib/env-loader.ts';
import type { ManagedClient } from '@mcp-z/client';
import { mcp } from '@mcp-z/mcp-sheets';
import { CallToolResultSchema } from '@modelcontextprotocol/sdk/types.js';
import assert from 'assert';
import { createStdioFixture } from '../lib/create-stdio-fixture.ts';
import { createMetadataFixture, METADATA_SHEET_ID, METADATA_SHEET_TITLE } from '../lib/metadata-fixture.ts';

describe('Sheets MCP Server Component Tests', () => {
  let client: ManagedClient;
  let fixture: ReturnType<typeof createStdioFixture>;

  before(async () => {
    fixture = createStdioFixture('live');
    client = await fixture.connect();
  });

  after(async () => {
    if (fixture) await fixture.close();
  });

  describe('MCP Protocol Component Testing', () => {
    describe('Live metadata invocation', () => {
      let metadataFixture: Awaited<ReturnType<typeof createMetadataFixture>>;

      before(async () => {
        metadataFixture = await createMetadataFixture();
      });

      after(async () => {
        if (metadataFixture) await metadataFixture.close();
      });

      it('invokes both metadata tools through the built package and preserves structured and text results', async () => {
        const workbook = CallToolResultSchema.parse(await client.callTool({ name: 'spreadsheet-metadata-get', arguments: { id: metadataFixture.id } }));
        assert.ok(!workbook.isError);
        const workbookResult = mcp.toolFactories.spreadsheetMetadataGet().config.outputSchema.parse(workbook.structuredContent).result;
        if (workbookResult.type !== 'success') assert.fail('Expected authenticated workbook metadata');
        assert.equal(workbookResult.metadata.spreadsheetId, metadataFixture.id);
        assert.equal(workbookResult.metadata.properties?.locale, 'en_CA');
        assert.equal(workbookResult.metadata.sheets?.length, 2);
        const workbookSheets = workbookResult.metadata.sheets ?? [];
        assert.deepEqual(workbookSheets.map((sheet) => sheet.properties?.title).sort(), ['Empty tab', METADATA_SHEET_TITLE].sort());
        for (const sheet of workbookSheets) assert.equal(sheet.data, undefined);

        const cells = CallToolResultSchema.parse(await client.callTool({ name: 'cells-metadata-get', arguments: { id: metadataFixture.id, gid: METADATA_SHEET_ID, ranges: ['E8'] } }));
        assert.ok(!cells.isError);
        const cellsResult = mcp.toolFactories.cellsMetadataGet().config.outputSchema.parse(cells.structuredContent).result;
        if (cellsResult.type !== 'success') assert.fail('Expected authenticated cell metadata');
        assert.equal(cellsResult.gid, String(METADATA_SHEET_ID));
        const sheet = cellsResult.metadata.sheets?.[0];
        assert.equal(sheet?.properties?.title, METADATA_SHEET_TITLE);
        assert.equal(sheet?.data?.[0]?.startRow, 7);
        assert.equal(sheet?.data?.[0]?.startColumn, 4);
        const note = sheet?.data?.[0]?.rowData?.[0]?.values?.[0];
        assert.equal(note?.note, 'Selected note');
        assert.equal(note?.userEnteredValue, undefined);
        for (const { response, result } of [
          { response: workbook, result: workbookResult },
          { response: cells, result: cellsResult },
        ]) {
          const text = response.content.find((item) => item.type === 'text');
          assert.ok(text && text.type === 'text');
          assert.deepEqual(JSON.parse(text.text), result);
        }
      });
    });
  });
});
