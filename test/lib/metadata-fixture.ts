// Network: disposable Google Sheets fixtures for metadata integration suites.
import { type sheets_v4, sheets as sheetsApi } from '@googleapis/sheets';
import createMiddlewareContext from './create-middleware-context.ts';
import { createTestSpreadsheet, deleteTestSpreadsheet } from './spreadsheet-helpers.ts';

export const METADATA_SHEET_TITLE = "Investor's ! Notes";
export const METADATA_SHEET_ID = 17;

export async function createMetadataFixture() {
  const context = await createMiddlewareContext();
  const id = await createTestSpreadsheet(await context.authProvider.getAccessToken(context.accountId), { title: `ci-metadata-${Date.now()}` });
  const sheets = sheetsApi({ version: 'v4', auth: context.auth });
  const source = { sheetId: METADATA_SHEET_ID, startRowIndex: 0, endRowIndex: 4, startColumnIndex: 0, endColumnIndex: 2 };
  const validation: sheets_v4.Schema$DataValidationRule = {
    condition: { type: 'ONE_OF_LIST', values: [{ userEnteredValue: 'Yes' }, { userEnteredValue: 'No' }] },
    strict: true,
    showCustomUi: true,
    inputMessage: 'Choose Yes or No',
  };
  try {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: id,
      requestBody: {
        requests: [
          { updateSpreadsheetProperties: { properties: { locale: 'en_CA', timeZone: 'America/Vancouver', autoRecalc: 'ON_CHANGE' }, fields: 'locale,timeZone,autoRecalc' } },
          { addSheet: { properties: { sheetId: METADATA_SHEET_ID, title: METADATA_SHEET_TITLE, gridProperties: { rowCount: 100, columnCount: 26, frozenRowCount: 2, frozenColumnCount: 1 } } } },
          { updateSheetProperties: { properties: { sheetId: 0, title: 'Empty tab', hidden: true }, fields: 'title,hidden' } },
          {
            updateCells: {
              start: { sheetId: METADATA_SHEET_ID, rowIndex: 0, columnIndex: 0 },
              rows: [
                { values: [{ userEnteredValue: { stringValue: 'Category' } }, { userEnteredValue: { stringValue: 'Amount' } }] },
                { values: [{ userEnteredValue: { stringValue: 'A' } }, { userEnteredValue: { numberValue: 10 } }] },
                { values: [{ userEnteredValue: { stringValue: 'B' } }, { userEnteredValue: { numberValue: 20 } }] },
                { values: [{ userEnteredValue: { stringValue: 'A' } }, { userEnteredValue: { numberValue: 30 } }] },
              ],
              fields: 'userEnteredValue',
            },
          },
          { addNamedRange: { namedRange: { name: 'MetadataSource', range: source } } },
          { addProtectedRange: { protectedRange: { range: source, description: 'Metadata protection', warningOnly: true } } },
          { addConditionalFormatRule: { index: 0, rule: { ranges: [source], booleanRule: { condition: { type: 'NUMBER_GREATER', values: [{ userEnteredValue: '15' }] }, format: { textFormat: { bold: true } } } } } },
          {
            addChart: {
              chart: {
                spec: {
                  title: 'Metadata chart',
                  basicChart: {
                    chartType: 'COLUMN',
                    headerCount: 1,
                    domains: [{ domain: { sourceRange: { sources: [{ ...source, endColumnIndex: 1 }] } } }],
                    series: [{ series: { sourceRange: { sources: [{ ...source, startColumnIndex: 1 }] } }, targetAxis: 'LEFT_AXIS' }],
                  },
                },
                position: { overlayPosition: { anchorCell: { sheetId: METADATA_SHEET_ID, rowIndex: 10, columnIndex: 12 } } },
              },
            },
          },
          {
            updateCells: {
              start: { sheetId: METADATA_SHEET_ID, rowIndex: 7, columnIndex: 4 },
              rows: [
                {
                  values: [
                    {
                      note: 'Selected note',
                      userEnteredValue: { formulaValue: '=1+1' },
                      userEnteredFormat: { textFormat: { bold: true } },
                    },
                  ],
                },
              ],
              fields: 'note,userEnteredValue,userEnteredFormat',
            },
          },
          { setDataValidation: { range: { sheetId: METADATA_SHEET_ID, startRowIndex: 6, endRowIndex: 7, startColumnIndex: 3, endColumnIndex: 4 }, rule: validation } },
          { updateCells: { start: { sheetId: METADATA_SHEET_ID, rowIndex: 19, columnIndex: 9 }, rows: [{ values: [{ note: 'Outside note', dataValidation: validation }] }], fields: 'note,dataValidation' } },
          {
            updateCells: {
              start: { sheetId: METADATA_SHEET_ID, rowIndex: 1, columnIndex: 12 },
              rows: [
                {
                  values: [
                    {
                      pivotTable: {
                        source,
                        rows: [{ sourceColumnOffset: 0, showTotals: true, sortOrder: 'ASCENDING' }],
                        values: [{ sourceColumnOffset: 1, summarizeFunction: 'SUM' }],
                      },
                    },
                  ],
                },
              ],
              fields: 'pivotTable',
            },
          },
        ],
      },
    });
  } catch (error) {
    await deleteTestSpreadsheet(await context.authProvider.getAccessToken(context.accountId), id, context.logger);
    throw error;
  }
  return {
    ...context,
    id,
    sheets,
    validation,
    async close() {
      await deleteTestSpreadsheet(await context.authProvider.getAccessToken(context.accountId), id, context.logger);
    },
  };
}
