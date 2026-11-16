import type { sheets_v4 } from '@googleapis/sheets';
import { sheets as sheetsApi } from '@googleapis/sheets';
import type { EnrichedExtra } from '@mcp-z/oauth-google';
import { schemas } from '@mcp-z/oauth-google';

const { AuthRequiredBranchSchema } = schemas;

import type { CallToolResult, ToolModule } from '@mcp-z/server';
import { ProtocolError, ProtocolErrorCode } from '@mcp-z/server';
import { z } from 'zod';
import { PASTE_ORIENTATIONS, PASTE_TYPES } from '../../constants.ts';
import { googleAuth } from '../../lib/google-auth.ts';
import { A1NotationSchema, SheetGidOutput, SpreadsheetIdOutput, SpreadsheetIdSchema, StrictSheetGidSchema } from '../../schemas/index.ts';
import { calculatePasteRange, cellBlockToA1, parseA1Notation, parseFiniteCellBlock, rangeReferenceToGridRange, validateA1Notation } from '../../spreadsheet/range-operations.ts';

const inputSchema = z.object({
  id: SpreadsheetIdSchema,
  sourceGid: StrictSheetGidSchema.describe('Source sheet ID (from URL gid={gid})'),
  sourceRange: A1NotationSchema.describe('Finite sheet-local A1 range to copy (e.g., "A1:C10"; no sheet prefix, no whole rows/columns)'),
  destinationGid: StrictSheetGidSchema.describe('Destination sheet ID in the same spreadsheet; may equal sourceGid'),
  destinationRange: A1NotationSchema.describe('Top-left cell (e.g., "E1") or a finite range whose height and width are whole multiples of the (oriented) source; a larger range repeats the source to fill it'),
  pasteType: z
    .enum(PASTE_TYPES)
    .default('PASTE_NORMAL')
    .describe(
      'What to paste. PASTE_NORMAL (default) = values, formulas, formats and merges; PASTE_VALUES = values only, no formats/formulas/merges; PASTE_FORMAT = format only, excluding data validation; PASTE_NO_BORDERS = like PASTE_NORMAL without borders; PASTE_FORMULA = formulas only; PASTE_DATA_VALIDATION = data validation only; PASTE_CONDITIONAL_FORMATTING = conditional formatting rules only'
    ),
  pasteOrientation: z.enum(PASTE_ORIENTATIONS).default('NORMAL').describe('NORMAL, or TRANSPOSE (rows become columns and vice versa)'),
});

const successBranchSchema = z.object({
  type: z.literal('success'),
  id: SpreadsheetIdOutput,
  spreadsheetTitle: z.string().describe('Title of the spreadsheet'),
  spreadsheetUrl: z.string().describe('URL of the spreadsheet'),
  sourceGid: SheetGidOutput,
  sourceSheetTitle: z.string().describe('Title of the source sheet'),
  sourceRange: z.string().describe('A1 range that was copied'),
  destinationGid: SheetGidOutput,
  destinationSheetTitle: z.string().describe('Title of the destination sheet'),
  destinationRange: z.string().describe('Destination range as requested'),
  pastedRange: z.string().describe('A1 range of cells written in the destination sheet'),
  pasteType: z.enum(PASTE_TYPES).describe('Paste type applied'),
  pasteOrientation: z.enum(PASTE_ORIENTATIONS).describe('Orientation applied'),
  sheetUrl: z.string().describe('URL of the destination sheet'),
});

const outputSchema = z.discriminatedUnion('type', [successBranchSchema, AuthRequiredBranchSchema]);

const config = {
  description:
    'Copy a cell range to another range in the same spreadsheet (same or different tab), like Google Sheets copy and paste-special. Default PASTE_NORMAL copies values, formulas (relative references adjust), formats and merges; pasteType selects one native paste-special mode. Cannot copy between spreadsheets: use sheet-copy-to first. To duplicate a whole tab or workbook use sheet-copy, sheet-copy-to or spreadsheet-copy. Use a1-notation prompt for range syntax.',
  inputSchema,
  outputSchema: z.object({
    result: outputSchema,
  }),
} as const;

export type Input = z.infer<typeof inputSchema>;
export type Output = z.infer<typeof outputSchema>;
type SuccessOutput = z.infer<typeof successBranchSchema>;

/** The slice of the Sheets client this tool uses; tests pass a plain fake. */
export interface CellsCopySheets {
  spreadsheets: {
    get(params: { spreadsheetId: string; fields: string }): Promise<{ data: sheets_v4.Schema$Spreadsheet }>;
    batchUpdate(params: { spreadsheetId: string; requestBody: sheets_v4.Schema$BatchUpdateSpreadsheetRequest }): Promise<unknown>;
  };
}

function invalid(message: string): ProtocolError {
  return new ProtocolError(ProtocolErrorCode.InvalidParams, message);
}

/** Validates a finite sheet-local A1 range, converting any failure to InvalidParams. */
function checkRange(label: string, notation: string) {
  try {
    validateA1Notation(notation);
    return parseFiniteCellBlock(notation);
  } catch (error) {
    throw invalid(`Invalid ${label}: ${error instanceof Error ? error.message : String(error)}`);
  }
}

export async function copyCells(sheets: CellsCopySheets, { id, sourceGid, sourceRange, destinationGid, destinationRange, pasteType, pasteOrientation }: Input): Promise<SuccessOutput> {
  // All range checks happen before any network call
  const source = checkRange('sourceRange', sourceRange);
  checkRange('destinationRange', destinationRange);
  let pasted: ReturnType<typeof calculatePasteRange>;
  try {
    pasted = calculatePasteRange(sourceRange, destinationRange, pasteOrientation);
  } catch (error) {
    throw invalid(error instanceof Error ? error.message : String(error));
  }

  const spreadsheetResponse = await sheets.spreadsheets.get({
    spreadsheetId: id,
    fields: 'properties.title,spreadsheetUrl,sheets.properties.sheetId,sheets.properties.title,sheets.properties.gridProperties',
  });
  const spreadsheetData = spreadsheetResponse.data;
  const findSheet = (gid: string) => spreadsheetData.sheets?.find((s) => String(s.properties?.sheetId) === gid)?.properties;
  const sourceSheet = findSheet(sourceGid);
  if (!sourceSheet) throw invalid(`Sheet not found: ${sourceGid}`);
  const destinationSheet = findSheet(destinationGid);
  if (!destinationSheet) throw invalid(`Sheet not found: ${destinationGid}`);

  const gridSize = (sheet: sheets_v4.Schema$SheetProperties) => {
    const rows = sheet.gridProperties?.rowCount;
    const columns = sheet.gridProperties?.columnCount;
    if (typeof rows !== 'number' || typeof columns !== 'number') throw new Error(`Could not read the grid size of sheet "${sheet.title ?? sheet.sheetId}"`);
    return { rows, columns };
  };
  const sourceGrid = gridSize(sourceSheet);
  const destinationGrid = gridSize(destinationSheet);
  const fits = (grid: { rows: number; columns: number }, endRow: number, endColumn: number) => endRow <= grid.rows && endColumn <= grid.columns;
  if (!fits(sourceGrid, source.endRow, source.endColumn)) {
    throw invalid(`Source range ${sourceRange} is outside the grid of "${sourceSheet.title}" (${sourceGrid.rows} rows x ${sourceGrid.columns} columns)`);
  }
  if (!fits(destinationGrid, pasted.endRow, pasted.endColumn)) {
    throw invalid(`Pasting ${sourceRange} at ${destinationRange} would write ${cellBlockToA1(pasted)}, outside the grid of "${destinationSheet.title}" (${destinationGrid.rows} rows x ${destinationGrid.columns} columns)`);
  }

  // pasteType is always sent explicitly; there is no fallback path
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: {
      requests: [
        {
          copyPaste: {
            source: rangeReferenceToGridRange(parseA1Notation(sourceRange), Number(sourceGid)),
            destination: rangeReferenceToGridRange(parseA1Notation(destinationRange), Number(destinationGid)),
            pasteType,
            pasteOrientation,
          },
        },
      ],
    },
  });

  const destinationSheetId = destinationSheet.sheetId ?? Number(destinationGid);
  return {
    type: 'success' as const,
    id,
    spreadsheetTitle: spreadsheetData.properties?.title ?? '',
    spreadsheetUrl: spreadsheetData.spreadsheetUrl ?? '',
    sourceGid,
    sourceSheetTitle: sourceSheet.title ?? sourceGid,
    sourceRange,
    destinationGid,
    destinationSheetTitle: destinationSheet.title ?? destinationGid,
    destinationRange,
    pastedRange: cellBlockToA1(pasted),
    pasteType,
    pasteOrientation,
    sheetUrl: `https://docs.google.com/spreadsheets/d/${id}/edit#gid=${destinationSheetId}`,
  };
}

async function handler(input: Input, extra: EnrichedExtra): Promise<CallToolResult> {
  const logger = extra.logger;
  const { id, sourceGid, destinationGid, pasteType, pasteOrientation } = input;
  logger.info('sheets.cells.copy called', { id, sourceGid, destinationGid, pasteType, pasteOrientation });

  try {
    const sheets = sheetsApi({ version: 'v4', auth: googleAuth(extra.authContext.auth) });
    const result: Output = await copyCells(sheets as unknown as CellsCopySheets, input);
    logger.info('sheets.cells.copy completed successfully', { id, pastedRange: result.type === 'success' ? result.pastedRange : undefined });
    return {
      content: [{ type: 'text' as const, text: JSON.stringify(result) }],
      structuredContent: { result },
    };
  } catch (error) {
    if (error instanceof ProtocolError) throw error;
    const message = error instanceof Error ? error.message : String(error);
    logger.error('Copy cells failed', { id, sourceGid, destinationGid, error: message });
    throw new ProtocolError(ProtocolErrorCode.InternalError, `Error copying cells: ${message}`, {
      stack: error instanceof Error ? error.stack : undefined,
    });
  }
}

export default function createTool() {
  return {
    name: 'cells-copy',
    config,
    handler,
  } satisfies ToolModule;
}
