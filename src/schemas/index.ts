import { z } from 'zod';

// Spreadsheet and sheet reference schemas
export const SpreadsheetRefSchema = z.string().min(1).describe('Spreadsheet reference: full URL, exact name, or partial name (not ID - use spreadsheet-find to get ID)');

export const SheetRefSchema = z.string().min(1).describe('Sheet reference: exact title or partial name match (not ID - use sheet-find to get ID)');

// Shared descriptions for consistency
const SPREADSHEET_ID_DESC = 'Spreadsheet ID (from URL d/{id})';
const SHEET_ID_DESC = 'Sheet ID (from URL gid={gid})';

// Input schemas (with validation)
// Note: z.coerce.string() converts numbers to strings, handling cases where MCP clients
// pass gid: 0 (number) instead of gid: "0" (string). This is critical for gid=0 which
// is the default sheet ID in new spreadsheets.
export const SpreadsheetIdSchema = z.string().min(1).describe(SPREADSHEET_ID_DESC);
export const SheetGidSchema = z.coerce.string().min(1).describe(SHEET_ID_DESC);

// Output schemas (for use in output/response schemas)
export const SpreadsheetIdOutput = z.string().describe(SPREADSHEET_ID_DESC);
export const SheetGidOutput = z.string().describe(SHEET_ID_DESC);

// Schema for individual sheet cells (can be string, number, boolean, or null)
// null represents empty cells - JSON cannot represent undefined, and Google API uses null for empty cells
// Infinity values are allowed as Google Sheets can contain them, but NaN is still rejected
export const SheetCellSchema = z.union([z.string(), z.number(), z.literal(Infinity), z.literal(-Infinity), z.boolean(), z.null()]);

// Schema for a row in a sheet search result
export const SheetRowSchema = z.object({
  rowIndex: z.number().int().positive().describe('1-based row number in the sheet'),
  values: z.array(SheetCellSchema).optional().describe('Cell values for this row (included when includeData=true)'),
  range: z.string().optional().describe('A1 notation range for this row (e.g., A5:D5)'),
});

// A1 notation validation - simplified for zod compatibility
export const A1NotationSchema = z.string().min(1).describe('A1 notation for cell range (e.g., A1, A1:B2, A:B, 1:2)');
