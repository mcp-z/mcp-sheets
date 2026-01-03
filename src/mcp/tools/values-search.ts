import type { EnrichedExtra } from '@mcp-z/oauth-google';
import { schemas } from '@mcp-z/oauth-google';

const { AuthRequiredBranchSchema } = schemas;

import type { ToolModule } from '@mcp-z/server';
import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js';
import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { google } from 'googleapis';
import { z } from 'zod';
import { SheetCellSchema, SheetGidSchema, SpreadsheetIdSchema } from '../../schemas/index.ts';

// Helper to convert column index to letter (0 = A, 1 = B, etc.)
function columnIndexToLetter(index: number): string {
  let letter = '';
  let num = index + 1;
  while (num > 0) {
    const remainder = (num - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    num = Math.floor((num - 1) / 26);
  }
  return letter;
}

const inputSchema = z.object({
  id: SpreadsheetIdSchema,
  gid: SheetGidSchema,
  query: z.coerce.string().trim().optional().describe('Text to search for in sheet cells. If empty, returns all data at specified granularity.'),
  select: z.enum(['cells', 'rows', 'columns']).describe('Granularity: cells (individual matches), rows (full matching rows), columns (columns with matching headers)'),
  values: z.boolean().optional().describe('Include cell values in response'),
  a1s: z.boolean().optional().describe('Include A1 notation references in response'),
  render: z.enum(['FORMATTED_VALUE', 'UNFORMATTED_VALUE', 'FORMULA']).optional().describe('How to render cell values. FORMATTED_VALUE (default): calculated with formatting. UNFORMATTED_VALUE: calculated without formatting. FORMULA: show formula text instead of result.'),
  matchCase: z.boolean().optional().describe('Case-sensitive matching. Default is false (case-insensitive).'),
});

const successBranchSchema = z.object({
  type: z.literal('success'),
  count: z.number().int().nonnegative().describe('Number of matches found'),
  a1s: z.array(z.string()).optional().describe('A1 notation references for matches (e.g., "B5", "A5:D5", "B:B")'),
  values: z
    .array(z.union([z.array(SheetCellSchema), SheetCellSchema]))
    .optional()
    .describe('Cell values for matches (arrays for rows/columns, single values for cells)'),
});

const outputSchema = z.discriminatedUnion('type', [successBranchSchema, AuthRequiredBranchSchema]);

const config = {
  description: 'Search spreadsheet and return matches at cell, row, or column granularity. Use a1-notation prompt for range syntax.',
  inputSchema,
  outputSchema: z.object({
    result: outputSchema,
  }),
} as const;

export type Input = z.infer<typeof inputSchema>;
export type Output = z.infer<typeof outputSchema>;

async function handler({ id, gid, query, select, values = false, a1s = false, render, matchCase = false }: Input, extra: EnrichedExtra): Promise<CallToolResult> {
  const logger = extra.logger;
  logger.info('sheets.values-search called', { id, gid, query, select, values, a1s, render, matchCase });

  try {
    const sheets = google.sheets({ version: 'v4', auth: extra.authContext.auth });

    // Get sheet details including grid dimensions
    const spreadsheetResponse = await sheets.spreadsheets.get({
      spreadsheetId: id,
      fields: 'sheets.properties.sheetId,sheets.properties.title,sheets.properties.gridProperties',
    });

    const sheet = spreadsheetResponse.data.sheets?.find((s) => String(s.properties?.sheetId) === gid);

    if (!sheet) {
      logger.info('sheets.values-search sheet not found', { id, gid, query });
      throw new McpError(ErrorCode.InvalidParams, 'Sheet not found');
    }

    const sheetTitle = sheet.properties?.title ?? '';

    // Use actual sheet dimensions from gridProperties, with sensible defaults
    // Google Sheets default for new sheets is 1000 rows x 26 columns
    const rowCount = sheet.properties?.gridProperties?.rowCount ?? 1000;
    const columnCount = sheet.properties?.gridProperties?.columnCount ?? 26;
    const endColumn = columnIndexToLetter(columnCount - 1);
    const fullRange = `${sheetTitle}!A1:${endColumn}${rowCount}`;

    logger.debug?.('sheets.values-search fetching range', { fullRange, rowCount, columnCount });

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: id,
      range: fullRange,
      valueRenderOption: render || 'FORMATTED_VALUE',
    });
    const res = response.data;
    const allRows = Array.isArray(res.values) ? (res.values as (string | number | boolean | null)[][]) : [];

    // Prepare query for matching (case-insensitive by default)
    const searchQuery = query ? (matchCase ? query : query.toLowerCase()) : null;

    let count = 0;
    const a1Array: string[] = [];
    const valuesArray: (string | number | boolean | null | (string | number | boolean | null)[])[] = [];

    // Helper to check if a cell matches the query
    const cellMatches = (cell: string | number | boolean | null): boolean => {
      if (!searchQuery) return true; // If no query, match all
      if (typeof cell !== 'string') return false;
      const cellValue = matchCase ? cell : cell.toLowerCase();
      return cellValue.includes(searchQuery);
    };

    if (select === 'cells') {
      for (let rowIdx = 0; rowIdx < allRows.length; rowIdx++) {
        const row = allRows[rowIdx];
        if (!row) continue;

        for (let colIdx = 0; colIdx < row.length; colIdx++) {
          const cell = row[colIdx];
          if (cellMatches(cell)) {
            count++;
            if (a1s) {
              const colLetter = columnIndexToLetter(colIdx);
              const rowNum = rowIdx + 1;
              a1Array.push(`${colLetter}${rowNum}`);
            }
            if (values) {
              valuesArray.push(cell ?? null);
            }
          }
        }
      }
    } else if (select === 'rows') {
      const matchingRows: number[] = [];
      for (let rowIdx = 0; rowIdx < allRows.length; rowIdx++) {
        const row = allRows[rowIdx];
        if (!row) continue;

        const matches = row.some((cell) => cellMatches(cell));
        if (matches) {
          matchingRows.push(rowIdx);
        }
      }

      count = matchingRows.length;
      for (const rowIdx of matchingRows) {
        const row = allRows[rowIdx];
        if (!row) continue; // Skip if row is undefined
        const rowNum = rowIdx + 1;

        if (a1s) {
          const colEnd = columnIndexToLetter(Math.max(0, row.length - 1));
          a1Array.push(`A${rowNum}:${colEnd}${rowNum}`);
        }
        if (values) {
          valuesArray.push(row);
        }
      }
    } else if (select === 'columns') {
      const headerRow = allRows[0] || [];
      const matchingCols: number[] = [];

      for (let colIdx = 0; colIdx < headerRow.length; colIdx++) {
        const header = headerRow[colIdx];
        if (cellMatches(header)) {
          matchingCols.push(colIdx);
        }
      }

      count = matchingCols.length;
      for (const colIdx of matchingCols) {
        const colLetter = columnIndexToLetter(colIdx);

        if (a1s) {
          a1Array.push(`${colLetter}:${colLetter}`);
        }
        if (values) {
          const columnValues = allRows.map((row) => (row && row[colIdx] !== undefined ? row[colIdx] : null));
          valuesArray.push(columnValues);
        }
      }
    }

    logger.info('sheets.values-search results', { count, select, hasA1s: a1s, hasValues: values });

    const result: Output = {
      type: 'success' as const,
      count,
      ...(a1s && a1Array.length > 0 && { a1s: a1Array }),
      ...(values && valuesArray.length > 0 && { values: valuesArray }),
    };

    return {
      content: [{ type: 'text' as const, text: JSON.stringify(result) }],
      structuredContent: { result },
    };
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    logger.error('sheets.values-search error', { error: message });

    throw new McpError(ErrorCode.InternalError, `Error searching spreadsheet: ${message}`, {
      stack: error instanceof Error ? error.stack : undefined,
    });
  }
}

export default function createTool() {
  return {
    name: 'values-search',
    config,
    handler,
  } satisfies ToolModule;
}
