/** Import CSV data to Google Sheets with range-based update (no deduplication) */

import type { EnrichedExtra } from '@mcp-z/oauth-google';
import { schemas } from '@mcp-z/oauth-google';

const { AuthRequiredBranchSchema } = schemas;

import type { ToolModule } from '@mcp-z/server';
import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js';
import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { parse } from 'csv-parse';
import { google } from 'googleapis';
import { z } from 'zod';
import { A1NotationSchema, SheetGidOutput, SheetGidSchema, SpreadsheetIdOutput, SpreadsheetIdSchema } from '../../schemas/index.ts';
import { getCsvReadStream } from '../../spreadsheet/csv-streaming.ts';

/** Batch size for Sheets API calls (1000 rows × avg 10 cols = 10K cells, well under 40K limit) */
const BATCH_SIZE = 1000;

const inputSchema = z.object({
  id: SpreadsheetIdSchema,
  gid: SheetGidSchema,
  sourceUri: z.string().trim().min(1).describe('CSV file URI (file://, http://, https://)'),
  startRange: A1NotationSchema.default('A1').describe('Top-left cell where CSV data starts (default: A1)'),
  valueInputOption: z.enum(['RAW', 'USER_ENTERED']).default('USER_ENTERED').describe('How to interpret values (RAW = exact, USER_ENTERED = parse formulas/dates)'),
  sourceHasHeaders: z.boolean().default(true).describe('First row is headers (metadata) - exclude from data range. Set to false to include first row as data.'),
});

const successBranchSchema = z.object({
  type: z.literal('success'),
  id: SpreadsheetIdOutput,
  gid: SheetGidOutput,
  spreadsheetTitle: z.string().describe('Spreadsheet title'),
  spreadsheetUrl: z.string().describe('Spreadsheet URL'),
  sheetTitle: z.string().describe('Sheet title'),
  sheetUrl: z.string().describe('Sheet URL'),
  updatedRange: z.string().describe('A1 notation range that was updated'),
  updatedRows: z.number().describe('Number of rows updated'),
  updatedColumns: z.number().describe('Number of columns updated'),
  updatedCells: z.number().describe('Number of cells updated'),
});

const outputSchema = z.discriminatedUnion('type', [successBranchSchema, AuthRequiredBranchSchema]);

const config = {
  description: 'Import CSV to sheet range. Overwrites existing data at startRange. Use rows-csv-append for database-style appends with deduplication.',
  inputSchema,
  outputSchema: z.object({
    result: outputSchema,
  }),
} as const;

export type Input = z.infer<typeof inputSchema>;
export type Output = z.infer<typeof outputSchema>;

async function handler({ id, gid, sourceUri, startRange, valueInputOption = 'USER_ENTERED', sourceHasHeaders }: Input, extra: EnrichedExtra): Promise<CallToolResult> {
  const logger = extra.logger;
  logger.info('sheets.values.csv-update called', { id, gid, sourceUri, startRange, valueInputOption, sourceHasHeaders });

  try {
    const sheets = google.sheets({ version: 'v4', auth: extra.authContext.auth });

    // Get spreadsheet and sheet info in single API call
    const spreadsheetResponse = await sheets.spreadsheets.get({
      spreadsheetId: id,
      fields: 'properties.title,spreadsheetUrl,sheets.properties.sheetId,sheets.properties.title',
    });

    const spreadsheetData = spreadsheetResponse.data;

    // Find the sheet by gid
    const sheet = spreadsheetData.sheets?.find((s) => String(s.properties?.sheetId) === gid);
    if (!sheet?.properties) {
      throw new McpError(ErrorCode.InvalidParams, `Sheet not found: ${gid}`);
    }

    const sheetTitle = sheet.properties.title ?? '';

    // Streaming CSV processing state
    let sourceHeaders: string[] = [];
    const allRows: (string | number | boolean | null)[][] = [];
    let totalCols = 0;

    // Get readable stream from CSV URI (no temp files!)
    const readStream = await getCsvReadStream(sourceUri);

    // Create CSV parser with native streaming
    const parser = readStream.pipe(
      parse({
        columns: !!sourceHasHeaders, // Parse first row as column names if source has headers
        skip_empty_lines: true,
        trim: true,
        cast: true, // Auto-convert numbers/booleans
        relax_column_count: true,
      })
    );

    // Stream and collect all rows (with batching for very large files)
    for await (const record of parser) {
      if (sourceHasHeaders) {
        // Extract source headers from first record
        if (sourceHeaders.length === 0) {
          sourceHeaders = Object.keys(record as Record<string, unknown>);
          totalCols = sourceHeaders.length;
          logger.info('sheets.values.csv-update source headers', { sourceHeaders, totalCols });
        }

        // Convert record to row array (exclude source headers from data range)
        // CSV values are strings/numbers/booleans/nulls from the parser
        const row = sourceHeaders.map((header) => (record as Record<string, string | number | boolean | null>)[header] ?? null);
        allRows.push(row);
      } else {
        // sourceHasHeaders=false: record is an array, include all rows (including first row)
        const row = record as (string | number | boolean | null)[];
        allRows.push(row);

        if (totalCols === 0) {
          totalCols = row.length;
        }
      }
    }

    if (allRows.length === 0) {
      throw new McpError(ErrorCode.InvalidParams, 'CSV file is empty');
    }

    // Prepare data for update (all rows)
    const dataToWrite: (string | number | boolean | null)[][] = allRows;

    // For large datasets, use batchUpdate to write in chunks
    // This respects the 40K cell limit per request
    const batchRequests = [];
    const currentRow = startRange.match(/[A-Z]+(\d+)/)?.[1] || '1';
    let currentRowNum = Number.parseInt(currentRow, 10);

    for (let i = 0; i < dataToWrite.length; i += BATCH_SIZE) {
      const batchData = dataToWrite.slice(i, i + BATCH_SIZE);
      const batchRange = `${sheetTitle}!${startRange.match(/[A-Z]+/)?.[0]}${currentRowNum}`;

      batchRequests.push({
        range: batchRange,
        values: batchData,
      });

      currentRowNum += batchData.length;
    }

    logger.info('sheets.values.csv-update batching', { totalBatches: batchRequests.length, batchSize: BATCH_SIZE });

    // Execute batch update
    const batchUpdateResponse = await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: id,
      requestBody: {
        valueInputOption,
        data: batchRequests,
      },
    });

    const totalUpdatedCells = batchUpdateResponse.data.totalUpdatedCells || 0;
    const totalUpdatedRows = batchUpdateResponse.data.totalUpdatedRows || 0;
    const totalUpdatedColumns = batchUpdateResponse.data.totalUpdatedColumns || 0;
    const firstUpdatedRange = batchUpdateResponse.data.responses?.[0]?.updatedRange || `${sheetTitle}!${startRange}`;

    logger.info('sheets.values.csv-update completed', { id, gid: sheet.properties.sheetId, updatedRange: firstUpdatedRange, updatedRows: totalUpdatedRows, sourceUri });

    const result: Output = {
      type: 'success' as const,
      id,
      gid: String(sheet.properties.sheetId ?? ''),
      spreadsheetTitle: spreadsheetData.properties?.title ?? '',
      spreadsheetUrl: spreadsheetData.spreadsheetUrl ?? `https://docs.google.com/spreadsheets/d/${id}`,
      sheetTitle,
      sheetUrl: `https://docs.google.com/spreadsheets/d/${id}/edit#gid=${sheet.properties.sheetId}`,
      updatedRange: firstUpdatedRange,
      updatedRows: totalUpdatedRows,
      updatedColumns: totalUpdatedColumns,
      updatedCells: totalUpdatedCells,
    };

    return {
      content: [{ type: 'text' as const, text: JSON.stringify(result) }],
      structuredContent: { result },
    };
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    logger.error('sheets.values.csv-update error', { error: message });

    throw new McpError(ErrorCode.InternalError, `Error updating values from CSV: ${message}`, {
      stack: error instanceof Error ? error.stack : undefined,
    });
  }
}

export default function createTool() {
  return {
    name: 'values-csv-update',
    config,
    handler,
  } satisfies ToolModule;
}
