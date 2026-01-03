import type { EnrichedExtra } from '@mcp-z/oauth-google';
import { schemas } from '@mcp-z/oauth-google';

const { AuthRequiredBranchSchema } = schemas;

import type { ToolModule } from '@mcp-z/server';
import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js';
import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { google } from 'googleapis';
import { z } from 'zod';
import { SheetGidOutput, SheetGidSchema, SpreadsheetIdOutput, SpreadsheetIdSchema } from '../../schemas/index.ts';
import { appendRows, mapRowsToHeader } from '../../spreadsheet/data-operations.ts';
import { ensureTabAndHeaders } from '../../spreadsheet/sheet-operations.ts';

const inputSchema = z.object({
  id: SpreadsheetIdSchema,
  gid: SheetGidSchema,
  rows: z
    .array(z.array(z.union([z.string(), z.number(), z.boolean(), z.null()])).min(1))
    .min(1)
    .describe('Array of rows, where each row is an array of cell values. Use null to skip a cell (preserve existing value), empty string "" to clear it.'),
  headers: z.array(z.string()).optional().describe('Column order/names - used for blank sheets or column mapping'),
  deduplicateBy: z.array(z.string()).optional().describe('Column names to use as composite key for deduplication'),
});

const successBranchSchema = z.object({
  type: z.literal('success'),
  id: SpreadsheetIdOutput,
  gid: SheetGidOutput,
  sheetTitle: z.string().describe('Sheet tab name'),
  updatedRows: z.number().describe('Number of rows appended'),
  rowsSkipped: z.number().optional().describe('Number of duplicate rows skipped'),
  sheetUrl: z.string().optional().describe('URL to view the sheet'),
});

const outputSchema = z.discriminatedUnion('type', [successBranchSchema, AuthRequiredBranchSchema]);

const config = {
  description: 'Add new rows to the bottom of an existing sheet with smart header handling and optional deduplication. BEST FOR: Structured database operations where spreadsheet has headers defining schema and rows represent records.',
  inputSchema,
  outputSchema: z.object({
    result: outputSchema,
  }),
} as const;

export type Input = z.infer<typeof inputSchema>;
export type Output = z.infer<typeof outputSchema>;

async function handler({ id, gid, rows, headers, deduplicateBy }: Input, extra: EnrichedExtra): Promise<CallToolResult> {
  const logger = extra.logger;
  logger.info('sheets.rows.append called', { id, gid, rowCount: rows.length, headers, deduplicateBy });

  try {
    const sheets = google.sheets({ version: 'v4', auth: extra.authContext.auth });

    // Get sheet details using the gid
    const spreadsheetResponse = await sheets.spreadsheets.get({
      spreadsheetId: id,
      fields: 'sheets.properties.sheetId,sheets.properties.title',
    });

    const sheet = spreadsheetResponse.data.sheets?.find((s) => String(s.properties?.sheetId) === gid);

    if (!sheet) {
      throw new McpError(ErrorCode.InvalidParams, 'Sheet not found');
    }

    const sheetTitle = sheet?.properties?.title ?? '';

    // Smart header handling and data processing
    let processedRows = rows;
    let currentHeaders: string[] = [];
    let rowsSkipped = 0;
    let existingKeySet: Set<string> = new Set();

    if (headers && headers.length > 0) {
      // Use ensureTabAndHeaders to handle header logic
      const headerResult = await ensureTabAndHeaders(sheets, {
        spreadsheetId: id,
        sheetTitle,
        requiredHeader: headers,
        keyColumns: deduplicateBy || [],
        logger,
      });

      currentHeaders = headerResult.header;
      existingKeySet = headerResult.keySet;

      // Map data rows to match current header order
      processedRows = mapRowsToHeader({ rows, header: currentHeaders, canonical: headers }) as (string | number | boolean | null)[][];
    }

    // Append rows with optional deduplication
    // If headers are empty but deduplication is requested, skip deduplication for empty spreadsheets
    const effectiveKeyColumns = currentHeaders.length === 0 && deduplicateBy && deduplicateBy.length > 0 ? [] : deduplicateBy || [];

    const appendResult = await appendRows(sheets, {
      spreadsheetId: id,
      sheetTitle,
      rows: processedRows,
      keySet: existingKeySet,
      keyColumns: effectiveKeyColumns,
      header: currentHeaders,
      logger,
    });

    rowsSkipped = appendResult.rowsSkipped || 0;
    // Only count data rows, not headers - headers are metadata, not items
    const updatedRows = appendResult.updatedRows || 0;

    const sheetUrl = `https://docs.google.com/spreadsheets/d/${id}/edit#gid=${gid}`;

    logger.info('sheets.rows.append completed', { id, gid, updatedRows, rowsSkipped });

    const result: Output = {
      type: 'success' as const,
      id,
      gid,
      sheetTitle,
      updatedRows,
      rowsSkipped, // Always include rowsSkipped, even when 0
      sheetUrl,
    };

    return {
      content: [{ type: 'text' as const, text: JSON.stringify(result) }],
      structuredContent: { result },
    };
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    logger.error('sheets.rows.append error', { error: message });

    throw new McpError(ErrorCode.InternalError, `Error appending rows: ${message}`, {
      stack: error instanceof Error ? error.stack : undefined,
    });
  }
}

export default function createTool() {
  return {
    name: 'rows-append',
    config,
    handler,
  } satisfies ToolModule;
}
