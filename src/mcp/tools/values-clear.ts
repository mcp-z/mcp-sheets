import type { EnrichedExtra } from '@mcp-z/oauth-google';
import { schemas } from '@mcp-z/oauth-google';

const { AuthRequiredBranchSchema } = schemas;

import type { ToolModule } from '@mcp-z/server';
import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js';
import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { google } from 'googleapis';
import { z } from 'zod';
import { A1NotationSchema, SheetGidOutput, SheetGidSchema, SpreadsheetIdOutput, SpreadsheetIdSchema } from '../../schemas/index.js';

const inputSchema = z.object({
  id: SpreadsheetIdSchema,
  gid: SheetGidSchema,
  ranges: z.array(A1NotationSchema).min(1).describe('A1 notation ranges to clear (e.g., ["A1:B5", "D3:D10"]). Clears values only, preserves formatting.'),
});

const successBranchSchema = z.object({
  type: z.literal('success'),
  id: SpreadsheetIdOutput,
  gid: SheetGidOutput,
  spreadsheetTitle: z.string().describe('Title of the spreadsheet'),
  spreadsheetUrl: z.string().describe('URL of the spreadsheet'),
  sheetTitle: z.string().describe('Title of the sheet'),
  sheetUrl: z.string().describe('URL of the sheet'),
  clearedRanges: z.array(z.string()).describe('A1 notation ranges that were cleared'),
});

const outputSchema = z.discriminatedUnion('type', [successBranchSchema, AuthRequiredBranchSchema]);

const config = {
  description: 'Clear cell values from one or more ranges. Clears values only - preserves formatting, validation, and other cell properties. Use a1-notation prompt for range syntax.',
  inputSchema,
  outputSchema: z.object({
    result: outputSchema,
  }),
} as const;

export type Input = z.infer<typeof inputSchema>;
export type Output = z.infer<typeof outputSchema>;

async function handler({ id, gid, ranges }: Input, extra: EnrichedExtra): Promise<CallToolResult> {
  const logger = extra.logger;
  logger.info('sheets.values.clear called', { id, gid, rangeCount: ranges.length });

  try {
    const sheets = google.sheets({ version: 'v4', auth: extra.authContext.auth });

    // Get spreadsheet and sheet info in single API call
    const spreadsheetResponse = await sheets.spreadsheets.get({
      spreadsheetId: id,
      fields: 'properties.title,spreadsheetUrl,sheets.properties.sheetId,sheets.properties.title',
    });

    const spreadsheetData = spreadsheetResponse.data;
    const spreadsheetTitle = spreadsheetData.properties?.title ?? '';
    const spreadsheetUrl = spreadsheetData.spreadsheetUrl ?? '';

    // Find sheet by gid
    const sheet = spreadsheetData.sheets?.find((s) => String(s.properties?.sheetId) === gid);
    if (!sheet?.properties) {
      logger.info('Sheet not found for clear', { id, gid, rangeCount: ranges.length });
      throw new McpError(ErrorCode.InvalidParams, `Sheet not found: ${gid}`);
    }

    const sheetTitle = sheet.properties.title ?? gid;
    const sheetGid = sheet.properties.sheetId ?? 0;
    const sheetUrl = `https://docs.google.com/spreadsheets/d/${id}/edit#gid=${sheetGid}`;

    // Prefix ranges with sheet title
    const prefixedRanges = ranges.map((range) => `${sheetTitle}!${range}`);

    logger.info('sheets.values.clear executing', { spreadsheetId: id, sheetTitle, prefixedRanges });

    // Use batchClear for efficiency (works for single or multiple ranges)
    const clearResponse = await sheets.spreadsheets.values.batchClear({
      spreadsheetId: id,
      requestBody: {
        ranges: prefixedRanges,
      },
    });

    const clearedRanges = clearResponse.data.clearedRanges || [];

    logger.info('sheets.values.clear completed successfully', {
      spreadsheetId: id,
      sheetTitle,
      clearedRangesCount: clearedRanges.length,
    });

    const result: Output = {
      type: 'success' as const,
      id,
      gid: String(sheetGid),
      spreadsheetTitle: spreadsheetTitle || '',
      spreadsheetUrl: spreadsheetUrl || '',
      sheetTitle,
      sheetUrl,
      clearedRanges,
    };

    return {
      content: [{ type: 'text' as const, text: JSON.stringify(result) }],
      structuredContent: { result },
    };
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    logger.error('Clear operation failed', { id, gid, rangeCount: ranges.length, error: message });

    throw new McpError(ErrorCode.InternalError, `Error clearing values: ${message}`, {
      stack: error instanceof Error ? error.stack : undefined,
    });
  }
}

export default function createTool() {
  return {
    name: 'values-clear',
    config,
    handler,
  } satisfies ToolModule;
}
