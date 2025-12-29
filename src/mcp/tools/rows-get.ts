import type { EnrichedExtra } from '@mcp-z/oauth-google';
import { schemas } from '@mcp-z/oauth-google';

const { AuthRequiredBranchSchema } = schemas;

import type { ToolModule } from '@mcp-z/server';
import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js';
import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { google } from 'googleapis';
import { z } from 'zod';
import { SheetCellSchema, SheetGidSchema, SpreadsheetIdSchema } from '../../schemas/index.js';

const inputSchema = z.object({
  id: SpreadsheetIdSchema,
  gid: SheetGidSchema,
  range: z.string().min(1).describe('A1 notation range to fetch (e.g., "B5", "A5:D5", "B:B", "5:5")'),
  render: z.enum(['FORMATTED_VALUE', 'UNFORMATTED_VALUE', 'FORMULA']).optional().describe('How to render cell values. FORMATTED_VALUE (default): calculated with formatting. UNFORMATTED_VALUE: calculated without formatting. FORMULA: show formula text instead of result.'),
});

// Success branch schema - uses rows: for consistency with standard vocabulary
const successBranchSchema = z.object({
  type: z.literal('success'),
  range: z.string().describe('The A1 notation range that was retrieved'),
  rows: z.array(z.array(SheetCellSchema)).describe('2D array of row data (each inner array is a row)'),
});

// Output schema with auth_required support
const outputSchema = z.discriminatedUnion('type', [successBranchSchema, AuthRequiredBranchSchema]);

const config = {
  description: 'Fetch row data from a specific range in A1 notation. Best used after values-search to get surrounding context. Use a1-notation prompt for syntax reference.',
  inputSchema,
  outputSchema: z.object({
    result: outputSchema,
  }),
} as const;

export type Input = z.infer<typeof inputSchema>;
export type Output = z.infer<typeof outputSchema>;

async function handler({ id, gid, range, render }: Input, extra: EnrichedExtra): Promise<CallToolResult> {
  const logger = extra.logger;
  logger.info('sheets.rows.get called', { id, gid, range, render });

  try {
    const sheets = google.sheets({ version: 'v4', auth: extra.authContext.auth });

    // Get sheet details using the gid to get sheet title
    const spreadsheetResponse = await sheets.spreadsheets.get({
      spreadsheetId: id,
      fields: 'sheets.properties.sheetId,sheets.properties.title',
    });

    const sheet = spreadsheetResponse.data.sheets?.find((s) => String(s.properties?.sheetId) === gid);

    if (!sheet) {
      logger.info('sheets.rows.get sheet not found', { id, gid, range });
      throw new McpError(ErrorCode.InvalidParams, `Sheet not found: ${gid}`);
    }

    const sheetTitle = sheet.properties?.title ?? '';

    // Construct full range with sheet title
    const fullRange = `${sheetTitle}!${range}`;

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: id,
      range: fullRange,
      valueRenderOption: render || 'FORMATTED_VALUE',
    });

    const res = response.data;
    const rows = Array.isArray(res.values) ? (res.values as (string | number | boolean | null)[][]) : [];

    logger.info('sheets.rows.get success', { id, gid, range, rowCount: rows.length });

    const result: Output = {
      type: 'success' as const,
      range: res.range || fullRange,
      rows,
    };

    return {
      content: [{ type: 'text' as const, text: JSON.stringify(result) }],
      structuredContent: { result },
    };
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    logger.error('sheets.rows.get error', { error: message });

    // Throw McpError for proper MCP error handling
    throw new McpError(ErrorCode.InternalError, `Error getting rows: ${message}`, {
      stack: error instanceof Error ? error.stack : undefined,
    });
  }
}

export default function createTool() {
  return {
    name: 'rows-get',
    config,
    handler,
  } satisfies ToolModule;
}
