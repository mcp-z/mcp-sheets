/** Get column names from Google Sheet (peek at first row only) */

import type { EnrichedExtra } from '@mcp-z/oauth-google';
import { schemas } from '@mcp-z/oauth-google';

const { AuthRequiredBranchSchema } = schemas;

import type { ToolModule } from '@mcp-z/server';
import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js';
import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { google } from 'googleapis';
import { z } from 'zod';
import { SheetGidSchema, SpreadsheetIdSchema } from '../../schemas/index.js';

const inputSchema = z.object({
  id: SpreadsheetIdSchema,
  gid: SheetGidSchema,
  render: z.enum(['FORMATTED_VALUE', 'UNFORMATTED_VALUE', 'FORMULA']).optional().describe('How to render cell values. FORMATTED_VALUE (default): calculated with formatting. UNFORMATTED_VALUE: calculated without formatting. FORMULA: show formula text instead of result.'),
});

// Success branch schema - uses columns: for consistency with standard vocabulary
const successBranchSchema = z.object({
  type: z.literal('success'),
  columns: z.array(z.string()).describe('First row values (column names) or empty if no rows'),
  isEmpty: z.boolean().describe('True if sheet has zero rows'),
});

const outputSchema = z.discriminatedUnion('type', [successBranchSchema, AuthRequiredBranchSchema]);

const config = {
  description: 'Get first row from Google Sheet. Returns columns array and isEmpty flag.',
  inputSchema,
  outputSchema: z.object({
    result: outputSchema,
  }),
} as const;

export type Input = z.infer<typeof inputSchema>;
export type Output = z.infer<typeof outputSchema>;

async function handler({ id, gid, render }: Input, extra: EnrichedExtra): Promise<CallToolResult> {
  const logger = extra.logger;
  logger.debug?.('sheets.columns.get called', { id, gid, render });

  try {
    const sheets = google.sheets({ version: 'v4', auth: extra.authContext.auth });

    // Get spreadsheet and sheet info in single API call
    const spreadsheetResponse = await sheets.spreadsheets.get({
      spreadsheetId: id,
      fields: 'sheets.properties.sheetId,sheets.properties.title',
    });

    // Find sheet by gid
    const sheet = spreadsheetResponse.data.sheets?.find((s) => String(s.properties?.sheetId) === gid);
    if (!sheet?.properties) {
      throw new McpError(ErrorCode.InvalidParams, `Sheet not found: ${gid}`);
    }

    const sheetTitle = sheet.properties.title ?? '';

    // Read first row only (A1:ZZZ1 should cover most reasonable sheets)
    const range = `'${sheetTitle}'!A1:ZZZ1`;
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: id,
      range,
      valueRenderOption: render || 'FORMATTED_VALUE',
    });

    const rows = response.data.values || [];
    const isEmpty = rows.length === 0;

    // Extract first row and convert to strings (column names)
    const firstRow = rows[0] || [];
    const columns = firstRow.map((value) => String(value ?? ''));

    const result: Output = {
      type: 'success' as const,
      columns,
      isEmpty,
    };

    logger.info?.('sheets.columns.get completed', { id, gid, columnCount: columns.length, isEmpty });
    return {
      content: [{ type: 'text' as const, text: JSON.stringify(result) }],
      structuredContent: { result },
    };
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    logger.error?.('sheets.columns.get error', { error: message });
    throw new McpError(ErrorCode.InternalError, `Error getting sheet columns: ${message}`, {
      stack: error instanceof Error ? error.stack : undefined,
    });
  }
}

export default function createTool() {
  return {
    name: 'columns-get',
    config,
    handler,
  } satisfies ToolModule;
}
