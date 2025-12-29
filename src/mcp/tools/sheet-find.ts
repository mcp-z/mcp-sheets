import type { EnrichedExtra } from '@mcp-z/oauth-google';
import { schemas } from '@mcp-z/oauth-google';

const { AuthRequiredBranchSchema } = schemas;

import type { ToolModule } from '@mcp-z/server';
import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js';
import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { google } from 'googleapis';
import { z } from 'zod';
import { SheetGidOutput, SheetRefSchema, SpreadsheetIdOutput, SpreadsheetIdSchema } from '../../schemas/index.js';
import { findSheetByRef } from '../../spreadsheet/sheet-operations.js';

const inputSchema = z.object({
  id: SpreadsheetIdSchema,
  sheetRef: SheetRefSchema,
});

const successBranchSchema = z.object({
  type: z.literal('success'),
  id: SpreadsheetIdOutput,
  gid: SheetGidOutput,
  title: z.string().describe('Sheet tab name'),
  sheetUrl: z.string().optional().describe('URL to view the sheet'),
});

const outputSchema = z.discriminatedUnion('type', [successBranchSchema, AuthRequiredBranchSchema]);

const config = {
  description: 'Find an existing sheet/tab within a known spreadsheet by title or GUID',
  inputSchema,
  outputSchema: z.object({
    result: outputSchema,
  }),
} as const;

export type Input = z.infer<typeof inputSchema>;
export type Output = z.infer<typeof outputSchema>;

async function handler({ id, sheetRef }: Input, extra: EnrichedExtra): Promise<CallToolResult> {
  const logger = extra.logger;
  logger.info('sheets.sheet.find called', { id, sheetRef });

  try {
    const sheets = google.sheets({ version: 'v4', auth: extra.authContext.auth });

    // Find sheet within the known spreadsheet
    const sheet = await findSheetByRef(sheets, id, sheetRef, logger);
    if (!sheet) {
      throw new McpError(ErrorCode.InvalidParams, 'Sheet not found');
    }

    const title = sheet?.properties?.title ?? String(sheetRef);
    const gid = sheet?.properties?.sheetId != null ? String(sheet.properties.sheetId) : '';
    const sheetUrl = gid ? `https://docs.google.com/spreadsheets/d/${id}/edit#gid=${gid}` : `https://docs.google.com/spreadsheets/d/${id}`;

    logger.info('sheets.sheet.find success', { id, gid, title });

    const result: Output = {
      type: 'success' as const,
      id,
      gid,
      title,
      sheetUrl,
    };

    return {
      content: [{ type: 'text' as const, text: JSON.stringify(result) }],
      structuredContent: { result },
    };
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    logger.error('sheets.sheet.find error', { error: message });

    throw new McpError(ErrorCode.InternalError, `Error finding sheet: ${message}`, {
      stack: error instanceof Error ? error.stack : undefined,
    });
  }
}

export default function createTool() {
  return {
    name: 'sheet-find',
    config,
    handler,
  } satisfies ToolModule;
}
