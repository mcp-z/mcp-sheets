import type { EnrichedExtra } from '@mcp-z/oauth-google';
import { schemas } from '@mcp-z/oauth-google';

const { AuthRequiredBranchSchema } = schemas;

import type { ToolModule } from '@mcp-z/server';
import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js';
import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { google } from 'googleapis';
import { z } from 'zod';
import { SheetGidOutput, SheetGidSchema, SpreadsheetIdOutput, SpreadsheetIdSchema } from '../../schemas/index.ts';

const inputSchema = z.object({
  id: SpreadsheetIdSchema,
  gid: SheetGidSchema,
  newTitle: z.coerce.string().trim().min(1).describe('New name for the sheet tab'),
});

// Success branch schema
const successBranchSchema = z.object({
  type: z.literal('success'),
  operationSummary: z.string().describe('Summary of the rename operation'),
  itemsProcessed: z.number().describe('Total items attempted (always 1 for single sheet)'),
  itemsChanged: z.number().describe('Successfully renamed sheets (always 1 on success)'),
  completedAt: z.string().describe('ISO datetime when operation completed'),
  id: SpreadsheetIdOutput,
  gid: SheetGidOutput,
  sheetUrl: z.string().describe('URL of the renamed sheet'),
  oldTitle: z.string().describe('Previous title of the sheet'),
  newTitle: z.string().describe('New title of the sheet'),
});

// Output schema with auth_required support
const outputSchema = z.discriminatedUnion('type', [successBranchSchema, AuthRequiredBranchSchema]);

const config = {
  description: 'Rename a sheet within a spreadsheet.',
  inputSchema,
  outputSchema: z.object({
    result: outputSchema,
  }),
} as const;

export type Input = z.infer<typeof inputSchema>;
export type Output = z.infer<typeof outputSchema>;

async function handler({ id, gid, newTitle }: Input, extra: EnrichedExtra): Promise<CallToolResult> {
  const logger = extra.logger;
  logger.info('sheets.sheet.rename called', { id, gid, newTitle });

  try {
    const sheets = google.sheets({ version: 'v4', auth: extra.authContext.auth });

    // First, get the current sheet title
    const spreadsheetInfo = await sheets.spreadsheets.get({
      spreadsheetId: id,
      fields: 'sheets.properties.sheetId,sheets.properties.title',
    });

    const sheetInfo = spreadsheetInfo.data.sheets?.find((s) => String(s.properties?.sheetId) === gid);
    if (!sheetInfo?.properties) {
      throw new McpError(ErrorCode.InvalidParams, `Sheet with gid "${gid}" not found in spreadsheet`);
    }

    const oldTitle = sheetInfo.properties.title || '';

    // Rename the sheet
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: id,
      requestBody: {
        requests: [
          {
            updateSheetProperties: {
              properties: { sheetId: Number(gid), title: newTitle },
              fields: 'title',
            },
          },
        ],
      },
    });

    logger.info('sheets.sheet.rename success', { id, gid, oldTitle, newTitle });

    const result: Output = {
      type: 'success' as const,
      operationSummary: `Renamed sheet "${oldTitle}" to "${newTitle}"`,
      itemsProcessed: 1,
      itemsChanged: 1,
      completedAt: new Date().toISOString(),
      id,
      gid,
      sheetUrl: `https://docs.google.com/spreadsheets/d/${id}/edit#gid=${gid}`,
      oldTitle,
      newTitle,
    };

    return {
      content: [
        {
          type: 'text' as const,
          text: JSON.stringify(result),
        },
      ],
      structuredContent: { result },
    };
  } catch (error) {
    if (error instanceof McpError) {
      throw error;
    }
    const message = error instanceof Error ? error.message : String(error);
    logger.error('sheets.sheet.rename error', { error: message });

    throw new McpError(ErrorCode.InternalError, `Error renaming sheet: ${message}`, {
      stack: error instanceof Error ? error.stack : undefined,
    });
  }
}

export default function createTool() {
  return {
    name: 'sheet-rename',
    config,
    handler,
  } satisfies ToolModule;
}
