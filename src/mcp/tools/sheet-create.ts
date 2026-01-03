import type { EnrichedExtra } from '@mcp-z/oauth-google';
import { schemas } from '@mcp-z/oauth-google';

const { AuthRequiredBranchSchema } = schemas;

import type { ToolModule } from '@mcp-z/server';
import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js';
import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { google } from 'googleapis';
import { z } from 'zod';
import { SheetGidOutput, SpreadsheetIdOutput, SpreadsheetIdSchema } from '../../schemas/index.ts';

const inputSchema = z.object({
  id: SpreadsheetIdSchema,
  sheetTitle: z.coerce.string().trim().min(1).describe('Name for the new sheet tab'),
});

// Success branch schema
const successBranchSchema = z.object({
  type: z.literal('success'),
  operationSummary: z.string().describe('Summary of the sheet creation operation'),
  itemsProcessed: z.number().describe('Total items attempted (always 1 for single sheet)'),
  itemsChanged: z.number().describe('Successfully created sheets (always 1 on success)'),
  completedAt: z.string().describe('ISO datetime when operation completed'),
  id: SpreadsheetIdOutput,
  gid: SheetGidOutput,
  sheetUrl: z.string().describe('URL of the created sheet'),
  sheetTitle: z.string().describe('Title of the created sheet'),
});

// Output schema with auth_required support
const outputSchema = z.discriminatedUnion('type', [successBranchSchema, AuthRequiredBranchSchema]);

const config = {
  description: 'Create a new sheet/tab in the spreadsheet/workbook',
  inputSchema,
  outputSchema: z.object({
    result: outputSchema,
  }),
} as const;

export type Input = z.infer<typeof inputSchema>;
export type Output = z.infer<typeof outputSchema>;

async function handler({ id, sheetTitle }: Input, extra: EnrichedExtra): Promise<CallToolResult> {
  const logger = extra.logger;
  logger.info('sheets.sheet.create called', { id, sheetTitle });

  try {
    const sheets = google.sheets({ version: 'v4', auth: extra.authContext.auth });
    const response = await sheets.spreadsheets.batchUpdate({
      spreadsheetId: id,
      requestBody: {
        requests: [{ addSheet: { properties: { title: sheetTitle } } }],
      },
    });
    const batchResult = response.data;
    const sheetId = batchResult.replies?.[0]?.addSheet?.properties?.sheetId;
    if (!sheetId) {
      throw new Error('Failed to retrieve sheetId from Google Sheets API response');
    }
    logger.info('sheets.sheet.create success', { id, sheetTitle, sheetId });

    const result: Output = {
      type: 'success' as const,
      operationSummary: `Created sheet "${sheetTitle}"`,
      itemsProcessed: 1,
      itemsChanged: 1,
      completedAt: new Date().toISOString(),
      id,
      gid: String(sheetId),
      sheetUrl: `https://docs.google.com/spreadsheets/d/${id}/edit#gid=${sheetId}`,
      sheetTitle,
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
    const message = error instanceof Error ? error.message : String(error);
    logger.error('sheets.sheet.create error', { error: message });

    // Throw McpError for proper MCP error handling
    throw new McpError(ErrorCode.InternalError, `Error creating sheet: ${message}`, {
      stack: error instanceof Error ? error.stack : undefined,
    });
  }
}

export default function createTool() {
  return {
    name: 'sheet-create',
    config,
    handler,
  } satisfies ToolModule;
}
