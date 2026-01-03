import type { EnrichedExtra } from '@mcp-z/oauth-google';
import { schemas } from '@mcp-z/oauth-google';

const { AuthRequiredBranchSchema } = schemas;

import type { ToolModule } from '@mcp-z/server';
import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js';
import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { google } from 'googleapis';
import { z } from 'zod';
import { SpreadsheetIdOutput, SpreadsheetIdSchema } from '../../schemas/index.ts';

const inputSchema = z.object({
  id: SpreadsheetIdSchema,
  newTitle: z.coerce.string().trim().min(1).describe('New name for the spreadsheet'),
});

// Success branch schema
const successBranchSchema = z.object({
  type: z.literal('success'),
  operationSummary: z.string().describe('Summary of the rename operation'),
  itemsProcessed: z.number().describe('Total items attempted (always 1)'),
  itemsChanged: z.number().describe('Successfully renamed (always 1 on success)'),
  completedAt: z.string().describe('ISO datetime when operation completed'),
  id: SpreadsheetIdOutput,
  spreadsheetUrl: z.string().describe('URL of the renamed spreadsheet'),
  oldTitle: z.string().describe('Previous title of the spreadsheet'),
  newTitle: z.string().describe('New title of the spreadsheet'),
});

// Output schema with auth_required support
const outputSchema = z.discriminatedUnion('type', [successBranchSchema, AuthRequiredBranchSchema]);

const config = {
  description: 'Rename a spreadsheet/workbook (the entire document, not individual sheets/tabs)',
  inputSchema,
  outputSchema: z.object({
    result: outputSchema,
  }),
} as const;

export type Input = z.infer<typeof inputSchema>;
export type Output = z.infer<typeof outputSchema>;

async function handler({ id, newTitle }: Input, extra: EnrichedExtra): Promise<CallToolResult> {
  const logger = extra.logger;
  logger.info('sheets.spreadsheet.rename called', { id, newTitle });

  try {
    const sheets = google.sheets({ version: 'v4', auth: extra.authContext.auth });

    // First, get the current spreadsheet title
    const spreadsheetInfo = await sheets.spreadsheets.get({
      spreadsheetId: id,
      fields: 'properties.title',
    });

    const oldTitle = spreadsheetInfo.data.properties?.title || '';

    // Rename the spreadsheet
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: id,
      requestBody: {
        requests: [
          {
            updateSpreadsheetProperties: {
              properties: { title: newTitle },
              fields: 'title',
            },
          },
        ],
      },
    });

    logger.info('sheets.spreadsheet.rename success', { id, oldTitle, newTitle });

    const result: Output = {
      type: 'success' as const,
      operationSummary: `Renamed spreadsheet "${oldTitle}" to "${newTitle}"`,
      itemsProcessed: 1,
      itemsChanged: 1,
      completedAt: new Date().toISOString(),
      id,
      spreadsheetUrl: `https://docs.google.com/spreadsheets/d/${id}`,
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
    const message = error instanceof Error ? error.message : String(error);
    logger.error('sheets.spreadsheet.rename error', { error: message });

    throw new McpError(ErrorCode.InternalError, `Error renaming spreadsheet: ${message}`, {
      stack: error instanceof Error ? error.stack : undefined,
    });
  }
}

export default function createTool() {
  return {
    name: 'spreadsheet-rename',
    config,
    handler,
  } satisfies ToolModule;
}
