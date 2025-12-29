import type { EnrichedExtra } from '@mcp-z/oauth-google';
import { schemas } from '@mcp-z/oauth-google';

const { AuthRequiredBranchSchema } = schemas;

import type { ToolModule } from '@mcp-z/server';
import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js';
import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { google } from 'googleapis';
import { z } from 'zod';
import { SpreadsheetIdSchema } from '../../schemas/index.js';

// Note: Using contextual descriptions for sourceId/newId since they describe different spreadsheets

const inputSchema = z.object({
  id: SpreadsheetIdSchema,
  newTitle: z.coerce.string().trim().min(1).optional().describe('Name for the copy (optional, defaults to "Copy of [original]")'),
});

// Success branch schema
const successBranchSchema = z.object({
  type: z.literal('success'),
  operationSummary: z.string().describe('Summary of the copy operation'),
  itemsProcessed: z.number().describe('Total items attempted (always 1)'),
  itemsChanged: z.number().describe('Successfully copied (always 1 on success)'),
  completedAt: z.string().describe('ISO datetime when operation completed'),
  sourceId: z.string().describe('Source spreadsheet ID'),
  sourceTitle: z.string().describe('Source spreadsheet title'),
  newId: z.string().describe('Copied spreadsheet ID'),
  newTitle: z.string().describe('Title of the copied spreadsheet'),
  spreadsheetUrl: z.string().describe('URL of the copied spreadsheet'),
});

// Output schema with auth_required support
const outputSchema = z.discriminatedUnion('type', [successBranchSchema, AuthRequiredBranchSchema]);

const config = {
  description: 'Copy an entire spreadsheet/workbook (all sheets, formatting, charts, named ranges, etc.). Creates in the same folder as the original. Uses Google Drive API.',
  inputSchema,
  outputSchema: z.object({
    result: outputSchema,
  }),
} as const;

export type Input = z.infer<typeof inputSchema>;
export type Output = z.infer<typeof outputSchema>;

async function handler({ id, newTitle }: Input, extra: EnrichedExtra): Promise<CallToolResult> {
  const logger = extra.logger;
  logger.info('sheets.spreadsheet.copy called', { id, newTitle });

  try {
    const drive = google.drive({ version: 'v3', auth: extra.authContext.auth });
    const sheets = google.sheets({ version: 'v4', auth: extra.authContext.auth });

    // Get the original spreadsheet title
    const sourceInfo = await sheets.spreadsheets.get({
      spreadsheetId: id,
      fields: 'properties.title',
    });

    const sourceTitle = sourceInfo.data.properties?.title || '';

    // Copy the spreadsheet using Drive API
    const copyResponse = await drive.files.copy({
      fileId: id,
      requestBody: {
        ...(newTitle && { name: newTitle }),
      },
    });

    const newId = copyResponse.data.id;
    const resultTitle = copyResponse.data.name || '';

    if (!newId) {
      throw new Error('Failed to retrieve new spreadsheet ID from API response');
    }

    logger.info('sheets.spreadsheet.copy success', { sourceId: id, newId, newTitle: resultTitle });

    const result: Output = {
      type: 'success' as const,
      operationSummary: `Copied spreadsheet "${sourceTitle}" to "${resultTitle}"`,
      itemsProcessed: 1,
      itemsChanged: 1,
      completedAt: new Date().toISOString(),
      sourceId: id,
      sourceTitle,
      newId,
      newTitle: resultTitle,
      spreadsheetUrl: `https://docs.google.com/spreadsheets/d/${newId}`,
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
    logger.error('sheets.spreadsheet.copy error', { error: message });

    throw new McpError(ErrorCode.InternalError, `Error copying spreadsheet: ${message}`, {
      stack: error instanceof Error ? error.stack : undefined,
    });
  }
}

export default function createTool() {
  return {
    name: 'spreadsheet-copy',
    config,
    handler,
  } satisfies ToolModule;
}
