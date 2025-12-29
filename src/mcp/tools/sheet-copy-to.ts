import type { EnrichedExtra } from '@mcp-z/oauth-google';
import { schemas } from '@mcp-z/oauth-google';

const { AuthRequiredBranchSchema } = schemas;

import type { ToolModule } from '@mcp-z/server';
import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js';
import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { google } from 'googleapis';
import { z } from 'zod';
import { SheetGidSchema, SpreadsheetIdSchema } from '../../schemas/index.js';

// Note: Using contextual descriptions for source/destination IDs since they describe different spreadsheets/sheets

const inputSchema = z.object({
  sourceId: SpreadsheetIdSchema.describe('Source spreadsheet ID'),
  sourceGid: SheetGidSchema.describe('Source sheet grid ID to copy'),
  destinationId: SpreadsheetIdSchema.describe('Destination spreadsheet ID'),
  newTitle: z.coerce.string().trim().min(1).optional().describe('New name for the copied sheet (optional, will use auto-generated name if not provided)'),
});

// Success branch schema
const successBranchSchema = z.object({
  type: z.literal('success'),
  operationSummary: z.string().describe('Summary of the copy operation'),
  itemsProcessed: z.number().describe('Total items attempted (always 1)'),
  itemsChanged: z.number().describe('Successfully copied (always 1 on success)'),
  completedAt: z.string().describe('ISO datetime when operation completed'),
  sourceId: z.string().describe('Source spreadsheet ID'),
  sourceGid: z.string().describe('Source sheet ID'),
  sourceTitle: z.string().describe('Source sheet title'),
  destinationId: z.string().describe('Destination spreadsheet ID'),
  destinationGid: z.string().describe('Copied sheet ID in destination'),
  destinationTitle: z.string().describe('Title of the copied sheet in destination'),
  sheetUrl: z.string().describe('URL of the copied sheet'),
  renamed: z.boolean().describe('Whether the sheet was renamed after copying'),
});

// Output schema with auth_required support
const outputSchema = z.discriminatedUnion('type', [successBranchSchema, AuthRequiredBranchSchema]);

const config = {
  description: 'Copy a sheet to another spreadsheet. Copies all data, formatting, and charts.',
  inputSchema,
  outputSchema: z.object({
    result: outputSchema,
  }),
} as const;

export type Input = z.infer<typeof inputSchema>;
export type Output = z.infer<typeof outputSchema>;

async function handler({ sourceId, sourceGid, destinationId, newTitle }: Input, extra: EnrichedExtra): Promise<CallToolResult> {
  const logger = extra.logger;
  logger.info('sheets.sheet.copyTo called', { sourceId, sourceGid, destinationId, newTitle });

  try {
    const sheets = google.sheets({ version: 'v4', auth: extra.authContext.auth });

    // Get source sheet info
    const sourceInfo = await sheets.spreadsheets.get({
      spreadsheetId: sourceId,
      fields: 'sheets.properties.sheetId,sheets.properties.title',
    });

    const sourceSheet = sourceInfo.data.sheets?.find((s) => String(s.properties?.sheetId) === sourceGid);
    if (!sourceSheet?.properties) {
      throw new McpError(ErrorCode.InvalidParams, `Source sheet with gid "${sourceGid}" not found in spreadsheet`);
    }

    const sourceTitle = sourceSheet.properties.title || '';

    // Copy the sheet to the destination spreadsheet
    const copyResponse = await sheets.spreadsheets.sheets.copyTo({
      spreadsheetId: sourceId,
      sheetId: Number(sourceGid),
      requestBody: {
        destinationSpreadsheetId: destinationId,
      },
    });

    const newSheetId = copyResponse.data.sheetId;
    let destinationTitle = copyResponse.data.title || '';

    if (!newSheetId) {
      throw new Error('Failed to retrieve new sheet ID from API response');
    }

    // If newTitle is provided, rename the sheet in the destination
    let renamed = false;
    if (newTitle && newTitle !== destinationTitle) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: destinationId,
        requestBody: {
          requests: [
            {
              updateSheetProperties: {
                properties: { sheetId: newSheetId, title: newTitle },
                fields: 'title',
              },
            },
          ],
        },
      });
      destinationTitle = newTitle;
      renamed = true;
    }

    logger.info('sheets.sheet.copyTo success', {
      sourceId,
      sourceGid,
      destinationId,
      destinationGid: String(newSheetId),
      renamed,
    });

    const result: Output = {
      type: 'success' as const,
      operationSummary: `Copied sheet "${sourceTitle}" to destination${renamed ? ` as "${destinationTitle}"` : ''}`,
      itemsProcessed: 1,
      itemsChanged: 1,
      completedAt: new Date().toISOString(),
      sourceId,
      sourceGid,
      sourceTitle,
      destinationId,
      destinationGid: String(newSheetId),
      destinationTitle,
      sheetUrl: `https://docs.google.com/spreadsheets/d/${destinationId}/edit#gid=${newSheetId}`,
      renamed,
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
    logger.error('sheets.sheet.copyTo error', { error: message });

    throw new McpError(ErrorCode.InternalError, `Error copying sheet to another spreadsheet: ${message}`, {
      stack: error instanceof Error ? error.stack : undefined,
    });
  }
}

export default function createTool() {
  return {
    name: 'sheet-copy-to',
    config,
    handler,
  } satisfies ToolModule;
}
