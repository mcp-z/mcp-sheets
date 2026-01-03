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
  dimension: z.enum(['ROWS', 'COLUMNS']).describe('Whether to move rows or columns'),
  startIndex: z.number().int().nonnegative().describe('Starting index of the range to move (0-based)'),
  endIndex: z.number().int().positive().describe('Ending index of the range to move (0-based, exclusive)'),
  destinationIndex: z.number().int().nonnegative().describe('Index where the rows/columns will be moved to (0-based)'),
});

const successBranchSchema = z.object({
  type: z.literal('success'),
  id: SpreadsheetIdOutput,
  gid: SheetGidOutput,
  spreadsheetTitle: z.string().describe('Title of the spreadsheet'),
  spreadsheetUrl: z.string().describe('URL of the spreadsheet'),
  sheetTitle: z.string().describe('Title of the sheet'),
  sheetUrl: z.string().describe('URL of the sheet'),
  dimension: z.enum(['ROWS', 'COLUMNS']).describe('Dimension that was moved'),
  sourceRange: z.object({
    startIndex: z.number().describe('Starting index of the moved range'),
    endIndex: z.number().describe('Ending index of the moved range'),
  }),
  destinationIndex: z.number().describe('Index where the range was moved to'),
  movedCount: z.number().describe('Number of rows/columns moved'),
});

const outputSchema = z.discriminatedUnion('type', [successBranchSchema, AuthRequiredBranchSchema]);

const config = {
  description: 'Move rows or columns within a sheet to a new position. Use 0-based indices. The destinationIndex is where the rows/columns will be moved TO.',
  inputSchema,
  outputSchema: z.object({
    result: outputSchema,
  }),
} as const;

export type Input = z.infer<typeof inputSchema>;
export type Output = z.infer<typeof outputSchema>;

async function handler({ id, gid, dimension, startIndex, endIndex, destinationIndex }: Input, extra: EnrichedExtra): Promise<CallToolResult> {
  const logger = extra.logger;
  logger.info('sheets.dimensions.move called', { id, gid, dimension, startIndex, endIndex, destinationIndex });

  // Validate indices
  if (startIndex >= endIndex) {
    throw new McpError(ErrorCode.InvalidParams, `startIndex (${startIndex}) must be less than endIndex (${endIndex})`);
  }

  // Check if destination is within source range (invalid move)
  if (destinationIndex > startIndex && destinationIndex < endIndex) {
    throw new McpError(ErrorCode.InvalidParams, `destinationIndex (${destinationIndex}) cannot be within the source range (${startIndex}-${endIndex})`);
  }

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
      logger.info('Sheet not found for move', { id, gid, dimension });
      throw new McpError(ErrorCode.InvalidParams, `Sheet not found: ${gid}`);
    }

    const sheetTitle = sheet.properties.title ?? gid;
    const sheetGid = sheet.properties.sheetId ?? 0;
    const sheetUrl = `https://docs.google.com/spreadsheets/d/${id}/edit#gid=${sheetGid}`;

    logger.info('sheets.dimensions.move executing', { spreadsheetId: id, sheetTitle, dimension, startIndex, endIndex, destinationIndex });

    // Execute the move dimension request
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: id,
      requestBody: {
        requests: [
          {
            moveDimension: {
              source: {
                sheetId: sheetGid,
                dimension,
                startIndex,
                endIndex,
              },
              destinationIndex,
            },
          },
        ],
      },
    });

    const movedCount = endIndex - startIndex;

    logger.info('sheets.dimensions.move completed successfully', {
      spreadsheetId: id,
      sheetTitle,
      dimension,
      movedCount,
    });

    const result: Output = {
      type: 'success' as const,
      id,
      gid: String(sheetGid),
      spreadsheetTitle: spreadsheetTitle || '',
      spreadsheetUrl: spreadsheetUrl || '',
      sheetTitle,
      sheetUrl,
      dimension,
      sourceRange: {
        startIndex,
        endIndex,
      },
      destinationIndex,
      movedCount,
    };

    return {
      content: [{ type: 'text' as const, text: JSON.stringify(result) }],
      structuredContent: { result },
    };
  } catch (error) {
    if (error instanceof McpError) {
      throw error;
    }
    const message = error instanceof Error ? error.message : String(error);
    logger.error('Move operation failed', { id, gid, dimension, startIndex, endIndex, destinationIndex, error: message });

    throw new McpError(ErrorCode.InternalError, `Error moving ${dimension.toLowerCase()}: ${message}`, {
      stack: error instanceof Error ? error.stack : undefined,
    });
  }
}

export default function createTool() {
  return {
    name: 'dimensions-move',
    config,
    handler,
  } satisfies ToolModule;
}
