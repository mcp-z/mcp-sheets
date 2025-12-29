import type { EnrichedExtra } from '@mcp-z/oauth-google';
import { schemas } from '@mcp-z/oauth-google';

const { AuthRequiredBranchSchema } = schemas;

import type { ToolModule } from '@mcp-z/server';
import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js';
import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { google, type sheets_v4 } from 'googleapis';
import { z } from 'zod';
import { SheetGidOutput, SheetGidSchema, SpreadsheetIdOutput, SpreadsheetIdSchema } from '../../schemas/index.js';
import { parseA1Notation, rangeReferenceToGridRange } from '../../spreadsheet/range-operations.js';

// RGB color schema (0-1 range for Google Sheets API)
const ColorSchema = z.object({
  red: z.number().min(0).max(1).describe('Red component (0-1)'),
  green: z.number().min(0).max(1).describe('Green component (0-1)'),
  blue: z.number().min(0).max(1).describe('Blue component (0-1)'),
});

// Number format schema
const NumberFormatSchema = z.object({
  type: z.enum(['TEXT', 'NUMBER', 'PERCENT', 'CURRENCY', 'DATE', 'TIME']).describe('Number format type'),
  pattern: z.string().optional().describe('Custom format pattern (e.g., "$#,##0.00" for currency)'),
});

// Border schema
const BorderSchema = z.object({
  style: z.enum(['SOLID', 'DASHED', 'DOTTED']).describe('Border line style'),
  color: ColorSchema.describe('Border color'),
});

// Input schema for format requests
const FormatRequestSchema = z.object({
  range: z.string().min(1).describe('A1 notation range to format (e.g., "A1:D10", "B:B", "5:5")'),
  backgroundColor: ColorSchema.optional().describe('Cell background color'),
  textColor: ColorSchema.optional().describe('Text color'),
  bold: z.boolean().optional().describe('Bold text'),
  fontSize: z.number().int().min(6).max(36).optional().describe('Font size in points'),
  horizontalAlignment: z.enum(['LEFT', 'CENTER', 'RIGHT']).optional().describe('Horizontal text alignment'),
  numberFormat: NumberFormatSchema.optional().describe('Number format pattern'),
  borders: BorderSchema.optional().describe('Cell borders'),
});

const inputSchema = z.object({
  id: SpreadsheetIdSchema,
  gid: SheetGidSchema,
  requests: z.array(FormatRequestSchema).min(1).max(50).describe('Array of formatting requests. Batch multiple ranges for efficiency.'),
});

const successBranchSchema = z.object({
  type: z.literal('success'),
  id: SpreadsheetIdOutput,
  gid: SheetGidOutput,
  sheetTitle: z.string().describe('Title of the formatted sheet'),
  sheetUrl: z.string().describe('URL of the formatted sheet'),
  successCount: z.number().int().nonnegative().describe('Number of format requests successfully applied'),
  failedRanges: z
    .array(
      z.object({
        range: z.string().describe('A1 notation of range that failed'),
        error: z.string().describe('Why formatting failed for this range'),
      })
    )
    .optional()
    .describe('Only populated if some ranges failed to format'),
});

const outputSchema = z.discriminatedUnion('type', [successBranchSchema, AuthRequiredBranchSchema]);

const config = {
  description: 'Apply formatting (colors, borders, fonts, alignment, number formats) to cell ranges without modifying data. Supports batch operations for efficiency. Colors use 0-1 RGB format. Best for creating professional, visually organized spreadsheets.',
  inputSchema,
  outputSchema: z.object({
    result: outputSchema,
  }),
} as const;

export type Input = z.infer<typeof inputSchema>;
export type Output = z.infer<typeof outputSchema>;

async function handler({ id, gid, requests }: Input, extra: EnrichedExtra): Promise<CallToolResult> {
  const logger = extra.logger;
  logger.info('sheets.cells.format called', {
    id,
    gid,
    requestCount: requests.length,
  });

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
      logger.info('Sheet not found for format cells', { id, gid, requestCount: requests.length });
      throw new McpError(ErrorCode.InvalidParams, `Sheet not found: ${gid}`);
    }

    const sheetTitle = sheet.properties.title ?? gid;
    const sheetId = sheet.properties.sheetId;
    const sheetUrl = `https://docs.google.com/spreadsheets/d/${id}/edit#gid=${sheetId}`;

    // Build batch update requests
    const batchRequests: sheets_v4.Schema$Request[] = [];
    const failedRanges: Array<{ range: string; error: string }> = [];

    for (const request of requests) {
      try {
        // Parse A1 notation to range reference
        const rangeRef = parseA1Notation(request.range);

        // Build grid range from range reference using helper function
        const gridRange = rangeReferenceToGridRange(rangeRef, sheetId);

        // Build cell format object
        const cellFormat: {
          backgroundColor?: sheets_v4.Schema$Color;
          textFormat?: sheets_v4.Schema$TextFormat;
          horizontalAlignment?: string;
          numberFormat?: sheets_v4.Schema$NumberFormat;
        } = {};
        const fields: string[] = [];

        // Background color
        if (request.backgroundColor) {
          cellFormat.backgroundColor = request.backgroundColor;
          fields.push('backgroundColor');
        }

        // Text format
        if (request.textColor || request.bold !== undefined || request.fontSize !== undefined) {
          cellFormat.textFormat = {};
          if (request.textColor) {
            cellFormat.textFormat.foregroundColor = request.textColor;
            fields.push('textFormat.foregroundColor');
          }
          if (request.bold !== undefined) {
            cellFormat.textFormat.bold = request.bold;
            fields.push('textFormat.bold');
          }
          if (request.fontSize !== undefined) {
            cellFormat.textFormat.fontSize = request.fontSize;
            fields.push('textFormat.fontSize');
          }
        }

        // Horizontal alignment
        if (request.horizontalAlignment) {
          cellFormat.horizontalAlignment = request.horizontalAlignment;
          fields.push('horizontalAlignment');
        }

        // Number format
        if (request.numberFormat) {
          const numberFormat: sheets_v4.Schema$NumberFormat = {
            type: request.numberFormat.type,
          };
          if (request.numberFormat.pattern !== undefined) {
            numberFormat.pattern = request.numberFormat.pattern;
          }
          cellFormat.numberFormat = numberFormat;
          fields.push('numberFormat');
        }

        // Add repeatCell request for this range
        if (fields.length > 0) {
          batchRequests.push({
            repeatCell: {
              range: gridRange,
              cell: {
                userEnteredFormat: cellFormat,
              },
              fields: `userEnteredFormat(${fields.join(',')})`,
            },
          });
        }

        // Add border formatting if specified
        if (request.borders) {
          batchRequests.push({
            updateBorders: {
              range: gridRange,
              top: {
                style: request.borders.style,
                color: request.borders.color,
              },
              bottom: {
                style: request.borders.style,
                color: request.borders.color,
              },
              left: {
                style: request.borders.style,
                color: request.borders.color,
              },
              right: {
                style: request.borders.style,
                color: request.borders.color,
              },
            },
          });
        }
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        logger.info('Failed to parse range for formatting', {
          range: request.range,
          error: message,
        });
        failedRanges.push({
          range: request.range,
          error: `Failed to parse range: ${message}`,
        });
      }
    }

    // Early return if all ranges failed
    if (batchRequests.length === 0) {
      const result: Output = {
        type: 'success' as const,
        id,
        gid: String(sheetId),
        sheetTitle,
        sheetUrl,
        successCount: 0,
        failedRanges: failedRanges.length > 0 ? failedRanges : undefined,
      };

      return {
        content: [{ type: 'text' as const, text: JSON.stringify(result) }],
        structuredContent: { result },
      };
    }

    logger.info('sheets.cells.format executing batch request', {
      spreadsheetId: id,
      sheetTitle,
      batchRequestsCount: batchRequests.length,
    });

    // Execute the batch update
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: id,
      requestBody: {
        requests: batchRequests,
      },
    });

    logger.info('sheets.cells.format completed successfully', {
      successCount: requests.length - failedRanges.length,
      failedCount: failedRanges.length,
    });

    const result: Output = {
      type: 'success' as const,
      id,
      gid: String(sheetId),
      sheetTitle,
      sheetUrl,
      successCount: requests.length - failedRanges.length,
      failedRanges: failedRanges.length > 0 ? failedRanges : undefined,
    };

    return {
      content: [{ type: 'text' as const, text: JSON.stringify(result) }],
      structuredContent: { result },
    };
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    logger.error('Format cells operation failed', {
      id,
      gid,
      requestCount: requests.length,
      error: message,
    });

    throw new McpError(ErrorCode.InternalError, `Error formatting cells: ${message}`, {
      stack: error instanceof Error ? error.stack : undefined,
    });
  }
}

export default function createTool() {
  return {
    name: 'cells-format',
    config,
    handler,
  } satisfies ToolModule;
}
