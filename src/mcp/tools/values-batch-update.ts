import type { EnrichedExtra } from '@mcp-z/oauth-google';
import { schemas } from '@mcp-z/oauth-google';

const { AuthRequiredBranchSchema } = schemas;

import type { ToolModule } from '@mcp-z/server';
import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js';
import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { google } from 'googleapis';
import { z } from 'zod';
import { A1NotationSchema, SheetCellSchema, SheetGidOutput, SheetGidSchema, SpreadsheetIdOutput, SpreadsheetIdSchema } from '../../schemas/index.ts';

// Types for updated data response
interface UpdatedDataItem {
  range: string;
  majorDimension: 'ROWS' | 'COLUMNS';
  values?: (string | number | boolean | null)[][];
}

// Input schema for values batch update requests
const ValuesBatchUpdateRequestSchema = z.object({
  range: A1NotationSchema.describe('A1 notation range defining the bounded target area. Data dimensions must match range dimensions. Example: D1:D100 requires exactly 100 rows of data. Use open-ended ranges like D1:D to write any number of rows.'),
  values: z.array(z.array(SheetCellSchema)).min(1).describe('2D array of values. Row count must match range height, column count must match range width. Use null to skip a cell (preserve existing value), empty string "" to clear it.'),
  majorDimension: z.enum(['ROWS', 'COLUMNS']).describe('Whether values represent rows or columns'),
});

const inputSchema = z.object({
  id: SpreadsheetIdSchema,
  gid: SheetGidSchema,
  requests: z.array(ValuesBatchUpdateRequestSchema).min(1).describe('Array of value update requests'),
  valueInputOption: z.enum(['RAW', 'USER_ENTERED']).describe('How input data should be interpreted (RAW = exact values, USER_ENTERED = parsed like user input)'),
  includeData: z.boolean().describe('Whether to include updated cell values in the response'),
});

const successBranchSchema = z.object({
  type: z.literal('success'),
  id: SpreadsheetIdOutput,
  gid: SheetGidOutput,
  spreadsheetTitle: z.string().describe('Title of the updated spreadsheet'),
  spreadsheetUrl: z.string().describe('URL of the updated spreadsheet'),
  sheetTitle: z.string().describe('Title of the updated sheet'),
  sheetUrl: z.string().describe('URL of the updated sheet'),
  totalUpdatedRows: z.number().int().nonnegative().describe('Total number of rows updated across all requests'),
  totalUpdatedColumns: z.number().int().nonnegative().describe('Total number of columns updated across all requests'),
  totalUpdatedCells: z.number().int().nonnegative().describe('Total number of cells updated across all requests'),
  updatedRanges: z.array(z.string()).describe('A1 notation ranges that were updated'),
  updatedData: z
    .array(
      z.object({
        range: z.string().describe('A1 notation range that was updated'),
        majorDimension: z.enum(['ROWS', 'COLUMNS']).describe('Dimension of the updated data'),
        values: z.array(z.array(SheetCellSchema)).optional().describe('Updated values (if includeData was true)'),
      })
    )
    .optional()
    .describe('Detailed information about updated data (if includeData was true)'),
});

const outputSchema = z.discriminatedUnion('type', [successBranchSchema, AuthRequiredBranchSchema]);

const config = {
  description: 'Batch update multiple cell ranges. RAW=exact values, USER_ENTERED=parsed like user input. Use a1-notation prompt for range syntax.',
  inputSchema,
  outputSchema: z.object({
    result: outputSchema,
  }),
} as const;

export type Input = z.infer<typeof inputSchema>;
export type Output = z.infer<typeof outputSchema>;

async function handler({ id, gid, requests, valueInputOption = 'USER_ENTERED', includeData = false }: Input, extra: EnrichedExtra): Promise<CallToolResult> {
  const logger = extra.logger;
  logger.info('sheets.values.batchUpdate called', {
    id,
    gid,
    requestCount: requests.length,
    valueInputOption,
    includeData,
  });

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
      logger.info('Sheet not found for batch update', { id, gid, requestCount: requests.length });
      throw new McpError(ErrorCode.InvalidParams, `Sheet not found: ${gid}`);
    }

    const sheetTitle = sheet.properties.title ?? gid;
    const sheetGid = sheet.properties.sheetId ?? 0;
    const sheetUrl = `https://docs.google.com/spreadsheets/d/${id}/edit#gid=${sheetGid}`;

    // Build batch update request with prefixed ranges
    const batchUpdateData = requests.map((req) => ({
      range: `${sheetTitle}!${req.range}`,
      values: req.values,
      majorDimension: req.majorDimension || 'ROWS',
    }));

    logger.info('sheets.values.batchUpdate executing batch request', {
      spreadsheetId: id,
      sheetTitle,
      batchUpdateDataCount: batchUpdateData.length,
    });

    // Execute the batch update
    const batchUpdateResponse = await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: id,
      requestBody: {
        valueInputOption,
        data: batchUpdateData,
        includeValuesInResponse: includeData,
        responseDateTimeRenderOption: 'FORMATTED_STRING',
        responseValueRenderOption: 'FORMATTED_VALUE',
      },
    });

    const updateResult = batchUpdateResponse.data;

    // Validate batch operation results and detect partial failures
    const responses = updateResult.responses || [];
    const expectedCount = requests.length;
    const actualCount = responses.length;

    if (actualCount !== expectedCount) {
      logger.error('Partial batch failure detected', {
        expectedOperations: expectedCount,
        completedOperations: actualCount,
        spreadsheetId: id,
        sheetTitle,
      });

      throw new McpError(ErrorCode.InternalError, `Batch operation partially failed: ${actualCount}/${expectedCount} operations completed`);
    }

    // Check for any failed operations (empty or null responses)
    const failedOperations = responses.filter((response, index) => {
      if (!response || !response.updatedRange) {
        logger.error('Failed operation detected', {
          operationIndex: index,
          requestedRange: requests[index]?.range,
          spreadsheetId: id,
          sheetTitle,
        });
        return true;
      }
      return false;
    });

    if (failedOperations.length > 0) {
      throw new McpError(ErrorCode.InternalError, `${failedOperations.length} operations failed to update ranges`);
    }

    // Extract updated ranges and calculate totals
    const updatedRanges = responses.map((response) => response.updatedRange || '').filter((range) => range);
    const totalUpdatedRows = updateResult.totalUpdatedRows || 0;
    const totalUpdatedColumns = updateResult.totalUpdatedColumns || 0;
    const totalUpdatedCells = updateResult.totalUpdatedCells || 0;

    // Build updated data response if requested
    let updatedData: UpdatedDataItem[] | undefined;
    if (includeData && updateResult.responses) {
      updatedData = updateResult.responses
        .filter((response) => response.updatedData)
        .map((response) => {
          const item: UpdatedDataItem = {
            range: response.updatedRange || '',
            majorDimension: (response.updatedData?.majorDimension as 'ROWS' | 'COLUMNS' | undefined) || 'ROWS',
          };
          const values = response.updatedData?.values as (string | number | boolean | null | undefined)[][] | undefined;
          if (values !== undefined) {
            // Map undefined values to null for JSON Schema compatibility
            item.values = values.map((row) => row.map((cell) => (cell === undefined ? null : cell)));
          }
          return item;
        });
    }

    logger.info('sheets.values.batchUpdate completed successfully', {
      totalUpdatedRows,
      totalUpdatedColumns,
      totalUpdatedCells,
      updatedRangesCount: updatedRanges.length,
    });

    const result: Output = {
      type: 'success' as const,
      id,
      gid: String(sheetGid),
      spreadsheetTitle: spreadsheetTitle || '',
      spreadsheetUrl: spreadsheetUrl || '',
      sheetTitle,
      sheetUrl,
      totalUpdatedRows,
      totalUpdatedColumns,
      totalUpdatedCells,
      updatedRanges,
      updatedData,
    };

    return {
      content: [{ type: 'text' as const, text: JSON.stringify(result) }],
      structuredContent: { result },
    };
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    logger.error('Batch update operation failed', {
      id,
      gid,
      requestCount: requests.length,
      valueInputOption,
      error: message,
    });

    throw new McpError(ErrorCode.InternalError, `Error batch updating values: ${message}`, {
      stack: error instanceof Error ? error.stack : undefined,
    });
  }
}

export default function createTool() {
  return {
    name: 'values-batch-update',
    config,
    handler,
  } satisfies ToolModule;
}
