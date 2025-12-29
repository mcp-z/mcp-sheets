import type { EnrichedExtra } from '@mcp-z/oauth-google';
import { schemas } from '@mcp-z/oauth-google';

const { AuthRequiredBranchSchema } = schemas;

import type { ToolModule } from '@mcp-z/server';
import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js';
import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { google } from 'googleapis';
import { z } from 'zod';
import { SheetGidOutput, SheetGidSchema, SpreadsheetIdOutput, SpreadsheetIdSchema } from '../../schemas/index.js';
import { buildDimensionRequest, calculateAffectedCount, DEFAULT_COLUMN_COUNT, DEFAULT_ROW_COUNT, type DimensionRequest, MAX_COLUMN_COUNT, MAX_ROW_COUNT, sortOperations } from './lib/dimension-operations.js';

// Input schema for dimension batch update requests
const DimensionRequestSchema = z.object({
  operation: z.enum(['insertDimension', 'deleteDimension', 'appendDimension']).describe('Type of dimension operation to perform'),
  dimension: z.enum(['ROWS', 'COLUMNS']).describe('Whether to operate on rows or columns'),
  startIndex: z.number().int().nonnegative().describe('Starting index for the operation (0-based)'),
  endIndex: z.number().int().nonnegative().optional().describe('Ending index for the operation (0-based, exclusive). Optional - if omitted, the range is unbounded (extends to the end of the sheet)'),
  inheritFromBefore: z.boolean().optional().describe('For insertDimension: whether new rows/columns inherit properties from the row/column before the insertion point'),
});

const inputSchema = z.object({
  id: SpreadsheetIdSchema,
  gid: SheetGidSchema,
  requests: z.array(DimensionRequestSchema).min(1).describe('Array of dimension update requests'),
});

// Success branch schema
const successBranchSchema = z.object({
  type: z.literal('success'),
  id: SpreadsheetIdOutput,
  gid: SheetGidOutput,
  spreadsheetTitle: z.string().describe('Title of the updated spreadsheet'),
  spreadsheetUrl: z.string().describe('URL of the updated spreadsheet'),
  sheetTitle: z.string().describe('Title of the updated sheet'),
  sheetUrl: z.string().describe('URL of the updated sheet'),
  totalOperations: z.number().int().nonnegative().describe('Total number of dimension operations performed'),
  operationResults: z
    .array(
      z.object({
        operation: z.enum(['insertDimension', 'deleteDimension', 'appendDimension']).describe('Type of operation that was performed'),
        dimension: z.enum(['ROWS', 'COLUMNS']).describe('Dimension that was operated on'),
        startIndex: z.number().int().nonnegative().describe('Starting index of the operation'),
        endIndex: z.number().int().nonnegative().optional().describe('Ending index of the operation (for insert/delete)'),
        affectedCount: z.number().int().nonnegative().describe('Number of rows/columns affected by this operation'),
      })
    )
    .describe('Detailed results for each dimension operation'),
  updatedDimensions: z
    .object({
      rows: z.number().int().nonnegative().describe('Total number of rows in the sheet after all operations'),
      columns: z.number().int().nonnegative().describe('Total number of columns in the sheet after all operations'),
    })
    .describe('Final dimensions of the sheet after all operations'),
});

const outputSchema = z.discriminatedUnion('type', [successBranchSchema, AuthRequiredBranchSchema]);

const config = {
  description: 'Batch update sheet dimensions by inserting, deleting, or appending rows/columns. Operations are atomic (all succeed or all fail) and execute in optimal order automatically.',
  inputSchema,
  outputSchema: z.object({
    result: outputSchema,
  }),
} as const;

export type Input = z.infer<typeof inputSchema>;
export type Output = z.infer<typeof outputSchema>;

async function handler({ id, gid, requests }: Input, extra: EnrichedExtra): Promise<CallToolResult> {
  const logger = extra.logger;
  logger.debug?.('sheets.dimensions.batchUpdate called', {
    id,
    gid,
    requestCount: requests.length,
    operations: requests.map((r) => ({ operation: r.operation, dimension: r.dimension, startIndex: r.startIndex, endIndex: r.endIndex })),
  });

  try {
    const sheets = google.sheets({ version: 'v4', auth: extra.authContext.auth });

    // Get spreadsheet and sheet info in single API call
    const spreadsheetResponse = await sheets.spreadsheets.get({
      spreadsheetId: id,
      fields: 'properties.title,spreadsheetUrl,sheets.properties.sheetId,sheets.properties.title,sheets.properties.gridProperties',
    });

    const spreadsheetData = spreadsheetResponse.data;
    const spreadsheetTitle = spreadsheetData.properties?.title ?? '';
    const spreadsheetUrl = spreadsheetData.spreadsheetUrl ?? '';

    // Find sheet by gid
    const sheet = spreadsheetData.sheets?.find((s) => String(s.properties?.sheetId) === gid);
    if (!sheet?.properties) {
      logger.warn?.('Sheet not found for dimensions batch update', { id, gid, requestCount: requests.length });
      throw new McpError(ErrorCode.InvalidParams, `Sheet not found: ${gid}`);
    }

    const sheetTitle = sheet.properties.title ?? gid;
    const sheetId = sheet.properties.sheetId;

    if (sheetId === undefined || sheetId === null) {
      logger.error?.('Sheet ID not available for dimensions batch update', { id, gid, sheetTitle });
      throw new McpError(ErrorCode.InternalError, `Sheet ID not available for ${gid}. Cannot perform dimension operations without valid sheet ID.`);
    }

    const sheetUrl = `https://docs.google.com/spreadsheets/d/${id}/edit#gid=${sheetId}`;

    // Get current sheet dimensions for response calculation
    // Note: Google Sheets API may not always provide gridProperties for older sheets
    // Fall back to Google's documented defaults for new sheets
    const currentRowCount = sheet.properties.gridProperties?.rowCount ?? DEFAULT_ROW_COUNT;
    const currentColumnCount = sheet.properties.gridProperties?.columnCount ?? DEFAULT_COLUMN_COUNT;

    // Sort operations for optimal execution order to prevent index conflicts
    // Delete operations are processed first (high to low index) to avoid shifting issues
    // Insert/append operations are processed after (low to high index)
    const sortedRequests = sortOperations(requests as DimensionRequest[]);

    // Build Google Sheets API batch update requests
    const batchRequests = sortedRequests.map((operation) => buildDimensionRequest(operation, sheetId));

    logger.debug?.('sheets.dimensions.batchUpdate executing batch request', {
      spreadsheetId: id,
      sheetTitle,
      sheetId,
      totalOperations: batchRequests.length,
      operationTypes: sortedRequests.map((r) => r.operation),
      operationDetails: sortedRequests.map((r, i) => ({
        index: i,
        operation: r.operation,
        dimension: r.dimension,
        startIndex: r.startIndex,
        endIndex: r.endIndex,
        affectedCount: calculateAffectedCount(r),
      })),
      currentDimensions: { rows: currentRowCount, columns: currentColumnCount },
    });

    // Execute the atomic batch update
    const batchUpdateResponse = await sheets.spreadsheets.batchUpdate({
      spreadsheetId: id,
      requestBody: {
        requests: batchRequests,
        includeSpreadsheetInResponse: false, // We don't need the full spreadsheet data back
      },
    });

    const updateResult = batchUpdateResponse.data;

    // Comprehensive validation of batch update results
    if (!updateResult) {
      logger.error?.('Dimensions batch update failed - no response data', {
        spreadsheetId: id,
        sheetTitle,
        requestCount: requests.length,
      });
      throw new McpError(ErrorCode.InternalError, 'Batch update failed: no response data received from Google Sheets API');
    }

    const replies = updateResult.replies || [];
    const expectedCount = requests.length;
    const actualCount = replies.length;

    // Validate operation count matches expectations
    if (actualCount !== expectedCount) {
      logger.error?.('Dimensions batch update failed - operation count mismatch', {
        expectedOperations: expectedCount,
        actualReplies: actualCount,
        spreadsheetId: id,
        sheetTitle,
        receivedReplies: replies.map((reply, index) => ({
          index,
          replyType: Object.keys(reply || {})[0] || 'empty',
        })),
      });

      throw new McpError(ErrorCode.InternalError, `Batch operation failed: expected ${expectedCount} operations, received ${actualCount} replies. This may indicate a partial failure or Google API issue.`);
    }

    // Validate each reply exists - Google Sheets API may return empty objects for successful operations
    for (let i = 0; i < replies.length; i++) {
      const reply = replies[i];
      const request = sortedRequests[i];

      if (!request) {
        logger.error?.('Dimensions batch update failed - missing request', {
          replyIndex: i,
          hasReply: !!reply,
          hasRequest: !!request,
          spreadsheetId: id,
          sheetTitle,
        });
        throw new McpError(ErrorCode.InternalError, `Operation ${i} failed: missing request data`);
      }

      // Note: Google Sheets API often returns empty objects {} for successful dimension operations
      // This is normal behavior and indicates success, not failure
      // We validate that the reply exists (even if empty) rather than checking specific keys
      if (reply === null || reply === undefined) {
        logger.error?.('Dimensions batch update failed - null reply', {
          operationIndex: i,
          expectedOperation: request.operation,
          spreadsheetId: id,
          sheetTitle,
        });
        throw new McpError(ErrorCode.InternalError, `Operation ${i} (${request.operation}) failed: null reply from Google Sheets API`);
      }
    }

    // Calculate final dimensions and operation results
    // Note: We calculate based on the operations we performed, but actual dimensions
    // may vary slightly due to Google Sheets internal behavior (e.g., minimum dimensions)
    let finalRowCount = currentRowCount;
    let finalColumnCount = currentColumnCount;

    const operationResults = sortedRequests.map((operation, _index) => {
      const affectedCount = calculateAffectedCount(operation);

      // Update dimension counts based on operation
      // Operations are applied in sorted order, so this should match actual result
      if (operation.dimension === 'ROWS') {
        if (operation.operation === 'insertDimension') {
          finalRowCount += affectedCount;
        } else if (operation.operation === 'appendDimension') {
          finalRowCount += affectedCount;
        } else if (operation.operation === 'deleteDimension') {
          finalRowCount = Math.max(1, finalRowCount - affectedCount); // Google Sheets minimum 1 row
        }
      } else if (operation.dimension === 'COLUMNS') {
        if (operation.operation === 'insertDimension') {
          finalColumnCount += affectedCount;
        } else if (operation.operation === 'appendDimension') {
          finalColumnCount += affectedCount;
        } else if (operation.operation === 'deleteDimension') {
          finalColumnCount = Math.max(1, finalColumnCount - affectedCount); // Google Sheets minimum 1 column
        }
      }

      return {
        operation: operation.operation,
        dimension: operation.dimension,
        startIndex: operation.startIndex,
        endIndex: operation.endIndex,
        affectedCount,
      };
    });

    // Validate final dimensions are within Google Sheets limits
    if (finalRowCount > MAX_ROW_COUNT) {
      logger.warn?.('Final row count exceeds Google Sheets maximum', {
        finalRowCount,
        maxRowCount: MAX_ROW_COUNT,
        spreadsheetId: id,
        sheetTitle,
      });
    }
    if (finalColumnCount > MAX_COLUMN_COUNT) {
      logger.warn?.('Final column count exceeds Google Sheets maximum', {
        finalColumnCount,
        maxColumnCount: MAX_COLUMN_COUNT,
        spreadsheetId: id,
        sheetTitle,
      });
    }

    logger.debug?.('sheets.dimensions.batchUpdate completed successfully', {
      totalOperations: requests.length,
      finalRowCount,
      finalColumnCount,
      operationResults: operationResults.length,
    });

    const result: Output = {
      type: 'success' as const,
      id,
      gid: String(sheetId),
      spreadsheetTitle: spreadsheetTitle || '',
      spreadsheetUrl: spreadsheetUrl || '',
      sheetTitle,
      sheetUrl,
      totalOperations: requests.length,
      operationResults,
      updatedDimensions: {
        rows: finalRowCount,
        columns: finalColumnCount,
      },
    };

    return {
      content: [{ type: 'text' as const, text: JSON.stringify(result) }],
      structuredContent: { result },
    };
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    logger.error?.('sheets.dimensions.batchUpdate error', { error: message });
    throw new McpError(ErrorCode.InternalError, `Error batch updating dimensions: ${message}`, {
      stack: error instanceof Error ? error.stack : undefined,
    });
  }
}

export default function createTool() {
  return {
    name: 'dimensions-batch-update',
    config,
    handler,
  } satisfies ToolModule;
}
