import type { EnrichedExtra } from '@mcp-z/oauth-google';
import { schemas } from '@mcp-z/oauth-google';

const { AuthRequiredBranchSchema } = schemas;

import type { ToolModule } from '@mcp-z/server';
import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js';
import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { google } from 'googleapis';
import { z } from 'zod';
import { SheetCellSchema, SheetGidOutput, SheetGidSchema, SpreadsheetIdOutput, SpreadsheetIdSchema } from '../../schemas/index.ts';
import { generateRowKey, type KeyGenerationStrategy, type Row, snapshotHeaderKeysAndPositions, type UpsertOptions, upsertByKey } from '../../spreadsheet/data-operations.ts';

// Input schema for columns update requests with enhanced validation
const inputSchema = z
  .object({
    id: SpreadsheetIdSchema,
    gid: SheetGidSchema,
    rows: z.array(z.array(SheetCellSchema)).min(1).max(1000).describe('Array of rows to upsert (max 1000 rows per request). Each row is an array of cell values matching the headers array'),
    headers: z.array(z.string().min(1).max(100)).min(1).max(50).describe('Array of column names/headers (max 50 columns). Must match the length of each row in the rows array'),
    updateBy: z.array(z.string().min(1).max(100)).min(1).max(10).describe('Array of column names to use as unique keys for matching existing rows (max 10 key columns). These columns must exist in the headers array'),
    behavior: z
      .enum(['add-or-update', 'update-only', 'add-only'])
      .default('add-or-update')
      .describe('Update behavior: add-or-update (default) adds new rows and updates existing, update-only skips new rows, add-only skips existing rows. BEST FOR: Bulk upsert operations with key matching in structured database or table contexts.'),
    valueInputOption: z.enum(['RAW', 'USER_ENTERED']).default('USER_ENTERED').describe('How input data should be interpreted (RAW = exact values, USER_ENTERED = parsed like user input with formulas, dates, etc.)'),
  })
  .refine(
    (data) => {
      // Validate that all updateBy columns exist in headers
      const missingColumns = data.updateBy.filter((col) => !data.headers.includes(col));
      return missingColumns.length === 0;
    },
    {
      message: 'All updateBy columns must exist in the headers array',
      path: ['updateBy'],
    }
  )
  .refine(
    (data) => {
      // Validate that all rows have the same length as headers
      const invalidRows = data.rows.findIndex((row) => row.length !== data.headers.length);
      return invalidRows === -1;
    },
    {
      message: 'All rows must have the same length as the headers array',
      path: ['rows'],
    }
  )
  .refine(
    (data) => {
      // Validate that updateBy columns are unique
      const uniqueColumns = new Set(data.updateBy);
      return uniqueColumns.size === data.updateBy.length;
    },
    {
      message: 'updateBy columns must be unique',
      path: ['updateBy'],
    }
  );

// Success branch schema
const successBranchSchema = z.object({
  type: z.literal('success'),
  id: SpreadsheetIdOutput,
  gid: SheetGidOutput,
  spreadsheetTitle: z.string().describe('Title of the updated spreadsheet'),
  spreadsheetUrl: z.string().describe('URL of the updated spreadsheet'),
  sheetTitle: z.string().describe('Title of the updated sheet'),
  sheetUrl: z.string().describe('URL of the updated sheet'),
  updatedRows: z.number().int().nonnegative().describe('Number of rows that were successfully updated or inserted'),
  insertedKeys: z.array(z.string()).describe('Keys of rows that were successfully inserted (new rows)'),
  rowsSkipped: z.number().int().nonnegative().describe('Number of rows that were skipped based on the update behavior and existing data'),
  headersAdded: z.array(z.string()).describe('Column headers that were added to the sheet (if any were missing)'),
  errors: z.array(z.string()).optional().describe('Any non-fatal errors encountered during the operation'),
});

const outputSchema = z.discriminatedUnion('type', [successBranchSchema, AuthRequiredBranchSchema]);

const config = {
  description: 'Update spreadsheet data by column headers with intelligent upsert logic. Supports adding missing columns, flexible update behaviors, and robust error handling. Uses column names as keys for matching existing rows, enabling context-aware data synchronization workflows.',
  inputSchema,
  outputSchema: z.object({
    result: outputSchema,
  }),
} as const;

export type Input = z.infer<typeof inputSchema>;
export type Output = z.infer<typeof outputSchema>;

// Helper function to ensure consistent key generation
function generateConsistentRowKey(row: Row, headers: string[], strategy: KeyGenerationStrategy): string {
  return generateRowKey(row, headers, strategy);
}

async function handler({ id, gid, rows, headers, updateBy, behavior = 'add-or-update', valueInputOption = 'USER_ENTERED' }: Input, extra: EnrichedExtra): Promise<CallToolResult> {
  const logger = extra.logger;
  logger.debug?.('sheets.columns.update called', {
    id,
    gid,
    rowCount: rows.length,
    headerCount: headers.length,
    updateByColumns: updateBy,
    behavior,
    valueInputOption,
  });

  try {
    const sheets = google.sheets({ version: 'v4', auth: extra.authContext.auth });

    // Note: Basic input validation is now handled by the enhanced Zod schema

    // EARLY VALIDATION: Validate that updateBy columns have non-empty values for all rows
    // This validation happens before finding spreadsheet/sheet to fail fast
    const emptyKeyErrors: string[] = [];
    rows.forEach((row, rowIndex) => {
      const keyValues = updateBy.map((keyCol) => {
        const colIndex = headers.indexOf(keyCol);
        return colIndex >= 0 ? String(row[colIndex] ?? '').trim() : '';
      });

      if (keyValues.some((val) => val === '')) {
        emptyKeyErrors.push(`Row ${rowIndex + 1} has empty key values for columns: ${updateBy.join(', ')}`);
      }
    });

    if (emptyKeyErrors.length > 0) {
      const message = `Silent data loss prevented - empty key columns detected:\n${emptyKeyErrors.join('\n')}`;
      logger.error?.('sheets.columns.update error', { error: message });
      throw new McpError(ErrorCode.InvalidParams, message);
    }

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
      logger.warn?.('Sheet not found for columns update', { id, gid, rowCount: rows.length });
      throw new McpError(ErrorCode.InvalidParams, `Sheet not found: ${gid}`);
    }

    const sheetTitle = sheet.properties.title ?? gid;
    const sheetGid = sheet.properties.sheetId ?? 0;
    const sheetUrl = `https://docs.google.com/spreadsheets/d/${id}/edit#gid=${sheetGid}`;

    // Configure upsert options based on behavior
    const upsertOptions: UpsertOptions = {
      keyStrategy: {
        keyColumns: updateBy,
        useProviderIdLogic: false, // Use standard key generation for column-based updates
        separator: '\\',
      },
      allowUpdates: behavior === 'add-or-update' || behavior === 'update-only',
      batchSize: 50,
      valueInputOption: valueInputOption,
    };

    // Get existing keys to implement behavior filtering
    const { keySet: existingKeys } = await snapshotHeaderKeysAndPositions(sheets, id, sheetTitle, updateBy, upsertOptions.keyStrategy);

    // Filter rows based on behavior using consistent key generation
    let processedRows = rows;
    if (behavior === 'add-only') {
      // Only include rows that don't exist yet
      processedRows = rows.filter((row) => {
        // Use the same key generation strategy as the upsert function
        const key = generateConsistentRowKey(row, headers, upsertOptions.keyStrategy);
        return key !== '' && !existingKeys.has(key);
      });
      logger.debug?.('Filtered rows for add-only behavior', {
        originalCount: rows.length,
        filteredCount: processedRows.length,
      });
    } else if (behavior === 'update-only') {
      // Only include rows that already exist
      processedRows = rows.filter((row) => {
        // Use the same key generation strategy as the upsert function
        const key = generateConsistentRowKey(row, headers, upsertOptions.keyStrategy);
        return key !== '' && existingKeys.has(key);
      });
      logger.debug?.('Filtered rows for update-only behavior', {
        originalCount: rows.length,
        filteredCount: processedRows.length,
      });
    }

    // Early return if no rows to process after behavior filtering
    if (processedRows.length === 0) {
      logger.info?.('No rows to process after behavior filtering', { behavior, originalCount: rows.length });
      const result: Output = {
        type: 'success' as const,
        id,
        gid: String(sheetGid),
        spreadsheetTitle: spreadsheetTitle || '',
        spreadsheetUrl: spreadsheetUrl || '',
        sheetTitle,
        sheetUrl,
        updatedRows: 0,
        insertedKeys: [],
        rowsSkipped: rows.length,
        headersAdded: [],
        errors: [`No rows matched behavior '${behavior}' - all ${rows.length} rows were skipped`],
      };
      return {
        content: [{ type: 'text' as const, text: JSON.stringify(result) }],
        structuredContent: { result },
      };
    }

    logger.debug?.('sheets.columns.update executing upsert operation', {
      spreadsheetId: id,
      sheetTitle,
      canonicalHeaders: headers,
      keyColumns: updateBy,
      behavior,
      processedRowsCount: processedRows.length,
    });

    // Execute the upsert operation using the shared function
    const upsertResult = await upsertByKey(sheets, {
      spreadsheetId: id,
      sheetTitle,
      rows: processedRows,
      canonicalHeaders: headers,
      options: upsertOptions,
      logger,
    });

    // Track headers that were added (if any)
    // This would require additional logic in upsertByKey to return added headers
    const headersAdded: string[] = [];

    // Calculate total rows skipped: behavior filter + upsert duplicates
    const rowsFilteredByBehavior = rows.length - processedRows.length;
    const totalRowsSkipped = upsertResult.rowsSkipped + rowsFilteredByBehavior;

    logger.debug?.('sheets.columns.update completed successfully', {
      updatedRows: upsertResult.updatedRows,
      insertedCount: upsertResult.inserted.length,
      rowsSkipped: totalRowsSkipped,
      rowsFilteredByBehavior,
      rowsSkippedByUpsert: upsertResult.rowsSkipped,
      errorsCount: upsertResult.errors?.length || 0,
    });

    const result: Output = {
      type: 'success' as const,
      id,
      gid: String(sheetGid),
      spreadsheetTitle: spreadsheetTitle || '',
      spreadsheetUrl: spreadsheetUrl || '',
      sheetTitle,
      sheetUrl,
      updatedRows: upsertResult.updatedRows,
      insertedKeys: upsertResult.inserted,
      rowsSkipped: totalRowsSkipped,
      headersAdded,
      errors: upsertResult.errors,
    };

    return {
      content: [{ type: 'text' as const, text: JSON.stringify(result) }],
      structuredContent: { result },
    };
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    logger.error?.('sheets.columns.update error', { error: message });
    throw new McpError(ErrorCode.InternalError, `Error updating columns: ${message}`, {
      stack: error instanceof Error ? error.stack : undefined,
    });
  }
}

export default function createTool() {
  return {
    name: 'columns-update',
    config,
    handler,
  } satisfies ToolModule;
}
