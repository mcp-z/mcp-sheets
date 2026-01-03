/** Import CSV data to Google Sheets with database-style row append and deduplication */

import type { EnrichedExtra } from '@mcp-z/oauth-google';
import { schemas } from '@mcp-z/oauth-google';

const { AuthRequiredBranchSchema } = schemas;

import type { ToolModule } from '@mcp-z/server';
import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js';
import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { parse } from 'csv-parse';
import { google } from 'googleapis';
import { z } from 'zod';
import { SheetGidOutput, SheetGidSchema, SpreadsheetIdOutput, SpreadsheetIdSchema } from '../../schemas/index.ts';
import { getCsvReadStream } from '../../spreadsheet/csv-streaming.ts';
import { buildDeduplicationKey } from '../../spreadsheet/deduplication-utils.ts';
import { ensureTabAndHeaders } from '../../spreadsheet/sheet-operations.ts';

// Header mapping schema: source/target can be string (name) or number (0-based index)
const HeaderMapItemSchema = z.object({
  source: z.union([z.string(), z.number().int().min(0)]).describe('CSV column: header name (string) or 0-based index (number)'),
  target: z.union([z.string(), z.number().int().min(0)]).describe('Sheet column: header name (string) or 0-based index (number)'),
});

const inputSchema = z.object({
  id: SpreadsheetIdSchema,
  gid: SheetGidSchema,
  sourceUri: z.string().trim().min(1).describe('CSV file URI (file://, http://, https://)'),
  sourceHasHeaders: z.boolean().default(true).describe('Source has header row for column name mapping. Set to false for data-only sources (numeric indices required).'),
  headerMap: z.array(HeaderMapItemSchema).describe('Column mappings from CSV to sheet'),
  deduplicateBy: z
    .array(z.union([z.string(), z.number().int().min(0)]))
    .optional()
    .describe('Sheet columns for deduplication (names or indices)'),
});

const successBranchSchema = z.object({
  type: z.literal('success'),
  id: SpreadsheetIdOutput,
  gid: SheetGidOutput,
  sheetTitle: z.string().describe('Sheet tab name'),
  updatedRows: z.number().describe('Number of rows appended'),
  rowsSkipped: z.number().optional().describe('Number of duplicate rows skipped'),
  sheetUrl: z.string().optional().describe('URL to view the sheet'),
});

const outputSchema = z.discriminatedUnion('type', [successBranchSchema, AuthRequiredBranchSchema]);

const config = {
  description: 'Import CSV to Google Sheets with column mapping and optional deduplication. Streams data for large files.',
  inputSchema,
  outputSchema: z.object({
    result: outputSchema,
  }),
} as const;

export type Input = z.infer<typeof inputSchema>;
export type Output = z.infer<typeof outputSchema>;

/** Batch size for Sheets API calls (1000 rows × 12 cols = 12K cells, well under 40K limit) */
const BATCH_SIZE = 1000;

/**
 * Resolve column reference to numeric index
 * @param ref Column reference (string name or number index)
 * @param headers Header row (null when sourceHasHeaders=false)
 * @param sourceHasHeaders Whether headers are present
 * @returns 0-based column index
 */
function resolveColumnReference(ref: string | number, headers: string[] | null, sourceHasHeaders: boolean, context: string): number {
  // If number, use directly as 0-based index
  if (typeof ref === 'number') {
    if (ref < 0) {
      throw new Error(`${context}: Column index must be >= 0, got ${ref}`);
    }
    return ref;
  }

  // If string, must be header name
  if (!sourceHasHeaders || !headers) {
    throw new Error(`${context}: String column reference "${ref}" requires sourceHasHeaders=true. Use numeric index when sourceHasHeaders=false.`);
  }

  const index = headers.indexOf(ref);
  if (index === -1) {
    throw new Error(`${context}: Header "${ref}" not found in [${headers.join(', ')}]`);
  }

  return index;
}

async function handler({ id, gid, sourceUri, sourceHasHeaders, headerMap, deduplicateBy }: Input, extra: EnrichedExtra): Promise<CallToolResult> {
  const logger = extra.logger;
  logger.info('sheets.rows.csv-append called', { id, gid, sourceUri, sourceHasHeaders, headerMap, deduplicateBy });

  try {
    if (headerMap.length === 0) {
      throw new McpError(ErrorCode.InvalidParams, 'headerMap cannot be empty');
    }

    // Validate: if sourceHasHeaders=false, all references must be numeric
    if (!sourceHasHeaders) {
      for (const { source, target } of headerMap) {
        if (typeof source === 'string') {
          throw new McpError(ErrorCode.InvalidParams, `sourceHasHeaders=false requires numeric indices. Got string source: "${source}"`);
        }
        if (typeof target === 'string') {
          throw new McpError(ErrorCode.InvalidParams, `sourceHasHeaders=false requires numeric indices. Got string target: "${target}"`);
        }
      }
      if (deduplicateBy) {
        for (const colRef of deduplicateBy) {
          if (typeof colRef === 'string') {
            throw new McpError(ErrorCode.InvalidParams, `sourceHasHeaders=false requires numeric indices in deduplicateBy. Got string: "${colRef}"`);
          }
        }
      }
    }

    const sheets = google.sheets({ version: 'v4', auth: extra.authContext.auth });

    // Get sheet details using the gid
    const spreadsheetResponse = await sheets.spreadsheets.get({
      spreadsheetId: id,
      fields: 'sheets.properties.sheetId,sheets.properties.title',
    });

    const sheet = spreadsheetResponse.data.sheets?.find((s) => String(s.properties?.sheetId) === gid);

    if (!sheet) {
      throw new McpError(ErrorCode.InvalidParams, 'Sheet not found');
    }

    const sheetTitle = sheet?.properties?.title ?? '';

    // Determine target headers (for sourceHasHeaders=true) or column count (for sourceHasHeaders=false)
    let sheetHeaders: string[] | null = null;
    let existingKeySet: Set<string> = new Set();
    const keyColumns: (string | number)[] = deduplicateBy || [];

    if (sourceHasHeaders) {
      // Extract target header names from headerMap
      const targetHeaderNames = headerMap.map(({ target }) => target).filter((t) => typeof t === 'string') as string[];

      // Use ensureTabAndHeaders to setup headers and fetch existing keys
      const headerResult = await ensureTabAndHeaders(sheets, {
        spreadsheetId: id,
        sheetTitle,
        requiredHeader: targetHeaderNames.length > 0 ? targetHeaderNames : null,
        keyColumns: keyColumns.filter((k) => typeof k === 'string') as string[],
        logger,
      });

      sheetHeaders = headerResult.header;
      existingKeySet = headerResult.keySet;
    } else {
      // sourceHasHeaders=false: Read existing data for deduplication (if needed)
      if (deduplicateBy && deduplicateBy.length > 0) {
        // Read data in chunks for memory efficiency with large sheets
        const CHUNK_SIZE = 1000;
        let startRow = 1;
        let hasMore = true;

        while (hasMore) {
          const endRow = startRow + CHUNK_SIZE - 1;
          const chunkRange = `${sheetTitle}!A${startRow}:ZZZ${endRow}`;

          const response = await sheets.spreadsheets.values.get({
            spreadsheetId: id,
            range: chunkRange,
          });

          const rows = response.data.values || [];

          for (const row of rows) {
            const key = buildDeduplicationKey(row, keyColumns, null, false);
            if (key.replace(/::/g, '') !== '') {
              existingKeySet.add(key);
            }
          }

          // Check if there are more rows to read
          if (rows.length < CHUNK_SIZE) {
            hasMore = false;
          } else {
            startRow += CHUNK_SIZE;
          }
        }

        logger.info('sheets.rows.csv-append existing keys loaded', { keyCount: existingKeySet.size });
      }
    }

    // Streaming CSV processing state
    let sourceHeaders: string[] | null = null;
    let batch: (string | number | boolean | null)[][] = [];
    let totalRows = 0;
    let rowsSkipped = 0;

    // Get readable stream from CSV URI (no temp files!)
    const readStream = await getCsvReadStream(sourceUri);

    // Create CSV parser
    const parser = readStream.pipe(
      parse({
        columns: !!sourceHasHeaders, // Parse first row as column names if source has headers
        skip_empty_lines: true,
        trim: true,
        cast: true, // Auto-convert numbers/booleans
        relax_column_count: true,
      })
    );

    // Helper to append batch to Sheets
    const appendBatch = async (rows: (string | number | boolean | null)[][]): Promise<void> => {
      if (rows.length === 0) return;

      await sheets.spreadsheets.values.append({
        spreadsheetId: id,
        range: `${sheetTitle}!A:A`,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values: rows, majorDimension: 'ROWS' },
      });

      logger.info('sheets.rows.csv-append batch appended', { batchSize: rows.length, totalRows });
    };

    // Resolve headerMap to numeric indices (do this after extracting CSV headers)
    let resolvedMap: Array<{ sourceIdx: number; targetIdx: number }> = [];

    // Stream and process records
    for await (const record of parser) {
      if (sourceHasHeaders) {
        // Extract source headers from first record
        if (sourceHeaders === null) {
          sourceHeaders = Object.keys(record as Record<string, unknown>);
          logger.info('sheets.rows.csv-append source headers', { sourceHeaders });

          // Resolve headerMap now that we have both source and sheet headers
          resolvedMap = headerMap.map(({ source, target }) => ({
            sourceIdx: resolveColumnReference(source, sourceHeaders, sourceHasHeaders, 'headerMap.source'),
            targetIdx: resolveColumnReference(target, sheetHeaders, sourceHasHeaders, 'headerMap.target'),
          }));
        }

        // Map source record to sheet row
        const sourceRow = sourceHeaders?.map((h) => (record as Record<string, unknown>)[h] ?? null);
        const maxTargetIdx = Math.max(...resolvedMap.map((m) => m.targetIdx));
        const sheetRow = new Array(maxTargetIdx + 1).fill(null);

        for (const { sourceIdx, targetIdx } of resolvedMap) {
          if (sourceRow) {
            sheetRow[targetIdx] = sourceRow[sourceIdx] ?? null;
          }
        }

        // Check deduplication
        if (deduplicateBy && deduplicateBy.length > 0) {
          const key = buildDeduplicationKey(sheetRow, keyColumns, sheetHeaders, sourceHasHeaders);
          if (existingKeySet.has(key)) {
            rowsSkipped++;
            continue; // Skip duplicate
          }
          existingKeySet.add(key); // Add to set for future deduplication
        }

        // Add to batch
        batch.push(sheetRow);
        totalRows++;
      } else {
        // sourceHasHeaders=false: record is an array (not an object)
        // Resolve map on first row
        if (resolvedMap.length === 0) {
          resolvedMap = headerMap.map(({ source, target }) => ({
            sourceIdx: resolveColumnReference(source, null, sourceHasHeaders, 'headerMap.source'),
            targetIdx: resolveColumnReference(target, null, sourceHasHeaders, 'headerMap.target'),
          }));
        }

        const sourceRow = record as unknown[];
        const maxTargetIdx = Math.max(...resolvedMap.map((m) => m.targetIdx));
        const sheetRow = new Array(maxTargetIdx + 1).fill(null);

        for (const { sourceIdx, targetIdx } of resolvedMap) {
          if (sourceIdx < sourceRow.length) {
            sheetRow[targetIdx] = sourceRow[sourceIdx] ?? null;
          }
        }

        // Check deduplication
        if (deduplicateBy && deduplicateBy.length > 0) {
          const key = buildDeduplicationKey(sheetRow, keyColumns, null, sourceHasHeaders);
          if (existingKeySet.has(key)) {
            rowsSkipped++;
            continue; // Skip duplicate
          }
          existingKeySet.add(key);
        }

        // Add to batch
        batch.push(sheetRow);
        totalRows++;
      }

      // Append batch when full
      if (batch.length >= BATCH_SIZE) {
        await appendBatch(batch);
        batch = []; // Clear batch
      }
    }

    // Flush remaining rows
    if (batch.length > 0) {
      await appendBatch(batch);
    }

    logger.info('sheets.rows.csv-append streaming complete', { totalRows, rowsSkipped });

    const updatedRows = totalRows;

    logger.info('sheets.rows.csv-append completed', { id, gid, updatedRows, rowsSkipped, sourceUri });

    const result: Output = {
      type: 'success' as const,
      id,
      gid,
      sheetTitle,
      updatedRows,
      rowsSkipped,
      sheetUrl: `https://docs.google.com/spreadsheets/d/${id}/edit#gid=${gid}`,
    };

    return {
      content: [{ type: 'text' as const, text: JSON.stringify(result) }],
      structuredContent: { result },
    };
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    logger.error('sheets.rows.csv-append error', { error: message });

    throw new McpError(ErrorCode.InternalError, `Error appending CSV rows: ${message}`, {
      stack: error instanceof Error ? error.stack : undefined,
    });
  }
}

export default function createTool() {
  return {
    name: 'rows-csv-append',
    config,
    handler,
  } satisfies ToolModule;
}
