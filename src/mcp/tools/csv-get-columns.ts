/** Get column names from CSV file (peek at first row only) */

import type { EnrichedExtra } from '@mcp-z/oauth-google';
import { schemas } from '@mcp-z/oauth-google';

const { AuthRequiredBranchSchema } = schemas;

import type { ToolModule } from '@mcp-z/server';
import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js';
import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { parse } from 'csv-parse';
import { z } from 'zod';
import { getCsvReadStream } from '../../spreadsheet/csv-streaming.js';

const inputSchema = z.object({
  sourceUri: z.string().trim().min(1).describe('CSV file URI (file://, http://, https://)'),
});

// Success branch schema - uses columns: for consistency with standard vocabulary
const successBranchSchema = z.object({
  type: z.literal('success'),
  columns: z.array(z.string()).describe('First row values (column names) or empty if no rows'),
  isEmpty: z.boolean().describe('True if CSV has zero rows'),
});

const outputSchema = z.discriminatedUnion('type', [successBranchSchema, AuthRequiredBranchSchema]);

const config = {
  description: 'Get first row from CSV file (streaming, no memory overhead). Returns columns array and isEmpty flag.',
  inputSchema,
  outputSchema: z.object({
    result: outputSchema,
  }),
} as const;

export type Input = z.infer<typeof inputSchema>;
export type Output = z.infer<typeof outputSchema>;

async function handler({ sourceUri }: Input, extra: EnrichedExtra): Promise<CallToolResult> {
  const logger = extra.logger;
  logger.debug?.('sheets.csv.get-columns called', { sourceUri });

  try {
    // Get readable stream from CSV URI (no temp files!)
    const readStream = await getCsvReadStream(sourceUri);

    // Create CSV parser without treating first row as column names
    // We just want to read the raw first row
    const parser = readStream.pipe(
      parse({
        columns: false, // Don't treat first row as headers
        skip_empty_lines: true,
        trim: true,
        cast: true, // Auto-convert numbers/booleans
        relax_column_count: true,
      })
    );

    // Read only the first row
    let firstRow: unknown[] = [];
    let rowCount = 0;

    for await (const row of parser) {
      firstRow = row;
      rowCount++;
      break; // Only read first row
    }

    // Convert first row to strings (column names)
    const columns = firstRow.map((value) => String(value ?? ''));
    const isEmpty = rowCount === 0;

    const result: Output = {
      type: 'success' as const,
      columns,
      isEmpty,
    };

    logger.info?.('sheets.csv.get-columns completed', { sourceUri, columnCount: columns.length, isEmpty });
    return {
      content: [{ type: 'text' as const, text: JSON.stringify(result) }],
      structuredContent: { result },
    };
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    logger.error?.('sheets.csv.get-columns error', { error: message });
    throw new McpError(ErrorCode.InternalError, `Error getting CSV columns: ${message}`, {
      stack: error instanceof Error ? error.stack : undefined,
    });
  }
}

export default function createTool() {
  return {
    name: 'csv-get-columns',
    config,
    handler,
  } satisfies ToolModule;
}
