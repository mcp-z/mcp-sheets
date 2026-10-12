/** Get column names from CSV file (peek at first row only) */

import type { EnrichedExtra } from '@mcp-z/oauth-google';
import { schemas } from '@mcp-z/oauth-google';

const { AuthRequiredBranchSchema } = schemas;

import type { CallToolResult, ToolModule } from '@mcp-z/server';
import { ProtocolError, ProtocolErrorCode } from '@mcp-z/server';
import { parse } from 'csv-parse';
import { pipeline } from 'stream/promises';
import { z } from 'zod';
import { getCsvReadStream } from '../../spreadsheet/csv-streaming.ts';

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
    const parser = parse({
      columns: false,
      skip_empty_lines: true,
      trim: true,
      cast: true,
      relax_column_count: true,
    });
    const readStream = await getCsvReadStream(sourceUri);
    const firstRowRead = new Error('First CSV row read');
    const completed = pipeline(readStream, parser).then(
      () => undefined,
      (error: unknown) => error
    );

    let firstRow: unknown[] = [];
    let rowCount = 0;
    try {
      const record = await parser[Symbol.asyncIterator]().next();
      if (!record.done) {
        firstRow = record.value;
        rowCount = 1;
        parser.destroy(firstRowRead);
      }
    } catch (error) {
      await completed;
      throw error;
    }

    const pipelineError = await completed;
    if (pipelineError !== undefined && pipelineError !== firstRowRead) throw pipelineError;

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
    throw new ProtocolError(ProtocolErrorCode.InternalError, `Error getting CSV columns: ${message}`, {
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
