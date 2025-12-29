import type { EnrichedExtra } from '@mcp-z/oauth-google';
import { schemas } from '@mcp-z/oauth-google';

const { AuthRequiredBranchSchema } = schemas;

import type { ToolModule } from '@mcp-z/server';
import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js';
import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { google } from 'googleapis';
import { z } from 'zod';
import { SheetGidSchema, SpreadsheetIdOutput, SpreadsheetIdSchema } from '../../schemas/index.js';

const MAX_BATCH_SIZE = 1000;

const inputSchema = z.object({
  id: SpreadsheetIdSchema,
  gids: z.array(SheetGidSchema).min(1).max(MAX_BATCH_SIZE).describe('Sheet grid IDs to permanently delete'),
});

const successBranchSchema = z.object({
  type: z.literal('success'),
  operationSummary: z.string().describe('Human-readable summary of the operation'),
  itemsProcessed: z.number().describe('Total sheets attempted to delete'),
  itemsChanged: z.number().describe('Number of sheets successfully deleted'),
  completedAt: z.string().describe('ISO timestamp when operation completed'),
  recoverable: z.literal(false).describe('Whether deletion can be undone (always false)'),
  id: SpreadsheetIdOutput,
  spreadsheetUrl: z.string().optional().describe('URL to view the spreadsheet'),
  failures: z
    .array(
      z.object({
        gid: z.string().describe('Grid ID of sheet that failed to delete'),
        error: z.string().describe('Error message explaining the failure'),
      })
    )
    .optional()
    .describe('Details of any sheets that failed to delete'),
});

const outputSchema = z.discriminatedUnion('type', [successBranchSchema, AuthRequiredBranchSchema]);

const config = {
  description: 'Permanently delete sheets from a spreadsheet. Cannot delete the last remaining sheet.',
  inputSchema,
  outputSchema: z.object({
    result: outputSchema,
  }),
} as const;

export type Input = z.infer<typeof inputSchema>;
export type Output = z.infer<typeof outputSchema>;

async function handler({ id, gids }: Input, extra: EnrichedExtra): Promise<CallToolResult> {
  const logger = extra.logger;
  logger.info('sheets.sheet.delete called', { id, count: gids.length });

  try {
    const sheets = google.sheets({ version: 'v4', auth: extra.authContext.auth });

    const results = await Promise.allSettled(
      gids.map(async (gid) => {
        await sheets.spreadsheets.batchUpdate({
          spreadsheetId: id,
          requestBody: {
            requests: [{ deleteSheet: { sheetId: Number(gid) } }],
          },
        });
        return gid;
      })
    );

    // Separate successes and failures
    const failures: Array<{ gid: string; error: string }> = [];

    results.forEach((result, index) => {
      const gid = gids[index];
      if (!gid) return;

      if (result.status === 'rejected') {
        const errorMessage = result.reason instanceof Error ? result.reason.message : String(result.reason);
        failures.push({ gid, error: errorMessage });
      }
    });

    const successCount = gids.length - failures.length;
    const failureCount = failures.length;
    const totalCount = gids.length;

    const summary = failureCount === 0 ? `Permanently deleted ${successCount} sheet${successCount === 1 ? '' : 's'}` : `Deleted ${successCount} of ${totalCount} sheet${totalCount === 1 ? '' : 's'} (${failureCount} failed)`;

    logger.info('sheets.sheet.delete completed', {
      totalCount,
      successCount,
      failureCount,
    });

    const result: Output = {
      type: 'success' as const,
      operationSummary: summary,
      itemsProcessed: totalCount,
      itemsChanged: successCount,
      completedAt: new Date().toISOString(),
      recoverable: false as const,
      id,
      spreadsheetUrl: `https://docs.google.com/spreadsheets/d/${id}`,
      ...(failures.length > 0 && { failures }),
    };

    return {
      content: [{ type: 'text' as const, text: JSON.stringify(result) }],
      structuredContent: { result },
    };
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    logger.error('sheets.sheet.delete error', { error: message });

    throw new McpError(ErrorCode.InternalError, `Error deleting sheet: ${message}`, {
      stack: error instanceof Error ? error.stack : undefined,
    });
  }
}

export default function createTool() {
  return {
    name: 'sheet-delete',
    config,
    handler,
  } satisfies ToolModule;
}
