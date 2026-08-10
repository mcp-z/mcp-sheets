import type { EnrichedExtra } from '@mcp-z/oauth-google';
import { schemas } from '@mcp-z/oauth-google';

const { AuthRequiredBranchSchema } = schemas;

import type { ToolModule } from '@mcp-z/server';
import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js';
import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { google, type sheets_v4 } from 'googleapis';
import { z } from 'zod';
import { SheetGidOutput, SheetGidSchema, SpreadsheetIdOutput, SpreadsheetIdSchema } from '../../schemas/index.ts';
import { buildTextFormatRuns, parseInlineMarkdown } from '../../spreadsheet/markdown-runs.ts';
import { parseA1Notation, rangeReferenceToGridRange } from '../../spreadsheet/range-operations.ts';

// Input schema for a single markdown-cell request
const MarkdownCellRequestSchema = z.object({
  cell: z.string().min(1).describe('Single-cell A1 reference (e.g., "I2"). Multi-cell ranges (e.g., "A1:B2") are rejected as a per-request failure.'),
  text: z.string().describe('Cell content as inline markdown: [label](url) links, bare URLs, **bold**, *italic*, ~~strikethrough~~. Block syntax and other unsupported markdown is written as literal text.'),
});

const inputSchema = z.object({
  id: SpreadsheetIdSchema,
  gid: SheetGidSchema,
  requests: z.array(MarkdownCellRequestSchema).min(1).max(50).describe('Array of cell markdown requests. Batch multiple cells for efficiency.'),
});

const successBranchSchema = z.object({
  type: z.literal('success'),
  id: SpreadsheetIdOutput,
  gid: SheetGidOutput,
  sheetTitle: z.string().describe('Title of the updated sheet'),
  sheetUrl: z.string().describe('URL of the updated sheet'),
  successCount: z.number().int().nonnegative().describe('Number of cells successfully written'),
  cells: z
    .array(
      z.object({
        cell: z.string().describe('A1 reference of the cell that was written'),
        text: z.string().describe('Resolved plain text written to the cell (markdown syntax stripped/literalized)'),
        linkCount: z.number().int().nonnegative().describe('Number of clickable links written into the cell'),
      })
    )
    .describe('Cells successfully written'),
  failedCells: z
    .array(
      z.object({
        cell: z.string().describe('A1 reference that failed'),
        error: z.string().describe('Why the cell failed (e.g., not a single-cell reference)'),
      })
    )
    .optional()
    .describe('Only populated if some cells failed'),
});

const outputSchema = z.discriminatedUnion('type', [successBranchSchema, AuthRequiredBranchSchema]);

const config = {
  description:
    'Replace cell content with text parsed from inline markdown. Supports [label](url) links, bare URLs, **bold**, *italic*, ~~strikethrough~~ — the only way to put MULTIPLE clickable links in one cell. Block syntax (headers, lists, tables) and other unsupported markdown is written as literal text. Writes text cells only — no formula or number coercion. For whole-cell formatting (colors, borders, number formats) use cells-format.',
  inputSchema,
  outputSchema: z.object({
    result: outputSchema,
  }),
} as const;

export type Input = z.infer<typeof inputSchema>;
export type Output = z.infer<typeof outputSchema>;

async function handler({ id, gid, requests }: Input, extra: EnrichedExtra): Promise<CallToolResult> {
  const logger = extra.logger;
  logger.info('sheets.values.markdownUpdate called', {
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
      logger.info('Sheet not found for markdown update', { id, gid, requestCount: requests.length });
      throw new McpError(ErrorCode.InvalidParams, `Sheet not found: ${gid}`);
    }

    const sheetTitle = sheet.properties.title ?? gid;
    // googleapis types sheetId as `number | null | undefined`; coerce for rangeReferenceToGridRange below (same pattern as values-batch-update.ts's sheetGid).
    const sheetId = sheet.properties.sheetId ?? 0;
    const sheetUrl = `https://docs.google.com/spreadsheets/d/${id}/edit#gid=${sheetId}`;

    // Build one updateCells request per cell
    const batchRequests: sheets_v4.Schema$Request[] = [];
    const cells: Array<{ cell: string; text: string; linkCount: number }> = [];
    const failedCells: Array<{ cell: string; error: string }> = [];

    for (const request of requests) {
      try {
        // Parse A1 notation and reject anything that isn't a single cell
        const rangeRef = parseA1Notation(request.cell);
        if (rangeRef.type !== 'cell') {
          throw new Error(`Not a single-cell reference: "${request.cell}"`);
        }

        const gridRange = rangeReferenceToGridRange(rangeRef, sheetId);
        const parsed = parseInlineMarkdown(request.text);
        const runs = buildTextFormatRuns(parsed);

        batchRequests.push({
          updateCells: {
            range: gridRange,
            rows: [
              {
                values: [
                  {
                    userEnteredValue: { stringValue: parsed.text },
                    textFormatRuns: runs,
                  },
                ],
              },
            ],
            fields: 'userEnteredValue,textFormatRuns',
          },
        });
        cells.push({ cell: request.cell, text: parsed.text, linkCount: parsed.linkCount });
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        logger.info('Failed to parse cell for markdown update', {
          cell: request.cell,
          error: message,
        });
        failedCells.push({
          cell: request.cell,
          error: `Failed to parse cell: ${message}`,
        });
      }
    }

    // Early return if all cells failed
    if (batchRequests.length === 0) {
      const result: Output = {
        type: 'success' as const,
        id,
        gid: String(sheetId),
        sheetTitle,
        sheetUrl,
        successCount: 0,
        cells: [],
        failedCells: failedCells.length > 0 ? failedCells : undefined,
      };

      return {
        content: [{ type: 'text' as const, text: JSON.stringify(result) }],
        structuredContent: { result },
      };
    }

    logger.info('sheets.values.markdownUpdate executing batch request', {
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

    logger.info('sheets.values.markdownUpdate completed successfully', {
      successCount: cells.length,
      failedCount: failedCells.length,
    });

    const result: Output = {
      type: 'success' as const,
      id,
      gid: String(sheetId),
      sheetTitle,
      sheetUrl,
      successCount: cells.length,
      cells,
      failedCells: failedCells.length > 0 ? failedCells : undefined,
    };

    return {
      content: [{ type: 'text' as const, text: JSON.stringify(result) }],
      structuredContent: { result },
    };
  } catch (error) {
    // Re-throw McpError as-is
    if (error instanceof McpError) {
      throw error;
    }

    const message = error instanceof Error ? error.message : String(error);
    logger.error('Markdown update operation failed', {
      id,
      gid,
      requestCount: requests.length,
      error: message,
    });

    throw new McpError(ErrorCode.InternalError, `Error writing markdown to cells: ${message}`, {
      stack: error instanceof Error ? error.stack : undefined,
    });
  }
}

export default function createTool() {
  return {
    name: 'values-markdown-update',
    config,
    handler,
  } satisfies ToolModule;
}
