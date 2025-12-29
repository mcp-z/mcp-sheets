import type { EnrichedExtra } from '@mcp-z/oauth-google';
import { schemas } from '@mcp-z/oauth-google';

const { AuthRequiredBranchSchema } = schemas;

import type { ToolModule } from '@mcp-z/server';
import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js';
import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { google } from 'googleapis';
import { z } from 'zod';
import { SheetGidOutput, SheetGidSchema, SpreadsheetIdOutput, SpreadsheetIdSchema } from '../../schemas/index.js';

const MAX_BATCH_SIZE = 100;

const copyItemSchema = z.object({
  newTitle: z.coerce.string().trim().min(1).describe('Name for the copied sheet'),
  insertIndex: z.number().int().min(0).optional().describe('Position to insert the new sheet (0-indexed)'),
});

const inputSchema = z.object({
  id: SpreadsheetIdSchema,
  gid: SheetGidSchema,
  copies: z.array(copyItemSchema).min(1).max(MAX_BATCH_SIZE).describe('Array of copies to create from the source sheet'),
});

// Created sheet info
const createdSheetSchema = z.object({
  gid: SheetGidOutput,
  title: z.string().describe('Title of the created sheet'),
  sheetUrl: z.string().describe('URL of the created sheet'),
});

// Success branch schema - uses items: for consistency with standard vocabulary
const successBranchSchema = z.object({
  type: z.literal('success'),
  operationSummary: z.string().describe('Summary of the copy operation'),
  itemsProcessed: z.number().describe('Total copies attempted'),
  itemsChanged: z.number().describe('Successfully created copies'),
  completedAt: z.string().describe('ISO datetime when operation completed'),
  id: SpreadsheetIdOutput,
  sourceGid: z.string().describe('Source sheet ID'),
  sourceTitle: z.string().describe('Source sheet title'),
  items: z.array(createdSheetSchema).describe('Information about created sheets'),
  failures: z
    .array(
      z.object({
        title: z.string(),
        error: z.string(),
      })
    )
    .optional()
    .describe('Failed copies with error messages'),
});

// Output schema with auth_required support
const outputSchema = z.discriminatedUnion('type', [successBranchSchema, AuthRequiredBranchSchema]);

const config = {
  description: 'Copy a sheet/tab within the same spreadsheet. Supports batch copying: create multiple copies from a single source sheet (e.g., create 12 monthly sheets from a template). Copies all data, formatting, charts, and conditional formatting verbatim.',
  inputSchema,
  outputSchema: z.object({
    result: outputSchema,
  }),
} as const;

export type Input = z.infer<typeof inputSchema>;
export type Output = z.infer<typeof outputSchema>;

async function handler({ id, gid, copies }: Input, extra: EnrichedExtra): Promise<CallToolResult> {
  const logger = extra.logger;
  logger.info('sheets.sheet.copy called', { id, gid, copyCount: copies.length });

  try {
    const sheets = google.sheets({ version: 'v4', auth: extra.authContext.auth });

    // First, get the source sheet info
    const spreadsheetInfo = await sheets.spreadsheets.get({
      spreadsheetId: id,
      fields: 'sheets.properties.sheetId,sheets.properties.title',
    });

    const sourceSheet = spreadsheetInfo.data.sheets?.find((s) => String(s.properties?.sheetId) === gid);
    if (!sourceSheet?.properties) {
      throw new McpError(ErrorCode.InvalidParams, `Source sheet with gid "${gid}" not found in spreadsheet`);
    }

    const sourceTitle = sourceSheet.properties.title || '';
    const sourceSheetId = Number(gid);

    // Execute copies using Promise.allSettled for partial failure handling
    const results = await Promise.allSettled(
      copies.map(async (copy) => {
        const response = await sheets.spreadsheets.batchUpdate({
          spreadsheetId: id,
          requestBody: {
            requests: [
              {
                duplicateSheet: {
                  sourceSheetId,
                  newSheetName: copy.newTitle,
                  ...(copy.insertIndex !== undefined && { insertSheetIndex: copy.insertIndex }),
                },
              },
            ],
          },
        });

        const newSheetId = response.data.replies?.[0]?.duplicateSheet?.properties?.sheetId;
        const newTitle = response.data.replies?.[0]?.duplicateSheet?.properties?.title;

        if (!newSheetId || !newTitle) {
          throw new Error('Failed to retrieve new sheet info from API response');
        }

        return {
          gid: String(newSheetId),
          title: newTitle,
          sheetUrl: `https://docs.google.com/spreadsheets/d/${id}/edit#gid=${newSheetId}`,
        };
      })
    );

    // Separate successes and failures
    const items: Array<{ gid: string; title: string; sheetUrl: string }> = [];
    const failures: Array<{ title: string; error: string }> = [];

    results.forEach((result, index) => {
      const copy = copies[index];
      if (!copy) return;

      if (result.status === 'fulfilled') {
        items.push(result.value);
      } else {
        const errorMessage = result.reason instanceof Error ? result.reason.message : String(result.reason);
        failures.push({ title: copy.newTitle, error: errorMessage });
      }
    });

    const successCount = items.length;
    const failureCount = failures.length;
    const totalCount = copies.length;

    const summary = failureCount === 0 ? `Created ${successCount} cop${successCount === 1 ? 'y' : 'ies'} of "${sourceTitle}"` : `Created ${successCount} of ${totalCount} cop${totalCount === 1 ? 'y' : 'ies'} (${failureCount} failed)`;

    logger.info('sheets.sheet.copy completed', { totalCount, successCount, failureCount });

    const result: Output = {
      type: 'success' as const,
      operationSummary: summary,
      itemsProcessed: totalCount,
      itemsChanged: successCount,
      completedAt: new Date().toISOString(),
      id,
      sourceGid: gid,
      sourceTitle,
      items,
      ...(failures.length > 0 && { failures }),
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
    logger.error('sheets.sheet.copy error', { error: message });

    throw new McpError(ErrorCode.InternalError, `Error copying sheet: ${message}`, {
      stack: error instanceof Error ? error.stack : undefined,
    });
  }
}

export default function createTool() {
  return {
    name: 'sheet-copy',
    config,
    handler,
  } satisfies ToolModule;
}
