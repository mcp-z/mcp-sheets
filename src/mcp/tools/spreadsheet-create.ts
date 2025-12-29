import type { EnrichedExtra } from '@mcp-z/oauth-google';
import { schemas } from '@mcp-z/oauth-google';

const { AuthRequiredBranchSchema } = schemas;

import type { ToolModule } from '@mcp-z/server';
import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js';
import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { google } from 'googleapis';
import { z } from 'zod';
import { SpreadsheetIdOutput } from '../../schemas/index.js';

const inputSchema = z.object({
  title: z.coerce.string().trim().min(1).describe('Title for the new spreadsheet'),
});

// Success branch schema
const successBranchSchema = z.object({
  type: z.literal('success'),
  operationSummary: z.string().describe('Summary of the spreadsheet creation operation'),
  itemsProcessed: z.number().describe('Total items attempted (always 1 for single spreadsheet)'),
  itemsChanged: z.number().describe('Successfully created spreadsheets (always 1 on success)'),
  completedAt: z.string().describe('ISO datetime when operation completed'),
  id: SpreadsheetIdOutput,
  spreadsheetUrl: z.string().describe('URL of the created spreadsheet'),
  title: z.string().describe('Title of the created spreadsheet'),
});

// Output schema with auth_required support
const outputSchema = z.discriminatedUnion('type', [successBranchSchema, AuthRequiredBranchSchema]);

const config = {
  description: 'Create a new spreadsheet',
  inputSchema,
  outputSchema: z.object({
    result: outputSchema,
  }),
} as const;

export type Input = z.infer<typeof inputSchema>;
export type Output = z.infer<typeof outputSchema>;

async function handler({ title }: Input, extra: EnrichedExtra): Promise<CallToolResult> {
  const logger = extra.logger;
  logger.info('sheets.spreadsheet.create called', { title });

  try {
    const sheets = google.sheets({ version: 'v4', auth: extra.authContext.auth });
    const response = await sheets.spreadsheets.create({
      requestBody: {
        properties: {
          title: title,
        },
      },
    });
    const res = response.data;
    const id = res.spreadsheetId ?? '';
    const url = res.spreadsheetUrl ?? '';

    logger.info('sheets.spreadsheet.create success', { id, title, url });

    const result: Output = {
      type: 'success' as const,
      operationSummary: `Created spreadsheet "${title}"`,
      itemsProcessed: 1,
      itemsChanged: 1,
      completedAt: new Date().toISOString(),
      id,
      spreadsheetUrl: url,
      title,
    };

    return {
      content: [
        {
          type: 'text' as const,
          text: JSON.stringify(result),
        },
      ],
      structuredContent: { result },
    };
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    logger.error('sheets.spreadsheet.create error', { error: message });

    // Throw McpError for proper MCP error handling
    throw new McpError(ErrorCode.InternalError, `Error creating spreadsheet: ${message}`, {
      stack: error instanceof Error ? error.stack : undefined,
    });
  }
}

export default function createTool() {
  return {
    name: 'spreadsheet-create',
    config,
    handler,
  } satisfies ToolModule;
}
