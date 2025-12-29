import type { EnrichedExtra } from '@mcp-z/oauth-google';
import { schemas } from '@mcp-z/oauth-google';

const { AuthRequiredBranchSchema } = schemas;

import type { ToolModule } from '@mcp-z/server';
import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js';
import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { google } from 'googleapis';
import { z } from 'zod';
import { SheetGidOutput, SpreadsheetIdOutput, SpreadsheetRefSchema } from '../../schemas/index.js';
import { findSpreadsheetsByRef } from '../../spreadsheet/spreadsheet-management.js';

/** Spreadsheet match result from findSpreadsheetsByRef */
interface SpreadsheetMatch {
  id: string;
  spreadsheetTitle: string | undefined;
  url: string | undefined;
  modifiedTime: string | null | undefined;
}

const inputSchema = z.object({
  spreadsheetRef: SpreadsheetRefSchema,
});

const successBranchSchema = z.object({
  type: z.literal('success'),
  items: z
    .array(
      z.object({
        id: SpreadsheetIdOutput,
        spreadsheetTitle: z.string().describe('Name of the spreadsheet'),
        spreadsheetUrl: z.string().optional().describe('URL to view the spreadsheet'),
        modifiedTime: z.string().optional().describe('Last modified timestamp (ISO format)'),
        sheets: z
          .array(
            z.object({
              gid: SheetGidOutput,
              sheetTitle: z.string().describe('Sheet tab name'),
            })
          )
          .describe('All sheets/tabs in this spreadsheet'),
      })
    )
    .describe('Matching spreadsheets sorted by modification time'),
});

const outputSchema = z.discriminatedUnion('type', [successBranchSchema, AuthRequiredBranchSchema]);

const config = {
  description: 'Find spreadsheet by ID, URL, or name. Returns all sheets/tabs in response.',
  inputSchema,
  outputSchema: z.object({
    result: outputSchema,
  }),
} as const;

export type Input = z.infer<typeof inputSchema>;
export type Output = z.infer<typeof outputSchema>;

async function handler({ spreadsheetRef }: Input, extra: EnrichedExtra): Promise<CallToolResult> {
  const logger = extra.logger;
  logger.info('sheets.spreadsheet.find called', { spreadsheetRef });

  try {
    const sheets = google.sheets({ version: 'v4', auth: extra.authContext.auth });
    const drive = google.drive({ version: 'v3', auth: extra.authContext.auth });

    const matches = (await findSpreadsheetsByRef(sheets, drive, spreadsheetRef)) as SpreadsheetMatch[];

    const sorted = (matches || []).slice().sort((a: SpreadsheetMatch, b: SpreadsheetMatch) => (b.modifiedTime ? new Date(b.modifiedTime).getTime() : 0) - (a.modifiedTime ? new Date(a.modifiedTime).getTime() : 0));

    const items = await Promise.all(
      sorted.map(async (m: SpreadsheetMatch) => {
        // Fetch sheet metadata for each spreadsheet
        const spreadsheetResponse = await sheets.spreadsheets.get({
          spreadsheetId: m.id,
          fields: 'sheets.properties.sheetId,sheets.properties.title',
        });

        const sheetsData = (spreadsheetResponse.data.sheets || []).map((sheet) => {
          const gid = String(sheet.properties?.sheetId ?? '');
          return {
            gid,
            sheetTitle: sheet.properties?.title ?? '',
          };
        });

        return {
          id: m.id,
          spreadsheetTitle: m.spreadsheetTitle ?? '',
          sheets: sheetsData,
          ...(m.url && { spreadsheetUrl: m.url }),
          ...(m.modifiedTime && { modifiedTime: m.modifiedTime }),
        };
      })
    );

    logger.info('sheets.spreadsheet.find success', { count: items.length });

    const result: Output = {
      type: 'success' as const,
      items,
    };

    return {
      content: [{ type: 'text' as const, text: JSON.stringify(result) }],
      structuredContent: { result },
    };
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    logger.error('sheets.spreadsheet.find error', { error: message });

    throw new McpError(ErrorCode.InternalError, `Error finding spreadsheet: ${message}`, {
      stack: error instanceof Error ? error.stack : undefined,
    });
  }
}

export default function createTool() {
  return {
    name: 'spreadsheet-find',
    config,
    handler,
  } satisfies ToolModule;
}
