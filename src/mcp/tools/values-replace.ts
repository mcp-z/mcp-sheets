import type { EnrichedExtra } from '@mcp-z/oauth-google';
import { schemas } from '@mcp-z/oauth-google';

const { AuthRequiredBranchSchema } = schemas;

import type { ToolModule } from '@mcp-z/server';
import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js';
import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { google, type sheets_v4 } from 'googleapis';
import { z } from 'zod';
import { A1NotationSchema, SheetGidSchema, SpreadsheetIdOutput, SpreadsheetIdSchema } from '../../schemas/index.ts';
import { parseA1Notation, rangeReferenceToGridRange } from '../../spreadsheet/range-operations.ts';

const inputSchema = z
  .object({
    id: SpreadsheetIdSchema,
    find: z.string().min(1).describe('Text or regex pattern to find'),
    replacement: z.string().describe('Replacement text. Use $1, $2 for regex capture groups. Empty string deletes matches.'),

    // Scope - defaults to all sheets if neither specified
    gid: SheetGidSchema.optional().describe('Limit to specific sheet. If omitted, searches all sheets.'),
    range: A1NotationSchema.optional().describe('Limit to specific A1 range within the sheet (requires gid)'),

    // Match options
    matchCase: z.boolean().optional().describe('Case-sensitive matching'),
    matchEntireCell: z.boolean().optional().describe('Only match entire cell content'),
    searchByRegex: z.boolean().optional().describe('Treat find as RE2 regex'),
    includeFormulas: z.boolean().optional().describe('Search within formula text'),
  })
  .refine((data) => !(data.range && !data.gid), { message: 'range requires gid' });

const successBranchSchema = z.object({
  type: z.literal('success'),
  id: SpreadsheetIdOutput,
  spreadsheetUrl: z.string().describe('URL of the spreadsheet'),
  occurrencesChanged: z.number().int().nonnegative().describe('Total replacements made'),
  valuesChanged: z.number().int().nonnegative().describe('Number of non-formula cells changed'),
  formulasChanged: z.number().int().nonnegative().describe('Number of formula cells changed'),
  rowsChanged: z.number().int().nonnegative().describe('Number of rows with replacements'),
  sheetsChanged: z.number().int().nonnegative().describe('Number of sheets with replacements'),
});

const outputSchema = z.discriminatedUnion('type', [successBranchSchema, AuthRequiredBranchSchema]);

const config = {
  description: 'Find and replace text across a spreadsheet. Searches all sheets by default, or limit with gid/range. Supports regex with capture groups ($1, $2).',
  inputSchema,
  outputSchema: z.object({
    result: outputSchema,
  }),
} as const;

export type Input = z.infer<typeof inputSchema>;
export type Output = z.infer<typeof outputSchema>;

async function handler({ id, find, replacement, gid, range, matchCase, matchEntireCell, searchByRegex, includeFormulas }: Input, extra: EnrichedExtra): Promise<CallToolResult> {
  const logger = extra.logger;
  logger.info('sheets.values-replace called', { id, find, replacement, gid, range });

  try {
    const sheets = google.sheets({ version: 'v4', auth: extra.authContext.auth });

    // Build FindReplaceRequest - only include defined options
    const findReplaceRequest: sheets_v4.Schema$FindReplaceRequest = {
      find,
      replacement,
      ...(matchCase !== undefined && { matchCase }),
      ...(matchEntireCell !== undefined && { matchEntireCell }),
      ...(searchByRegex !== undefined && { searchByRegex }),
      ...(includeFormulas !== undefined && { includeFormulas }),
    };

    // Set scope based on gid/range
    if (!gid) {
      // No gid = search all sheets
      findReplaceRequest.allSheets = true;
    } else {
      // Need to resolve sheet to get numeric sheetId
      const spreadsheetResponse = await sheets.spreadsheets.get({
        spreadsheetId: id,
        fields: 'sheets.properties.sheetId,sheets.properties.title',
      });

      const sheet = spreadsheetResponse.data.sheets?.find((s) => String(s.properties?.sheetId) === gid);
      if (!sheet) {
        throw new McpError(ErrorCode.InvalidParams, `Sheet not found: ${gid}`);
      }

      // Note: sheetId can be 0 which is falsy, so check explicitly for undefined/null
      const sheetId = sheet.properties?.sheetId;
      if (sheetId === undefined || sheetId === null) {
        throw new McpError(ErrorCode.InternalError, 'Sheet properties not available');
      }

      if (!range) {
        // gid but no range = search specific sheet
        findReplaceRequest.sheetId = sheetId;
      } else {
        // gid + range = search specific range
        const rangeRef = parseA1Notation(range);
        findReplaceRequest.range = rangeReferenceToGridRange(rangeRef, sheetId);
      }
    }

    // Execute batchUpdate
    const response = await sheets.spreadsheets.batchUpdate({
      spreadsheetId: id,
      requestBody: {
        requests: [{ findReplace: findReplaceRequest }],
      },
    });

    const findReplaceResponse = response.data.replies?.[0]?.findReplace;
    const spreadsheetUrl = `https://docs.google.com/spreadsheets/d/${id}/edit`;

    logger.info('sheets.values-replace completed', {
      occurrencesChanged: findReplaceResponse?.occurrencesChanged ?? 0,
      sheetsChanged: findReplaceResponse?.sheetsChanged ?? 0,
    });

    const result: Output = {
      type: 'success' as const,
      id,
      spreadsheetUrl,
      occurrencesChanged: findReplaceResponse?.occurrencesChanged ?? 0,
      valuesChanged: findReplaceResponse?.valuesChanged ?? 0,
      formulasChanged: findReplaceResponse?.formulasChanged ?? 0,
      rowsChanged: findReplaceResponse?.rowsChanged ?? 0,
      sheetsChanged: findReplaceResponse?.sheetsChanged ?? 0,
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
    logger.error('Replace operation failed', {
      id,
      find,
      replacement,
      gid,
      range,
      error: message,
    });

    throw new McpError(ErrorCode.InternalError, `Error replacing values: ${message}`, {
      stack: error instanceof Error ? error.stack : undefined,
    });
  }
}

export default function createTool() {
  return {
    name: 'values-replace',
    config,
    handler,
  } satisfies ToolModule;
}
