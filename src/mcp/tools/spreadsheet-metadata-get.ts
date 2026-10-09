import { sheets as sheetsApi } from '@googleapis/sheets';
import type { EnrichedExtra } from '@mcp-z/oauth-google';
import { schemas } from '@mcp-z/oauth-google';
import type { CallToolResult, ToolModule } from '@mcp-z/server';
import { ProtocolError, ProtocolErrorCode } from '@mcp-z/server';
import { z } from 'zod';
import { googleAuth } from '../../lib/google-auth.ts';
import { SpreadsheetIdOutput, SpreadsheetIdSchema } from '../../schemas/index.ts';

const { AuthRequiredBranchSchema } = schemas;
const FIELD_MASK = 'spreadsheetId,properties,namedRanges,developerMetadata,dataSources,dataSourceSchedules,sheets(properties,charts,conditionalFormats,protectedRanges,tables,basicFilter,filterViews,bandedRanges,merges,slicers,developerMetadata)';

const inputSchema = z.object({ id: SpreadsheetIdSchema }).strict();
const metadataSchema = z
  .object({
    spreadsheetId: z.string(),
    properties: z.object({}).passthrough().optional(),
    sheets: z.array(z.object({ properties: z.object({}).passthrough().optional() }).passthrough()).optional(),
  })
  .passthrough();
const outputSchema = z.discriminatedUnion('type', [
  z.object({
    type: z.literal('success'),
    id: SpreadsheetIdOutput,
    metadata: metadataSchema.describe('Google spreadsheet envelope with full workbook and sheet properties and selected structural collections. No grid data, values, or direct cell formatting.'),
    fieldMask: z.string(),
  }),
  AuthRequiredBranchSchema,
]);
const config = {
  description: 'Read full workbook and sheet properties plus charts, named ranges, protections, conditional formats, tables, filters, banded ranges, merges, slicers, developer metadata, data sources and schedules. Returns all sheets without grid data, values or direct cell formatting.',
  annotations: { readOnlyHint: true },
  inputSchema,
  outputSchema: z.object({ result: outputSchema }),
} as const;

export type Input = z.infer<typeof inputSchema>;
export type Output = z.infer<typeof outputSchema>;

async function handler({ id }: Input, extra: EnrichedExtra): Promise<CallToolResult> {
  const logger = extra.logger;
  logger.info('sheets.spreadsheet.metadata.get called', { id });
  try {
    const sheets = sheetsApi({ version: 'v4', auth: googleAuth(extra.authContext.auth) });
    const response = await sheets.spreadsheets.get({ spreadsheetId: id, fields: FIELD_MASK });
    const result: Output = outputSchema.parse({ type: 'success', id, metadata: response.data, fieldMask: FIELD_MASK });
    logger.info('sheets.spreadsheet.metadata.get success', { id });
    return {
      content: [{ type: 'text', text: JSON.stringify(result) }],
      structuredContent: { result },
    };
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    logger.error('sheets.spreadsheet.metadata.get error', { id, error: message });
    throw new ProtocolError(ProtocolErrorCode.InternalError, `Error getting spreadsheet metadata: ${message}`, {
      stack: error instanceof Error ? error.stack : undefined,
    });
  }
}

export default function createTool() {
  return { name: 'spreadsheet-metadata-get', config, handler } satisfies ToolModule;
}
