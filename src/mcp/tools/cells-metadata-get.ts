import { sheets as sheetsApi } from '@googleapis/sheets';
import type { EnrichedExtra } from '@mcp-z/oauth-google';
import { schemas } from '@mcp-z/oauth-google';
import type { CallToolResult, ToolModule } from '@mcp-z/server';
import { ProtocolError, ProtocolErrorCode } from '@mcp-z/server';
import { z } from 'zod';
import { googleAuth } from '../../lib/google-auth.ts';
import { SheetGidOutput, SpreadsheetIdOutput, SpreadsheetIdSchema } from '../../schemas/index.ts';
import { calculateFiniteRangeCellCount } from '../../spreadsheet/range-operations.ts';

const { AuthRequiredBranchSchema } = schemas;
const MAX_RANGES = 50;
const MAX_CELLS = 250_000;
const FIELD_MASK = 'spreadsheetId,sheets(properties(sheetId,title),data(startRow,startColumn,rowData(values(note,dataValidation,pivotTable,dataSourceTable))))';

const rangesSchema = z
  .array(z.string().trim().min(1))
  .min(1)
  .max(MAX_RANGES)
  .superRefine((ranges, context) => {
    try {
      const totalCells = ranges.reduce((sum, range) => sum + calculateFiniteRangeCellCount(range), 0);
      if (totalCells > MAX_CELLS) throw new Error(`Requested ranges exceed the ${MAX_CELLS.toLocaleString('en-US')} aggregate cell limit`);
    } catch (error) {
      context.addIssue({ code: 'custom', message: error instanceof Error ? error.message : String(error) });
    }
  });
const inputSchema = z
  .object({
    id: SpreadsheetIdSchema,
    gid: z
      .union([z.string().min(1), z.number()])
      .transform(String)
      .describe('Sheet ID (from URL gid={gid})'),
    ranges: rangesSchema.describe(`Required finite sheet-local A1 ranges, such as A1 or B2:D10, for notes, dataValidation, pivotTable and dataSourceTable. Maximum ${MAX_RANGES} ranges and ${MAX_CELLS.toLocaleString('en-US')} aggregate cells, counting overlapping ranges separately.`),
  })
  .strict();
const definitionSchema = z.object({}).passthrough();
const cellSchema = z
  .object({
    note: z.string().optional(),
    dataValidation: definitionSchema.optional(),
    pivotTable: definitionSchema.optional(),
    dataSourceTable: definitionSchema.optional(),
  })
  .passthrough();
const metadataSchema = z
  .object({
    spreadsheetId: z.string(),
    sheets: z
      .array(
        z
          .object({
            properties: z.object({ sheetId: z.number().int().optional(), title: z.string().optional() }).passthrough().optional(),
            data: z
              .array(
                z
                  .object({
                    startRow: z.number().int().nonnegative().optional(),
                    startColumn: z.number().int().nonnegative().optional(),
                    rowData: z.array(z.object({ values: z.array(cellSchema).optional() }).passthrough()).optional(),
                  })
                  .passthrough()
              )
              .optional(),
          })
          .passthrough()
      )
      .optional(),
  })
  .passthrough();
const outputSchema = z.discriminatedUnion('type', [
  z.object({
    type: z.literal('success'),
    id: SpreadsheetIdOutput,
    gid: SheetGidOutput,
    metadata: metadataSchema.describe('Selected Google cell metadata with sheet identity and sparse GridData offsets. Omitted startRow/startColumn mean zero; empty rows and cells retain their positions. Values and direct formatting are excluded.'),
    fieldMask: z.string(),
    requestedRanges: z.array(z.string()).describe('Sheet-qualified A1 ranges sent to Google, in request order.'),
  }),
  AuthRequiredBranchSchema,
]);
const config = {
  description: 'Read notes, data validation rules, pivot table and data-source table definitions in required finite ranges on one sheet. Preserves sparse row/column offsets. Excludes cell values, formulas and direct formatting. DATA_SOURCE and OBJECT sheets are unsupported.',
  annotations: { readOnlyHint: true },
  inputSchema,
  outputSchema: z.object({ result: outputSchema }),
} as const;

export type Input = z.infer<typeof inputSchema>;
export type Output = z.infer<typeof outputSchema>;

async function handler(input: Input, extra: EnrichedExtra): Promise<CallToolResult> {
  const parsed = inputSchema.safeParse(input);
  if (!parsed.success) throw new ProtocolError(ProtocolErrorCode.InvalidParams, parsed.error.message);
  const { id, gid, ranges } = parsed.data;
  const logger = extra.logger;
  logger.info('sheets.cells.metadata.get called', { id, gid, rangeCount: ranges.length });
  try {
    const sheets = sheetsApi({ version: 'v4', auth: googleAuth(extra.authContext.auth) });
    // Resolve the title and type without fetching structures. DATA_SOURCE sheets cannot accept specific ranges.
    const preflight = await sheets.spreadsheets.get({ spreadsheetId: id, fields: 'sheets(properties(sheetId,title,sheetType))' });
    const properties = preflight.data.sheets?.find((sheet) => String(sheet.properties?.sheetId) === gid)?.properties;
    if (!properties) throw new ProtocolError(ProtocolErrorCode.InvalidParams, `Sheet not found: ${gid}`);
    if (properties.sheetType === 'DATA_SOURCE' || properties.sheetType === 'OBJECT') {
      throw new ProtocolError(ProtocolErrorCode.InvalidParams, `Cell metadata ranges are not supported on ${properties.sheetType} sheets: ${gid}`);
    }
    if (properties.title == null) throw new Error(`Sheet title not available for ${gid}`);
    const title = `'${properties.title.replace(/'/g, "''")}'`;
    const requestedRanges = ranges.map((range) => `${title}!${range}`);
    const response = await sheets.spreadsheets.get({ spreadsheetId: id, fields: FIELD_MASK, ranges: requestedRanges });
    const result: Output = outputSchema.parse({ type: 'success', id, gid, metadata: response.data, fieldMask: FIELD_MASK, requestedRanges });
    logger.info('sheets.cells.metadata.get success', { id, gid, rangeCount: ranges.length });
    return {
      content: [{ type: 'text', text: JSON.stringify(result) }],
      structuredContent: { result },
    };
  } catch (error) {
    if (error instanceof ProtocolError) throw error;
    const message = error instanceof Error ? error.message : String(error);
    logger.error('sheets.cells.metadata.get error', { id, gid, error: message });
    throw new ProtocolError(ProtocolErrorCode.InternalError, `Error getting cell metadata: ${message}`, {
      stack: error instanceof Error ? error.stack : undefined,
    });
  }
}

export default function createTool() {
  return { name: 'cells-metadata-get', config, handler } satisfies ToolModule;
}
