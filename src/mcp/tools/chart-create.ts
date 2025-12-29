import type { EnrichedExtra } from '@mcp-z/oauth-google';
import { schemas } from '@mcp-z/oauth-google';

const { AuthRequiredBranchSchema } = schemas;

import type { ToolModule } from '@mcp-z/server';
import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js';
import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { google, type sheets_v4 } from 'googleapis';
import { z } from 'zod';
import { SheetGidOutput, SheetGidSchema, SpreadsheetIdOutput, SpreadsheetIdSchema } from '../../schemas/index.js';
import { parseA1Notation, rangeReferenceToGridRange } from '../../spreadsheet/range-operations.js';

// Input schema for chart position
const ChartPositionSchema = z.object({
  anchorCell: z.string().min(1).describe('A1 notation cell where chart top-left anchors (e.g., "F2")'),
  offsetX: z.number().int().default(0).describe('Horizontal pixel offset from anchor'),
  offsetY: z.number().int().default(0).describe('Vertical pixel offset from anchor'),
});

const inputSchema = z.object({
  id: SpreadsheetIdSchema,
  gid: SheetGidSchema,
  chartType: z.enum(['PIE', 'BAR', 'COLUMN', 'LINE']).describe('Type of chart to create'),
  dataRange: z.string().min(1).describe('A1 notation range containing chart data including headers (e.g., "A1:C10")'),
  title: z.string().optional().describe('Chart title displayed at top'),
  position: ChartPositionSchema,
  legend: z.enum(['BOTTOM', 'RIGHT', 'TOP', 'LEFT', 'NONE']).default('BOTTOM').describe('Legend position'),
  is3D: z.boolean().default(false).describe('Render as 3D chart (PIE only)'),
});

// Success branch schema
const successBranchSchema = z.object({
  type: z.literal('success'),
  id: SpreadsheetIdOutput,
  gid: SheetGidOutput,
  sheetTitle: z.string().describe('Title of the sheet containing the chart'),
  sheetUrl: z.string().describe('URL of the sheet containing the chart'),
  chartId: z.number().int().describe('Unique chart ID for future updates/deletion'),
  anchorCell: z.string().describe('Where chart was anchored'),
});

// Output schema with auth_required support
const outputSchema = z.discriminatedUnion('type', [successBranchSchema, AuthRequiredBranchSchema]);

const config = {
  description: 'Create charts (pie, bar, column, line) from spreadsheet data ranges. Charts anchor to specific cells with optional pixel offsets. Data range should include headers. PIE charts use 2 columns (labels, values). BAR/COLUMN/LINE charts use first row as headers. Best for visualizing spreadsheet data.',
  inputSchema,
  outputSchema: z.object({
    result: outputSchema,
  }),
} as const;

export type Input = z.infer<typeof inputSchema>;
export type Output = z.infer<typeof outputSchema>;

// Parse A1 notation anchor cell to row/column indices
function parseAnchorCell(anchorCell: string): { rowIndex: number; columnIndex: number } {
  // Simple A1 notation parser for single cells
  const match = anchorCell.match(/^([A-Z]+)(\d+)$/);
  if (!match || !match[1] || !match[2]) {
    throw new Error(`Invalid anchor cell format: ${anchorCell}`);
  }

  const colLetters = match[1];
  const rowNumber = match[2];

  // Convert column letters to index (A=0, B=1, ..., Z=25, AA=26, etc.)
  let columnIndex = 0;
  for (let i = 0; i < colLetters.length; i++) {
    columnIndex = columnIndex * 26 + (colLetters.charCodeAt(i) - 65 + 1);
  }
  columnIndex--; // Convert to 0-based

  // Convert row number to index (1-based to 0-based)
  const rowIndex = Number.parseInt(rowNumber, 10) - 1;

  return { rowIndex, columnIndex };
}

async function handler({ id, gid, chartType, dataRange, title, position, legend = 'BOTTOM', is3D = false }: Input, extra: EnrichedExtra): Promise<CallToolResult> {
  const logger = extra.logger;
  logger.info('sheets.chart.create called', {
    id,
    gid,
    chartType,
    dataRange,
    title,
    position,
    legend,
    is3D,
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
      logger.info('Sheet not found for chart create', { id, gid, chartType });
      throw new McpError(ErrorCode.InvalidParams, `Sheet not found: ${gid}`);
    }

    const sheetTitle = sheet.properties.title ?? gid;
    const sheetId = sheet.properties.sheetId;
    const sheetUrl = `https://docs.google.com/spreadsheets/d/${id}/edit#gid=${sheetId}`;

    // Validate 3D only for PIE charts
    // Note: Google Sheets API only supports 3D for PIE charts, not COLUMN/BAR/LINE
    if (is3D && chartType !== 'PIE') {
      logger.info('3D mode not supported for this chart type', {
        chartType,
        is3D,
      });
      throw new McpError(ErrorCode.InvalidParams, `3D mode is only supported for PIE charts, not ${chartType}`);
    }

    // Parse anchor cell to row/column indices
    let anchorRowIndex: number;
    let anchorColumnIndex: number;
    try {
      const anchorIndices = parseAnchorCell(position.anchorCell);
      anchorRowIndex = anchorIndices.rowIndex;
      anchorColumnIndex = anchorIndices.columnIndex;
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      logger.info('Failed to parse anchor cell', {
        anchorCell: position.anchorCell,
        error: message,
      });
      throw new McpError(ErrorCode.InvalidParams, `Failed to parse anchor cell: ${message}`);
    }

    // Parse data range to grid range
    let dataGridRange: sheets_v4.Schema$GridRange;
    try {
      const rangeRef = parseA1Notation(dataRange);
      dataGridRange = rangeReferenceToGridRange(rangeRef, sheetId);

      // Validate that required properties are defined for chart creation
      if (dataGridRange.startColumnIndex === undefined || dataGridRange.startColumnIndex === null) {
        throw new Error('Data range must include column information');
      }
      if (dataGridRange.endColumnIndex === undefined || dataGridRange.endColumnIndex === null) {
        throw new Error('Data range must include column information');
      }
      if (dataGridRange.startRowIndex === undefined || dataGridRange.startRowIndex === null) {
        throw new Error('Data range must include row information');
      }
      if (dataGridRange.endRowIndex === undefined || dataGridRange.endRowIndex === null) {
        throw new Error('Data range must include row information');
      }
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      logger.info('Failed to parse data range', {
        dataRange,
        error: message,
      });
      throw new McpError(ErrorCode.InvalidParams, `Failed to parse data range: ${message}`);
    }

    // Extract validated properties with defined types
    const dataStartRowIndex = dataGridRange.startRowIndex as number;
    const dataEndRowIndex = dataGridRange.endRowIndex as number;
    const dataStartColumnIndex = dataGridRange.startColumnIndex as number;
    const dataEndColumnIndex = dataGridRange.endColumnIndex as number;
    const dataSheetId = dataGridRange.sheetId ?? sheetId;

    // Build chart spec with properly split domain and series ranges
    let chartSpec: sheets_v4.Schema$ChartSpec;

    logger.info('Building chart spec', {
      chartType,
      dataGridRange,
      sheetId,
    });

    // Map legend position to API format (add _LEGEND suffix)
    const legendPositionMap: Record<string, string> = {
      BOTTOM: 'BOTTOM_LEGEND',
      RIGHT: 'RIGHT_LEGEND',
      TOP: 'TOP_LEGEND',
      LEFT: 'LEFT_LEGEND',
      NONE: 'NO_LEGEND',
    };
    const apiLegendPosition = legendPositionMap[legend] || 'BOTTOM_LEGEND';

    if (chartType === 'PIE') {
      // PIE charts use a separate pieChart property, not basicChart
      const baseSheetId = dataSheetId;

      const domainRange = {
        sheetId: baseSheetId,
        startRowIndex: dataStartRowIndex,
        endRowIndex: dataEndRowIndex,
        startColumnIndex: dataStartColumnIndex,
        endColumnIndex: dataStartColumnIndex + 1, // First column only (labels)
      };

      const seriesRange = {
        sheetId: baseSheetId,
        startRowIndex: dataStartRowIndex,
        endRowIndex: dataEndRowIndex,
        startColumnIndex: dataStartColumnIndex + 1, // Second column (values)
        endColumnIndex: dataStartColumnIndex + 2,
      };

      logger.info('PIE chart ranges', {
        dataGridRange,
        domainRange,
        seriesRange,
        apiLegendPosition,
      });

      chartSpec = {
        pieChart: {
          // Use pieChart for PIE charts
          legendPosition: legend === 'NONE' ? 'NO_LEGEND' : apiLegendPosition,
          domain: {
            // Note: domain/series are direct objects, not arrays
            sourceRange: {
              sources: [domainRange],
            },
          },
          series: {
            sourceRange: {
              sources: [seriesRange],
            },
          },
          threeDimensional: is3D, // 3D is directly on pieChart
        },
      };
      if (title) {
        chartSpec.title = title;
      }
    } else {
      // For BAR, COLUMN, LINE charts: domain is first column, each subsequent column is a series
      const baseSheetId = dataSheetId;

      const domainRange = {
        sheetId: baseSheetId,
        startRowIndex: dataStartRowIndex,
        endRowIndex: dataEndRowIndex,
        startColumnIndex: dataStartColumnIndex,
        endColumnIndex: dataStartColumnIndex + 1, // First column only
      };

      // Calculate number of data columns (excluding the first domain column)
      const numDataColumns = dataEndColumnIndex - dataStartColumnIndex - 1;
      const series = [];

      // Create a series for each data column
      for (let i = 0; i < numDataColumns; i++) {
        const seriesRange = {
          sheetId: baseSheetId,
          startRowIndex: dataStartRowIndex,
          endRowIndex: dataEndRowIndex,
          startColumnIndex: dataStartColumnIndex + 1 + i,
          endColumnIndex: dataStartColumnIndex + 2 + i,
        };

        series.push({
          series: {
            sourceRange: {
              sources: [seriesRange],
            },
          },
        });
      }

      // If no data columns were found, default to using the entire range as a fallback
      if (series.length === 0) {
        series.push({
          series: {
            sourceRange: {
              sources: [dataGridRange],
            },
          },
        });
      }

      const basicChart: sheets_v4.Schema$BasicChartSpec = {
        chartType, // BAR, COLUMN, LINE are valid for basicChart
        headerCount: 1, // First row is headers
        // Note: Google Sheets API does not support threeDimensional for basicChart (BAR/COLUMN/LINE)
        // threeDimensional is only supported for pieChart
        domains: [
          {
            domain: {
              sourceRange: {
                sources: [domainRange],
              },
            },
          },
        ],
        series,
      };
      if (legend !== 'NONE') {
        basicChart.legendPosition = apiLegendPosition;
      }

      chartSpec = { basicChart };
      if (title) {
        chartSpec.title = title;
      }
    }

    // Build embedded object position
    const embeddedObjectPosition = {
      overlayPosition: {
        anchorCell: {
          sheetId,
          rowIndex: anchorRowIndex,
          columnIndex: anchorColumnIndex,
        },
        offsetXPixels: position.offsetX,
        offsetYPixels: position.offsetY,
      },
    };

    const requestBody = {
      requests: [
        {
          addChart: {
            chart: {
              spec: chartSpec,
              position: embeddedObjectPosition,
            },
          },
        },
      ],
    };

    logger.info('sheets.chart.create executing addChart request', {
      spreadsheetId: id,
      sheetTitle,
      chartType,
      dataRange,
      anchorCell: position.anchorCell,
      requestBody: JSON.stringify(requestBody),
    });

    // Execute the addChart request
    const response = await sheets.spreadsheets.batchUpdate({
      spreadsheetId: id,
      requestBody,
    });

    // Extract chart ID from response
    const replies = response.data.replies || [];
    if (replies.length === 0 || !replies[0]?.addChart?.chart?.chartId) {
      logger.error('Chart creation failed - no chart ID returned', {
        spreadsheetId: id,
        sheetTitle,
        chartType,
      });
      throw new McpError(ErrorCode.InternalError, 'Chart creation failed: no chart ID returned from Google Sheets API');
    }

    const chartId = replies[0].addChart.chart.chartId;

    logger.info('sheets.chart.create completed successfully', {
      chartId,
      chartType,
      anchorCell: position.anchorCell,
    });

    const result: Output = {
      type: 'success' as const,
      id,
      gid: String(sheetId),
      sheetTitle,
      sheetUrl,
      chartId,
      anchorCell: position.anchorCell,
    };

    return {
      content: [{ type: 'text' as const, text: JSON.stringify(result) }],
      structuredContent: { result },
    };
  } catch (e) {
    const error = e as Error & { response?: { data?: unknown; status?: number } };
    const message = error.message || String(e);
    logger.error('Chart create operation failed', {
      id,
      gid,
      chartType,
      dataRange,
      error: message,
      response: error.response?.data,
      status: error.response?.status,
    });

    throw new McpError(ErrorCode.InternalError, `Error creating chart: ${message}`, {
      stack: e instanceof Error ? e.stack : undefined,
    });
  }
}

export default function createTool() {
  return {
    name: 'chart-create',
    config,
    handler,
  } satisfies ToolModule;
}
