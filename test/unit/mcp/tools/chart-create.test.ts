import type { Logger, LoopbackOAuthProvider } from '@mcp-z/oauth-google';
import assert from 'assert';
import crypto from 'crypto';
import fs from 'fs/promises';
import type { OAuth2Client } from 'google-auth-library';
import { google } from 'googleapis';
import * as path from 'path';
import createTool, { type Input, type Output } from '../../../../src/mcp/tools/chart-create.ts';
import createValuesBatchUpdateTool, { type Input as ValuesBatchUpdateInput } from '../../../../src/mcp/tools/values-batch-update.ts';
import { createExtra, type TypedHandler } from '../../../lib/create-extra.ts';
import createMiddlewareContext from '../../../lib/create-middleware-context.ts';
import { createTestSpreadsheet, deleteTestSpreadsheet } from '../../../lib/spreadsheet-helpers.ts';

/**
 * RANGE ALLOCATION MAP - chart-create.test.ts
 *
 * All tests use shared sheet (gid: 0) with non-overlapping data ranges and chart positions.
 *
 * Data Ranges:
 * - A1:B5 = Pie chart data (Category, Value) - used by PIE chart tests
 * - E1:G5 = Bar chart data (Month, Sales, Costs) - used by BAR/COLUMN/LINE chart tests
 *
 * Chart Anchor Positions (to avoid visual overlap):
 * - D2  = Test 1: 2D PIE chart
 * - H2  = Test 1: 3D PIE chart
 * - J2  = Test 2: BAR chart
 * - J12 = Test 2: COLUMN chart
 * - J22 = Test 2: LINE chart
 * - L2  = Test 3: Offset chart
 * - L12 = Test 3: No legend chart
 * - L22 = Test 3: No title chart
 * - N2  = Test 4: (validation errors - no charts created)
 * - P2  = Test 5: Domain/series split test
 * - R2  = Test 6: A1 range test
 *
 * Next available data range: I1:K5
 * Next available chart position: T2
 */

describe('chart-create tool (service-backed tests)', () => {
  // Shared test resources
  let sharedSpreadsheetId: string;
  let auth: OAuth2Client;
  let authProvider: LoopbackOAuthProvider;
  let logger: Logger;
  let accountId: string;
  let tmpDir: string;
  let handler: TypedHandler<Input>;
  let valuesBatchUpdateHandler: TypedHandler<ValuesBatchUpdateInput>;

  // All tests use shared default sheet (gid: 0)
  const sharedSheetId = 0;

  // Helper to add sample data to sheet for chart testing
  async function addBothChartDataTypes(): Promise<void> {
    // Add both pie and bar chart data in non-overlapping ranges
    const requests = [
      {
        range: 'A1:B5',
        values: [
          ['Category', 'Value'],
          ['Red', 30],
          ['Green', 25],
          ['Blue', 20],
          ['Yellow', 25],
        ],
        majorDimension: 'ROWS' as const,
      },
      {
        range: 'E1:G5',
        values: [
          ['Month', 'Sales', 'Costs'],
          ['Jan', 1000, 800],
          ['Feb', 1200, 900],
          ['Mar', 1100, 850],
          ['Apr', 1300, 950],
        ],
        majorDimension: 'ROWS' as const,
      },
    ];

    await valuesBatchUpdateHandler(
      {
        id: sharedSpreadsheetId,
        gid: String(sharedSheetId),
        requests: requests,
        valueInputOption: 'USER_ENTERED',
        includeData: false,
      },
      createExtra()
    );
  }

  before(async () => {
    try {
      // Create temporary directory
      tmpDir = path.join('.tmp', `chart-create-tests-${crypto.randomUUID()}`);
      await fs.mkdir(tmpDir, { recursive: true });

      // Get middleware for tool creation
      const middlewareContext = await createMiddlewareContext();
      authProvider = middlewareContext.authProvider;
      logger = middlewareContext.logger;
      auth = middlewareContext.auth;
      const middleware = middlewareContext.middleware;
      accountId = middlewareContext.accountId;
      const tool = createTool();
      const wrappedTool = middleware.withToolAuth(tool);
      handler = wrappedTool.handler;
      const valuesBatchUpdateTool = createValuesBatchUpdateTool();
      const wrappedValuesBatchUpdateTool = middleware.withToolAuth(valuesBatchUpdateTool);
      valuesBatchUpdateHandler = wrappedValuesBatchUpdateTool.handler;

      // Create shared spreadsheet for all tests (use default sheet to minimize write operations)
      const title = `ci-chart-create-tests-${Date.now()}`;
      sharedSpreadsheetId = await createTestSpreadsheet(await authProvider.getAccessToken(accountId), { title });

      // Pre-populate default sheet with both pie and bar chart data in non-overlapping ranges
      await addBothChartDataTypes();
    } catch (error) {
      logger.error('Failed to initialize test resources:', { error });
      throw error;
    }
  });

  after(async () => {
    // Cleanup shared spreadsheet (automatically deletes all sheets within it)
    const accessToken = await authProvider.getAccessToken(accountId);
    await deleteTestSpreadsheet(accessToken, sharedSpreadsheetId, logger);

    // Cleanup temporary directory
    await fs.rm(tmpDir, { recursive: true, force: true });
  });

  it('[D2,H2] chart_create creates PIE charts (2D and 3D)', async () => {
    // Test 2D PIE chart at D2
    const response2D = await handler(
      {
        id: sharedSpreadsheetId,
        gid: String(sharedSheetId),
        chartType: 'PIE',
        dataRange: 'A1:B5',
        title: 'Sales by Category',
        position: {
          anchorCell: 'D2',
          offsetX: 0,
          offsetY: 0,
        },
        legend: 'BOTTOM',
        is3D: false,
      },
      createExtra()
    );

    const structured2D = response2D.structuredContent?.result as Output | undefined;
    if (structured2D?.type === 'success') {
      assert.ok(typeof structured2D.chartId === 'number', 'Should return valid chartId');
      assert.strictEqual(structured2D.anchorCell, 'D2', 'Should anchor at D2');
    } else {
      assert.fail('Expected success result for 2D PIE chart');
    }

    // Test 3D PIE chart at H2 (different position to avoid overlap)
    const response3D = await handler(
      {
        id: sharedSpreadsheetId,
        gid: String(sharedSheetId),
        chartType: 'PIE',
        dataRange: 'A1:B5',
        title: '3D Pie Chart',
        position: {
          anchorCell: 'H2',
          offsetX: 0,
          offsetY: 0,
        },
        legend: 'RIGHT',
        is3D: true,
      },
      createExtra()
    );

    const structured3D = response3D.structuredContent?.result as Output | undefined;
    if (structured3D?.type === 'success') {
      assert.ok(typeof structured3D.chartId === 'number', 'Should return valid chartId');
    } else {
      assert.fail('Expected success result for 3D PIE chart');
    }
  });

  it('[J2,J12,J22] chart_create creates BAR, COLUMN, and LINE charts', async () => {
    // Test BAR chart at J2
    const responseBar = await handler(
      {
        id: sharedSpreadsheetId,
        gid: String(sharedSheetId),
        chartType: 'BAR',
        dataRange: 'E1:G5',
        title: 'Monthly Report',
        position: {
          anchorCell: 'J2',
          offsetX: 10,
          offsetY: 10,
        },
        legend: 'TOP',
        is3D: false,
      },
      createExtra()
    );

    const structuredBar = responseBar.structuredContent?.result as Output | undefined;
    if (structuredBar?.type === 'success') {
      assert.ok(typeof structuredBar.chartId === 'number', 'Should return valid chartId');
    } else {
      assert.fail('Expected success result for BAR chart');
    }

    // Test COLUMN chart at J12
    const responseColumn = await handler(
      {
        id: sharedSpreadsheetId,
        gid: String(sharedSheetId),
        chartType: 'COLUMN',
        dataRange: 'E1:G5',
        title: 'Sales vs Costs',
        position: {
          anchorCell: 'J12',
          offsetX: 0,
          offsetY: 0,
        },
        legend: 'BOTTOM',
        is3D: false,
      },
      createExtra()
    );

    const structuredColumn = responseColumn.structuredContent?.result as Output | undefined;
    if (structuredColumn?.type === 'success') {
      assert.ok(typeof structuredColumn.chartId === 'number', 'Should return valid chartId');
    } else {
      assert.fail('Expected success result for COLUMN chart');
    }

    // Test LINE chart at J22
    const responseLine = await handler(
      {
        id: sharedSpreadsheetId,
        gid: String(sharedSheetId),
        chartType: 'LINE',
        dataRange: 'E1:G5',
        title: 'Trends Over Time',
        position: {
          anchorCell: 'J22',
          offsetX: 0,
          offsetY: 0,
        },
        legend: 'LEFT',
        is3D: false,
      },
      createExtra()
    );

    const structuredLine = responseLine.structuredContent?.result as Output | undefined;
    if (structuredLine?.type === 'success') {
      assert.ok(typeof structuredLine.chartId === 'number', 'Should return valid chartId');
    } else {
      assert.fail('Expected success result for LINE chart');
    }
  });

  it('[L2,L12,L22] chart_create edge cases: offsets, no legend, no title, and 3D validation', async () => {
    // Test: positions chart with pixel offsets at L2
    const responseOffset = await handler(
      {
        id: sharedSpreadsheetId,
        gid: String(sharedSheetId),
        chartType: 'PIE',
        dataRange: 'A1:B5',
        title: 'Offset Chart',
        position: {
          anchorCell: 'L2',
          offsetX: 50,
          offsetY: 25,
        },
        legend: 'BOTTOM',
        is3D: false,
      },
      createExtra()
    );

    const structuredOffset = responseOffset.structuredContent?.result as Output | undefined;
    if (structuredOffset?.type === 'success') {
      assert.strictEqual(structuredOffset.anchorCell, 'L2', 'Should anchor at L2');
    } else {
      assert.fail('Expected success result with offsets');
    }

    // Test: chart with no legend at L12
    const responseNoLegend = await handler(
      {
        id: sharedSpreadsheetId,
        gid: String(sharedSheetId),
        chartType: 'PIE',
        dataRange: 'A1:B5',
        title: 'No Legend Chart',
        position: {
          anchorCell: 'L12',
          offsetX: 0,
          offsetY: 0,
        },
        legend: 'NONE',
        is3D: false,
      },
      createExtra()
    );

    const structuredNoLegend = responseNoLegend.structuredContent?.result as Output | undefined;
    if (structuredNoLegend?.type !== 'success') {
      assert.fail('Expected success result with no legend');
    }

    // Test: chart without title at L22
    const responseNoTitle = await handler(
      {
        id: sharedSpreadsheetId,
        gid: String(sharedSheetId),
        chartType: 'PIE',
        dataRange: 'A1:B5',
        // title omitted
        position: {
          anchorCell: 'L22',
          offsetX: 0,
          offsetY: 0,
        },
        legend: 'BOTTOM',
        is3D: false,
      },
      createExtra()
    );

    const structuredNoTitle = responseNoTitle.structuredContent?.result as Output | undefined;
    if (structuredNoTitle?.type === 'success') {
      assert.ok(typeof structuredNoTitle.chartId === 'number', 'Should return valid chartId');
    } else {
      assert.fail('Expected success result without title');
    }
  });

  it('[N2] chart_create rejects 3D for unsupported chart types', async () => {
    // Test: rejects 3D for BAR charts (validation errors - no charts created)
    await assert.rejects(
      async () => {
        await handler(
          {
            id: sharedSpreadsheetId,
            gid: String(sharedSheetId),
            chartType: 'BAR',
            dataRange: 'E1:G5',
            title: 'Invalid 3D BAR',
            position: {
              anchorCell: 'N2',
              offsetX: 0,
              offsetY: 0,
            },
            legend: 'BOTTOM',
            is3D: true,
          },
          createExtra()
        );
      },
      (error: unknown) => {
        assert.ok(error instanceof Error && error.message.includes('3D mode is only supported for PIE charts'), 'Error should mention 3D limitation');
        return true;
      }
    );

    // Test: rejects 3D for COLUMN charts (Google Sheets API doesn't support 3D COLUMN)
    await assert.rejects(
      async () => {
        await handler(
          {
            id: sharedSpreadsheetId,
            gid: String(sharedSheetId),
            chartType: 'COLUMN',
            dataRange: 'E1:G5',
            title: 'Invalid 3D COLUMN',
            position: {
              anchorCell: 'N2',
              offsetX: 0,
              offsetY: 0,
            },
            legend: 'BOTTOM',
            is3D: true,
          },
          createExtra()
        );
      },
      (error: unknown) => {
        assert.ok(error instanceof Error && error.message.includes('3D mode is only supported for PIE charts'), 'Error should mention 3D limitation');
        return true;
      }
    );

    // Test: rejects 3D for LINE charts
    await assert.rejects(
      async () => {
        await handler(
          {
            id: sharedSpreadsheetId,
            gid: String(sharedSheetId),
            chartType: 'LINE',
            dataRange: 'E1:G5',
            title: 'Invalid 3D LINE',
            position: {
              anchorCell: 'N2',
              offsetX: 0,
              offsetY: 0,
            },
            legend: 'BOTTOM',
            is3D: true,
          },
          createExtra()
        );
      },
      (error: unknown) => {
        assert.ok(error instanceof Error && error.message.includes('3D mode is only supported for PIE charts'), 'Error should mention 3D limitation');
        return true;
      }
    );
  });

  it('[P2] chart_create should properly split domain and series ranges for PIE chart', async () => {
    const response = await handler(
      {
        id: sharedSpreadsheetId,
        gid: String(sharedSheetId),
        chartType: 'PIE',
        dataRange: 'A1:B5',
        title: 'Domain/Series Split Test',
        position: {
          anchorCell: 'P2',
          offsetX: 0,
          offsetY: 0,
        },
        legend: 'RIGHT',
        is3D: false,
      },
      createExtra()
    );

    const structured = response.structuredContent?.result as Output | undefined;
    if (structured?.type !== 'success') {
      assert.fail(`Expected success but got: ${JSON.stringify(structured)}`);
    }
    assert.ok(structured.chartId, 'Chart should have been created with an ID');

    // Fetch the spreadsheet to verify chart configuration
    const sheets = google.sheets({ version: 'v4', auth: auth });
    const getResponse = await sheets.spreadsheets.get({
      spreadsheetId: sharedSpreadsheetId,
      includeGridData: false,
    });

    const chartSheet = getResponse.data.sheets?.find((s) => s.properties?.sheetId === sharedSheetId);
    const chart = chartSheet?.charts?.find((c) => c.chartId === structured.chartId);
    assert.ok(chart, 'Chart should exist in spreadsheet');

    // PIE charts use pieChart, not basicChart
    const pieChart = chart.spec?.pieChart;
    assert.ok(pieChart, 'Chart should have pieChart spec');

    // Check that domain and series are properly configured
    const domainRange = pieChart.domain?.sourceRange?.sources?.[0];
    const seriesRange = pieChart.series?.sourceRange?.sources?.[0];

    assert.ok(domainRange, 'Chart should have a domain range');
    assert.ok(seriesRange, 'Chart should have a series range');

    // Verify the ranges are correctly split
    // Domain should use only column A (labels): startColumnIndex=0, endColumnIndex=1
    assert.strictEqual(domainRange.startColumnIndex, 0, 'Domain should start at column A');
    assert.strictEqual(domainRange.endColumnIndex, 1, 'Domain should end at column B (exclusive)');
    assert.strictEqual(domainRange.startRowIndex, 0, 'Domain should start at row 1');
    assert.strictEqual(domainRange.endRowIndex, 5, 'Domain should end at row 5');

    // Series should use only column B (values): startColumnIndex=1, endColumnIndex=2
    assert.strictEqual(seriesRange.startColumnIndex, 1, 'Series should start at column B');
    assert.strictEqual(seriesRange.endColumnIndex, 2, 'Series should end at column C (exclusive)');
    assert.strictEqual(seriesRange.startRowIndex, 0, 'Series should start at row 1');
    assert.strictEqual(seriesRange.endRowIndex, 5, 'Series should end at row 5');
  });

  it('[R2] chart_create respects A1 range for data and does not use entire sheet', async () => {
    // Add additional data in columns I-J that should NOT be included in chart
    const sheets = google.sheets({ version: 'v4', auth: auth });

    // Get spreadsheet metadata to find sheet by ID
    const spreadsheet = await sheets.spreadsheets.get({
      spreadsheetId: sharedSpreadsheetId,
      includeGridData: false,
    });

    // Find the sheet with matching sheetId
    const sheet = spreadsheet.data.sheets?.find((s) => s.properties?.sheetId === sharedSheetId);

    if (!sheet?.properties?.title) {
      throw new Error(`Could not find sheet with ID ${sharedSheetId}`);
    }

    const sheetTitle = sheet.properties.title;

    await sheets.spreadsheets.values.update({
      spreadsheetId: sharedSpreadsheetId,
      range: `'${sheetTitle}'!I1:J5`,
      valueInputOption: 'RAW',
      requestBody: {
        values: [
          ['Should Not', 'Be Included'],
          ['Data1', '100'],
          ['Data2', '200'],
          ['Data3', '300'],
          ['Data4', '400'],
        ],
      },
    });

    // Create chart with ONLY data range A1:B5
    const response = await handler(
      {
        id: sharedSpreadsheetId,
        gid: String(sharedSheetId),
        chartType: 'PIE',
        dataRange: 'A1:B5', // Should use ONLY this range, not I1:J5
        title: 'Limited Range Chart',
        position: {
          anchorCell: 'R2',
          offsetX: 0,
          offsetY: 0,
        },
        legend: 'BOTTOM',
        is3D: false,
      },
      createExtra()
    );

    const structured = response.structuredContent?.result as Output | undefined;
    if (structured?.type !== 'success') {
      assert.fail('Expected success result');
    }
    assert.ok(typeof structured.chartId === 'number', 'Should return valid chartId');

    // Verify the chart was created with the correct data range by reading the spreadsheet
    const spreadsheetData = await sheets.spreadsheets.get({
      spreadsheetId: sharedSpreadsheetId,
      includeGridData: false,
    });

    const chartSheet = spreadsheetData.data.sheets?.find((s) => s.properties?.sheetId === sharedSheetId);
    assert.ok(chartSheet, 'Chart sheet should exist');

    const charts = chartSheet.charts || [];
    assert.ok(charts.length > 0, 'Sheet should have at least one chart');

    // Find our chart by chartId
    const ourChart = charts.find((c) => c.chartId === structured.chartId);
    assert.ok(ourChart, 'Our chart should exist in the sheet');

    // Verify the chart uses the correct data range
    const chartSpec = ourChart.spec;
    assert.ok(chartSpec?.pieChart, 'Chart should have pieChart spec for PIE charts');

    // For PIE charts, domain is a single object, not an array
    const domainSources = chartSpec.pieChart.domain?.sourceRange?.sources || [];
    assert.ok(domainSources.length > 0, 'Chart should have domain data sources');

    const domainGridRange = domainSources[0];
    assert.ok(domainGridRange, 'Chart should have domain gridRange');

    // Verify the domain gridRange matches A1:B5 (column A for labels)
    // If the bug exists, this will fail because it would include the entire sheet
    assert.strictEqual(domainGridRange.startRowIndex, 0, 'Domain should start at row 0 (A1)');
    assert.strictEqual(domainGridRange.endRowIndex, 5, 'Domain should end at row 5');
    assert.strictEqual(domainGridRange.startColumnIndex, 0, 'Domain should start at column 0 (A)');
    assert.strictEqual(domainGridRange.endColumnIndex, 1, 'Domain should end at column 1 (A)');
  });
});
