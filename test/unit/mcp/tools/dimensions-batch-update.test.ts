import type { Logger, LoopbackOAuthProvider } from '@mcp-z/oauth-google';
import assert from 'assert';
import crypto from 'crypto';
import fs from 'fs/promises';
import * as path from 'path';
import createTool, { type Input, type Output } from '../../../../src/mcp/tools/dimensions-batch-update.js';
import { createExtra, type TypedHandler } from '../../../lib/create-extra.js';
import createMiddlewareContext from '../../../lib/create-middleware-context.js';
import { createTestSpreadsheet, deleteTestSpreadsheet } from '../../../lib/spreadsheet-helpers.js';

describe('dimensions-batch-update tool (service-backed tests)', () => {
  // Shared test resources
  let sharedSpreadsheetId: string;
  let authProvider: LoopbackOAuthProvider;
  let logger: Logger;
  let accountId: string;
  let handler: TypedHandler<Input>;
  let tmpDir: string;

  before(async () => {
    try {
      // Create temporary directory
      tmpDir = path.join('.tmp', `dimensions-batch-update-tests-${crypto.randomUUID()}`);
      await fs.mkdir(tmpDir, { recursive: true });

      // Get middleware for tool creation
      const middlewareContext = await createMiddlewareContext();
      authProvider = middlewareContext.authProvider;
      logger = middlewareContext.logger;
      const middleware = middlewareContext.middleware;
      accountId = middlewareContext.accountId;
      const tool = createTool();
      const wrappedTool = middleware.withToolAuth(tool);
      handler = wrappedTool.handler;

      // Create shared spreadsheet for all tests
      const title = `ci-dimensions-batch-update-tests-${Date.now()}`;
      sharedSpreadsheetId = await createTestSpreadsheet(await authProvider.getAccessToken(accountId), { title });
    } catch (error) {
      logger.error('Failed to initialize test resources:', { error });
      throw error;
    }
  });

  after(async () => {
    // Cleanup resources - fail fast on errors
    const accessToken = await authProvider.getAccessToken(accountId);
    await deleteTestSpreadsheet(accessToken, sharedSpreadsheetId, logger);
    await fs.rm(tmpDir, { recursive: true, force: true });
  });

  it('dimensions_batchUpdate handles delete and multiple operations with optimal ordering', async () => {
    // Use default sheet (gid: 0) to minimize write operations
    const testSheetId = 0;

    // CONSOLIDATED TEST: Validates both single deleteDimension and multiple operations in one execution
    // Test multiple operations in intentionally sub-optimal order
    // Handler should automatically sort them: delete -> insert -> append
    const requests = [
      // Delete columns B and C (indices 1-3)
      {
        operation: 'deleteDimension' as const,
        dimension: 'COLUMNS' as const,
        startIndex: 1,
        endIndex: 3,
      },
      // These will be reordered by the handler
      {
        operation: 'appendDimension' as const, // Should be executed last
        dimension: 'ROWS' as const,
        startIndex: 0,
      },
      {
        operation: 'insertDimension' as const, // Should be executed second (after deletes)
        dimension: 'COLUMNS' as const,
        startIndex: 1,
        endIndex: 3,
        inheritFromBefore: false,
      },
      {
        operation: 'deleteDimension' as const, // Should be executed first
        dimension: 'ROWS' as const,
        startIndex: 500,
        endIndex: 600,
      },
    ];

    const response = await handler(
      {
        id: sharedSpreadsheetId,
        gid: String(testSheetId),
        requests,
      },
      createExtra()
    );

    // Validate complete response structure according to outputSchema
    assert.ok(response, 'Handler returned no result');

    // Validate structuredContent.result matches outputSchema
    const structured = response.structuredContent?.result as Output | undefined;
    assert.ok(structured, 'Response missing structuredContent.result');

    // Fail fast with clear error message if operation failed
    if (structured?.type !== 'success') {
      assert.fail('Dimension batch operation failed');
    }

    // Validate response structure matches DimensionsBatchUpdateResponseSchema
    assert.ok(typeof structured.id === 'string', 'Item missing valid id');
    assert.ok(typeof structured.spreadsheetTitle === 'string', 'Item missing valid spreadsheetTitle');
    assert.ok(typeof structured.spreadsheetUrl === 'string', 'Item missing valid spreadsheetUrl');
    assert.ok(typeof structured.sheetTitle === 'string', 'Item missing valid sheetTitle');
    assert.ok(typeof structured.gid === 'string', 'Item missing valid gid');
    assert.ok(typeof structured.sheetUrl === 'string', 'Item missing valid sheetUrl');
    assert.ok(typeof structured.totalOperations === 'number', 'Item missing valid totalOperations');
    assert.ok(Array.isArray(structured.operationResults), 'Item missing valid operationResults array');
    assert.ok(structured.updatedDimensions, 'Item missing valid updatedDimensions');
    assert.ok(typeof structured.updatedDimensions.rows === 'number', 'Item missing valid updatedDimensions.rows');
    assert.ok(typeof structured.updatedDimensions.columns === 'number', 'Item missing valid updatedDimensions.columns');

    // Validate content array matches outputSchema requirements
    const content = response.content;
    assert.ok(Array.isArray(content), 'Response missing content array');
    assert.ok(content.length > 0, 'Content array is empty');
    const firstContent = content[0];
    assert.ok(firstContent, 'First content item missing');
    assert.strictEqual(firstContent.type, 'text', 'Content item missing text type');
    if (firstContent.type === 'text') {
      assert.ok(typeof firstContent.text === 'string', 'Content item missing text field');
    }

    // Validate operation details
    assert.equal(structured.totalOperations, 4, 'Should have 4 operations');
    assert.equal(structured.operationResults.length, 4, 'Should have 4 operation results');

    // Operations should be in the order they were executed (optimal order)
    // The handler sorts them internally: delete -> insert -> append
    const ops = structured.operationResults;

    // First two operations should be deletes (highest priority)
    const op0 = ops[0];
    const op1 = ops[1];
    if (op0) {
      assert.equal(op0.operation, 'deleteDimension', 'First operation should be deleteDimension');
    }
    if (op1) {
      assert.equal(op1.operation, 'deleteDimension', 'Second operation should be deleteDimension');
    }

    // Verify the column delete operation details
    const colDeleteOp = ops.find((op: { dimension: string; operation: string }) => op.dimension === 'COLUMNS' && op.operation === 'deleteDimension');
    assert.ok(colDeleteOp, 'Should have column delete operation');
    assert.equal(colDeleteOp.startIndex, 1, 'Column delete start index should be 1');
    assert.equal(colDeleteOp.endIndex, 3, 'Column delete end index should be 3');
    assert.equal(colDeleteOp.affectedCount, 2, 'Should delete 2 columns (1-3 exclusive)');

    // Verify the row delete operation details
    const rowDeleteOp = ops.find((op: { dimension: string; operation: string }) => op.dimension === 'ROWS' && op.operation === 'deleteDimension');
    assert.ok(rowDeleteOp, 'Should have row delete operation');
    assert.equal(rowDeleteOp.affectedCount, 100, 'Should delete 100 rows');

    // Third operation should be the insert
    const op2 = ops[2];
    if (op2) {
      assert.equal(op2.operation, 'insertDimension', 'Third operation should be insertDimension');
      assert.equal(op2.dimension, 'COLUMNS', 'Insert operation should target COLUMNS');
      assert.equal(op2.affectedCount, 2, 'Should insert 2 columns');
    }

    // Fourth operation should be the append
    const op3 = ops[3];
    if (op3) {
      assert.equal(op3.operation, 'appendDimension', 'Fourth operation should be appendDimension');
      assert.equal(op3.dimension, 'ROWS', 'Append operation should target ROWS');
      assert.equal(op3.affectedCount, 1, 'Should append 1 row');
    }

    // Validate final dimensions
    // Rows: 1000 (default) - 100 (deleted) + 1 (appended) = 901
    assert.equal(structured.updatedDimensions.rows, 901, 'Should have 901 rows after all operations');
    // Columns: 26 (default) - 2 (deleted) + 2 (inserted) = 26
    assert.equal(structured.updatedDimensions.columns, 26, 'Should have 26 columns after all operations');
  });
}); // End describe block
