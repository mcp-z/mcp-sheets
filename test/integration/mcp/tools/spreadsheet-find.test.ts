import '../../../lib/env-loader.ts';
import { mcp } from '@mcp-z/mcp-sheets';
import type { EnrichedExtra } from '@mcp-z/oauth-google';
import type { CallToolResult } from '@mcp-z/server';
import assert from 'assert';
import crypto from 'crypto';
import fs from 'fs/promises';
import * as path from 'path';
import type { Input, Output } from '../../../../src/mcp/tools/spreadsheet-find.ts';
import { createExtra } from '../../../lib/create-extra.ts';
import createMiddlewareContext from '../../../lib/create-middleware-context.ts';
import { sharedSpreadsheet } from '../../../lib/spreadsheet-helpers.ts';

describe('spreadsheet-find', () => {
  // Shared test resources
  let sharedSpreadsheetId: string;
  let handler: (input: Input, extra: EnrichedExtra) => Promise<CallToolResult>;
  let tmpDir: string;

  before(async () => {
    try {
      // Create temporary directory
      tmpDir = path.join('.tmp', `spreadsheet-find-tests-${crypto.randomUUID()}`);
      await fs.mkdir(tmpDir, { recursive: true });

      // Get middleware for tool creation
      const middlewareContext = await createMiddlewareContext();
      const middleware = middlewareContext.middleware;
      const tool = mcp.toolFactories.spreadsheetFind();
      const wrappedTool = middleware.withToolAuth(tool);
      handler = wrappedTool.handler;

      // Create shared spreadsheet for all tests
      sharedSpreadsheetId = await sharedSpreadsheet();
    } catch (error) {
      console.error('Failed to initialize test resources:', error);
      throw error;
    }
  });

  after(async () => {
    // Cleanup resources - fail fast on errors
    if (tmpDir) await fs.rm(tmpDir, { recursive: true, force: true });
  });

  it('sheets-spreadsheet-find locates by id', async () => {
    // By id using shared spreadsheet
    const byId = await handler({ spreadsheetRef: sharedSpreadsheetId }, createExtra());
    const byIdStructured = (byId.structuredContent as { result?: unknown } | undefined)?.result as Output | undefined;
    if (byIdStructured?.type !== 'success') {
      assert.fail('Spreadsheet find operation failed');
    }
    assert.ok(Array.isArray(byIdStructured?.items), 'expected items array');
    const idMatch = byIdStructured?.items?.find((it: { id: string }) => it.id === sharedSpreadsheetId);
    assert.ok(idMatch, 'id not present in items');
    assert.ok(typeof idMatch.spreadsheetUrl === 'string', 'expected spreadsheetUrl on item');
    assert.ok(typeof idMatch.spreadsheetTitle === 'string', 'expected spreadsheetTitle on item');
    assert.ok(Array.isArray(idMatch.sheets), 'expected sheets array on item');

    // Skip name search test to avoid Drive indexing delays in CI
    // Name-based search relies on Google Drive's search index which updates with eventual consistency
  });
});
