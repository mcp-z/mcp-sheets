import type { Logger, LoopbackOAuthProvider } from '@mcp-z/oauth-google';
import assert from 'assert';
import crypto from 'crypto';
import fs from 'fs/promises';
import type { OAuth2Client } from 'google-auth-library';
import { google } from 'googleapis';
import * as path from 'path';
import createTool, { type Input, type Output } from '../../../../src/mcp/tools/sheet-create.js';
import { createExtra, type TypedHandler } from '../../../lib/create-extra.js';
import createMiddlewareContext from '../../../lib/create-middleware-context.js';
import { createTestSpreadsheet, deleteTestSpreadsheet } from '../../../lib/spreadsheet-helpers.js';

// Shared test resources
let sharedSpreadsheetId: string;
let auth: OAuth2Client;
let authProvider: LoopbackOAuthProvider;
let logger: Logger;
let accountId: string;
let handler: TypedHandler<Input>;
let tmpDir: string;

before(async () => {
  try {
    // Create temporary directory
    tmpDir = path.join('.tmp', `sheet-create-tests-${crypto.randomUUID()}`);
    await fs.mkdir(tmpDir, { recursive: true });

    // Get middleware for tool creation and auth for close operations
    const middlewareContext = await createMiddlewareContext();
    auth = middlewareContext.auth;
    logger = middlewareContext.logger;
    const middleware = middlewareContext.middleware;
    const tool = createTool();
    const wrappedTool = middleware.withToolAuth(tool);
    handler = wrappedTool.handler;
    authProvider = middlewareContext.authProvider;
    accountId = middlewareContext.accountId;

    // Create shared spreadsheet for all tests
    const title = `ci-sheet-create-tests-${Date.now()}`;
    const accessToken = await authProvider.getAccessToken(accountId);
    sharedSpreadsheetId = await createTestSpreadsheet(accessToken, { title });
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

it('sheet_create adds a sheet (tab) to a spreadsheet and is removable', async () => {
  // Call sheet_create to add a new tab to shared spreadsheet
  const sheetTitle = `tab-${Date.now()}`;
  const res = await handler({ id: sharedSpreadsheetId, sheetTitle }, createExtra());
  assert.ok(res && res.structuredContent && res.content, 'missing structured result for sheet_create');
  const branch = res.structuredContent?.result as Output | undefined;
  assert.ok(branch, 'missing structured result for sheet_create');
  assert.equal(branch.type, 'success');
  if (branch.type === 'success') {
    assert.equal(branch.id, sharedSpreadsheetId, 'sheet_create did not return spreadsheet id');
    assert.ok(branch.gid, 'sheet_create should return gid');
    assert.equal(typeof branch.gid, 'string', 'gid should be a string');

    // Verify sheet exists via Sheets API
    const client = google.sheets({ version: 'v4', auth });
    const info = await client.spreadsheets.get({ spreadsheetId: sharedSpreadsheetId, fields: 'sheets.properties' });
    const exists = Array.isArray(info.data.sheets) && info.data.sheets.some((s) => s?.properties?.title === sheetTitle);
    assert.ok(exists, 'Expected sheet tab to exist after sheet_create');
  }
});
