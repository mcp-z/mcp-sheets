/**
 * Create a minimal test spreadsheet and return its id.
 * Keeps setup DRY for tests inside servers/mcp-sheets.
 */

import { OAuth2Client } from 'google-auth-library';
import { google } from 'googleapis';
import type { GoogleApiError, Logger } from '../../src/types.js';

export async function createTestSpreadsheet(accessToken: string, opts: { title?: string } = {}): Promise<string> {
  const title = opts.title || `ci-test-spreadsheet-${Date.now()}`;

  const auth = new OAuth2Client();
  auth.setCredentials({ access_token: accessToken });
  const sheets = google.sheets({ version: 'v4', auth });
  const response = await sheets.spreadsheets.create({ requestBody: { properties: { title } } });
  const id = response.data.spreadsheetId;
  if (!id) throw new Error('createTestSpreadsheet: expected spreadsheet id');
  return id;
}

/**
 * Delete a test spreadsheet created with createTestSpreadsheet.
 * Throws on any error - close failures indicate test problems that need to be visible.
 */
export async function deleteTestSpreadsheet(accessToken: string, id: string, logger: Logger): Promise<void> {
  try {
    const auth = new OAuth2Client();
    auth.setCredentials({ access_token: accessToken });
    const drive = google.drive({ version: 'v3', auth });
    await drive.files.delete({ fileId: id });
    logger.debug('Test spreadsheet close successful', { spreadsheetId: id });
  } catch (e: unknown) {
    const error = e as GoogleApiError;
    logger.error('Test spreadsheet close failed', {
      spreadsheetId: id,
      error: e instanceof Error ? e.message : String(e),
      status: error?.status || error?.statusCode,
      code: error?.code,
    });
    throw e; // Always throw - if we're deleting it, it should exist
  }
}

/**
 * Create a test sheet (tab) within an existing spreadsheet.
 * Returns the sheet id.
 */
export async function createTestSheet(accessToken: string, spreadsheetId: string, opts: { title?: string } = {}): Promise<number> {
  const title = opts.title || `ci-test-sheet-${Date.now()}`;

  const auth = new OAuth2Client();
  auth.setCredentials({ access_token: accessToken });
  const sheets = google.sheets({ version: 'v4', auth });
  const resp = await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [{ addSheet: { properties: { title } } }] } });
  const sid = resp.data.replies?.[0]?.addSheet?.properties?.sheetId;
  return Number(sid);
}
