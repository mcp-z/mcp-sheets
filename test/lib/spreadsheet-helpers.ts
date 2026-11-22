/**
 * Create a minimal test spreadsheet and return its id.
 * Keeps setup DRY for tests inside servers/mcp-sheets.
 */

import { drive as driveApi } from '@googleapis/drive';
import { auth as googleAuthPlus, sheets as sheetsApi } from '@googleapis/sheets';
import crypto from 'crypto';
import type { GoogleApiError, Logger } from '../../src/types.ts';
import createMiddlewareContext from './create-middleware-context.ts';
import { countSpreadsheetCreate, onRunEnd, reserveRawRequest } from './paced-oauth-provider.ts';
import { RunRegistry } from './run-registry.ts';

/** One id per test process; names every spreadsheet the run creates. */
export const RUN_ID = crypto.randomUUID();

// Every spreadsheet this run creates and has not yet deleted.
const registry = new RunRegistry();
let signalsInstalled = false;
let sharedId: Promise<string> | undefined;
let context: ReturnType<typeof createMiddlewareContext> | undefined;

function runContext() {
  context ??= createMiddlewareContext();
  return context;
}

function register(id: string, title: string): void {
  registry.add(id, title);
  if (signalsInstalled) return;
  signalsInstalled = true;
  onRunEnd(cleanupRun);
  const signals = ['SIGINT', 'SIGTERM'] as const;
  const handlers = signals.map((signal) => {
    const handler = async () => {
      for (let i = 0; i < signals.length; i++) process.removeListener(signals[i] as string, handlers[i] as () => void);
      try {
        await cleanupRun();
      } catch (error) {
        console.error('Run cleanup failed:', error);
      }
      process.kill(process.pid, signal);
    };
    return handler;
  });
  signals.forEach((signal, i) => process.once(signal, handlers[i] as () => void));
}

/** Register a spreadsheet a tool under test made. A 'create' counts toward the run's creates; a 'copy' does not. */
export function trackTestSpreadsheet(id: string, kind: 'create' | 'copy', title: string): void {
  if (kind === 'create') countSpreadsheetCreate();
  register(id, title);
}

/**
 * Delete every spreadsheet no delete has been attempted on. A spreadsheet whose delete already failed is not
 * tried again; the call throws naming each spreadsheet left behind.
 */
export async function cleanupRun(): Promise<void> {
  if (sharedId) await sharedId.catch(() => {});
  sharedId = undefined;
  if (registry.size === 0) return;
  const ctx = await runContext();
  const accessToken = await ctx.authProvider.getAccessToken(ctx.accountId);
  await registry.cleanup((id) => driveDelete(accessToken, id));
}

/**
 * The one spreadsheet shared by the run's specs; created on the first call.
 * Its default tab (gid 0, Sheet1) is never written or renamed by anyone: that keeps the columns-update gid-0 assertion exact.
 */
export function sharedSpreadsheet(): Promise<string> {
  sharedId ??= (async () => {
    const ctx = await runContext();
    return createTestSpreadsheet(await ctx.authProvider.getAccessToken(ctx.accountId), { title: `ci-shared-${RUN_ID}` });
  })();
  return sharedId;
}

export type TestTab = { gid: string; sheetId: number; title: string; a1: (range: string) => string };

/** Add tabs in one request with default grid properties; the tabs are addressed by gid afterwards. */
export async function addTestTabs(spreadsheetId: string, titles: string[]): Promise<TestTab[]> {
  const ctx = await runContext();
  const sheets = sheetsApi({ version: 'v4', auth: ctx.auth });
  const resp = await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: titles.map((title) => ({ addSheet: { properties: { title } } })) } });
  return titles.map((title, i) => {
    const sheetId = resp.data.replies?.[i]?.addSheet?.properties?.sheetId;
    if (typeof sheetId !== 'number') throw new Error(`addTestTabs: no sheet id for "${title}"`);
    return { gid: String(sheetId), sheetId, title, a1: (range: string) => `'${title.replace(/'/g, "''")}'!${range}` };
  });
}

/** Remove tabs in one request. A gid that does not exist fails the call. */
export async function deleteTestTabs(spreadsheetId: string, gids: Array<string | number>): Promise<void> {
  if (gids.length === 0) return;
  const ctx = await runContext();
  const sheets = sheetsApi({ version: 'v4', auth: ctx.auth });
  await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: gids.map((gid) => ({ deleteSheet: { sheetId: Number(gid) } })) } });
}

export async function createTestSpreadsheet(accessToken: string, opts: { title?: string } = {}): Promise<string> {
  const title = opts.title || `ci-test-spreadsheet-${Date.now()}`;

  const auth = new googleAuthPlus.OAuth2();
  auth.setCredentials({ access_token: accessToken });
  const sheets = sheetsApi({ version: 'v4', auth });
  await reserveRawRequest();
  const response = await sheets.spreadsheets.create({ requestBody: { properties: { title } } });
  const id = response.data.spreadsheetId;
  if (!id) throw new Error('createTestSpreadsheet: expected spreadsheet id');
  countSpreadsheetCreate();
  register(id, title);
  return id;
}

async function driveDelete(accessToken: string, fileId: string): Promise<void> {
  const auth = new googleAuthPlus.OAuth2();
  auth.setCredentials({ access_token: accessToken });
  const drive = driveApi({ version: 'v3', auth });
  await reserveRawRequest();
  await drive.files.delete({ fileId });
}

/**
 * Delete a test spreadsheet created with createTestSpreadsheet. The registry records the attempt; a failed
 * delete is never retried and is named when the run ends.
 * Throws on any error - close failures indicate test problems that need to be visible.
 */
export async function deleteTestSpreadsheet(accessToken: string, id: string, logger: Logger): Promise<void> {
  try {
    await registry.remove(id, (fileId) => driveDelete(accessToken, fileId));
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

  const auth = new googleAuthPlus.OAuth2();
  auth.setCredentials({ access_token: accessToken });
  const sheets = sheetsApi({ version: 'v4', auth });
  await reserveRawRequest();
  const resp = await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [{ addSheet: { properties: { title } } }] } });
  const sid = resp.data.replies?.[0]?.addSheet?.properties?.sheetId;
  return Number(sid);
}
