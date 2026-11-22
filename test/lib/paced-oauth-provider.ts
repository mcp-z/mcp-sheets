import { LoopbackOAuthProvider } from '@mcp-z/oauth-google';
import pThrottle from 'p-throttle';

// Shared by all test clients. Fifty requests/minute leaves room below Sheets'
// sixty reads or writes/minute; provider-backed clients request fresh auth per call.
const interval = 1200;
let authorizations = 0;
let rawRequests = 0;
let spreadsheetCreates = 0;
const runEndHooks: Array<() => Promise<void>> = [];
const reserve = pThrottle({ limit: 1, interval, strict: true })(() => {});

/** Reserve the shared request slot for a client using a static access token. */
export async function reserveRawRequest(): Promise<void> {
  await reserve();
  rawRequests++;
}

export class PacedOAuthProvider extends LoopbackOAuthProvider {
  override async getAccessToken(accountId?: string): Promise<string> {
    await reserve();
    authorizations++;
    return super.getAccessToken(accountId);
  }
}

/** Count one spreadsheet created for the run, by the helper or by a tool under test. */
export function countSpreadsheetCreate(): void {
  spreadsheetCreates++;
}

/** Register work that runs once when the run ends, before the pacing line prints. */
export function onRunEnd(hook: () => Promise<void>): void {
  runEndHooks.push(hook);
}

/** Run the registered end-of-run hooks; every hook runs and the failures are thrown together. */
export async function runEndHooksAll(): Promise<void> {
  const results = await Promise.allSettled(runEndHooks.splice(0).map((hook) => hook()));
  const failures = results.filter((r): r is PromiseRejectedResult => r.status === 'rejected').map((r) => r.reason);
  if (failures.length > 0) throw new Error(`run cleanup failed: ${failures.map((f) => (f instanceof Error ? f.message : String(f))).join('; ')}`);
}

export function pacingSummary(): { authorizations: number; rawRequests: number; spreadsheetCreates: number; intervalMs: number } {
  return { authorizations, rawRequests, spreadsheetCreates, intervalMs: interval };
}
