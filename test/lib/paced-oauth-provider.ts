import { LoopbackOAuthProvider } from '@mcp-z/oauth-google';
import pThrottle from 'p-throttle';

// Shared by all test clients. Fifty requests/minute leaves room below Sheets'
// sixty reads or writes/minute; provider-backed clients request fresh auth per call.
const interval = 1200;
let authorizations = 0;
let rawRequests = 0;
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

export function pacingSummary(): { authorizations: number; rawRequests: number; intervalMs: number } {
  return { authorizations, rawRequests, intervalMs: interval };
}
