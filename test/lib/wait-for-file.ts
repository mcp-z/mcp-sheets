import { setTimeout as delay } from 'timers/promises';
import type { GoogleApiError } from '../../src/types.ts';

export default async function waitForFile(drive: { files: { list: (args: Record<string, unknown>) => Promise<{ data?: { files?: Array<{ id?: string; name?: string }> } }> } }, q: string, opts: { interval?: number; timeout?: number } = {}): Promise<Array<{ id?: string; name?: string }> | undefined> {
  const initialInterval = typeof opts.interval === 'number' ? opts.interval : 200;
  const timeout = typeof opts.timeout === 'number' ? opts.timeout : 8000;
  const maxInterval = 1000;
  const start = Date.now();
  let currentInterval = initialInterval;

  while (true) {
    if (Date.now() - start > timeout) throw new Error('waitForFile: timeout waiting for file');
    try {
      const list = await drive.files.list({ q, spaces: 'drive', fields: 'files(id, name)' });
      if (Array.isArray(list?.data?.files) && list.data?.files?.length > 0) return list.data?.files;
    } catch (error: unknown) {
      const err = error as GoogleApiError;
      // Only ignore expected API errors (404, transient network issues)
      if (err?.response?.status === 404 || err?.code === 'ENOTFOUND') {
        continue; // Expected during file propagation
      }
      // Re-throw unexpected errors for diagnosis
      throw error;
    }
    await delay(currentInterval);
    // Exponential backoff with cap
    currentInterval = Math.min(currentInterval * 1.5, maxInterval);
  }
}
