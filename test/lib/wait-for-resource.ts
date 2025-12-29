import { setTimeout as delay } from 'timers/promises';

export default async function waitForResource<T>(check: () => Promise<T>, opts: { interval?: number; timeout?: number } = {}): Promise<T> {
  const interval = typeof opts.interval === 'number' ? opts.interval : 500;
  const timeout = typeof opts.timeout === 'number' ? opts.timeout : 10000;
  const start = Date.now();
  while (true) {
    if (Date.now() - start > timeout) throw new Error('waitForResource: timeout waiting for resource');
    try {
      const ok = await check();
      if (ok) return ok;
    } catch {
      // ignore
    }
    await delay(interval);
  }
}
