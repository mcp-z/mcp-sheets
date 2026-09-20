import { fileURLToPath } from 'node:url';
import { loadEnv } from 'tsds-lib';

const envPath = fileURLToPath(new URL('../../.env.test', import.meta.url));
const { error } = loadEnv({ path: envPath });
if (error && (error as NodeJS.ErrnoException).code !== 'ENOENT') throw error;

/**
 * Read a variable the suite cannot run without, failing immediately and by name.
 *
 * Use instead of process.env.VAR || 'default': a default lets the suite pass against
 * invalid configuration and hides the real fault.
 */
export function requiredEnv(key: string): string {
  const value = process.env[key];
  if (!value) throw new Error(`Environment variable ${key} is required. Add it to .env.test (see .env.test.example).`);
  return value;
}
