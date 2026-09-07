/**
 * Loads `.env.test` from the package root into `process.env`, and exports `requiredEnv`.
 *
 * Must be the first import in every test file: imports evaluate in order, so anything
 * imported ahead of this reads an unpopulated `process.env`.
 *
 * Parsed here rather than with `process.loadEnvFile`, which needs Node 20.12 and is
 * undefined at every engines floor these packages declare (>=16, >=18, >=20), so
 * `test:engines` would fail on the floor leg. Semantics match it: a variable already
 * present in the environment wins, and no dependency is used.
 */

import fs from 'fs';
import path from 'path';

function parse(source: string): Record<string, string> {
  const values: Record<string, string> = {};

  for (const line of source.split('\n')) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith('#')) continue;

    const separator = trimmed.indexOf('=');
    if (separator <= 0) continue;

    const key = trimmed.slice(0, separator).trim();
    let value = trimmed.slice(separator + 1).trim();
    const quote = value[0];
    if (value.length >= 2 && (quote === '"' || quote === "'") && value.endsWith(quote)) value = value.slice(1, -1);

    values[key] = value;
  }

  return values;
}

try {
  const values = parse(fs.readFileSync(path.resolve(process.cwd(), '.env.test'), 'utf8'));
  for (const key of Object.keys(values)) {
    if (process.env[key] === undefined) process.env[key] = values[key];
  }
} catch (error) {
  // A package with no .env.test is valid; anything else is a real failure.
  if ((error as NodeJS.ErrnoException).code !== 'ENOENT') throw error;
}

/**
 * Read a variable the suite cannot run without, failing immediately and by name.
 *
 * Use instead of `process.env.VAR || 'default'`: a default lets the suite pass against
 * invalid configuration and hides the real fault.
 */
export function requiredEnv(key: string): string {
  const value = process.env[key];
  if (!value) throw new Error(`Environment variable ${key} is required. Add it to .env.test (see .env.test.example).`);
  return value;
}
