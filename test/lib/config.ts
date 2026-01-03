import { parseConfig } from '../../src/setup/config.ts';
import type { ServerConfig } from '../../src/types.ts';

export function createConfig(args?: string[], env?: Record<string, string | undefined>): ServerConfig {
  return parseConfig(args ?? ['--headless'], env ?? process.env);
}
