import { parseConfig } from '../../src/setup/config.js';
import type { ServerConfig } from '../../src/types.js';

export function createConfig(args?: string[], env?: Record<string, string | undefined>): ServerConfig {
  return parseConfig(args ?? ['--headless'], env ?? process.env);
}
