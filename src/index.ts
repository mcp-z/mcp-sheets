import { createConfig, handleVersionHelp } from './setup/config.ts';
import { createHTTPServer } from './setup/http.ts';
import { createStdioServer } from './setup/stdio.ts';
import type { ServerConfig } from './types.ts';

export { GOOGLE_SCOPE } from './constants.ts';
export * as mcp from './mcp/index.ts';
export * as schemas from './schemas/index.ts';
export * as setup from './setup/index.ts';
export * from './types.ts';

export async function startServer(config: ServerConfig): Promise<void> {
  const { logger, close } = config.transport.type === 'stdio' ? await createStdioServer(config) : await createHTTPServer(config);

  process.on('SIGINT', async () => {
    await close();
    process.exit(0);
  });

  logger.info(`Server started with ${config.transport.type} transport`);
  await new Promise(() => {});
}

export default async function main(): Promise<void> {
  // Check for help/version flags FIRST, before config parsing
  const versionHelpResult = handleVersionHelp(process.argv);
  if (versionHelpResult.handled) {
    console.log(versionHelpResult.output);
    process.exit(0);
  }

  // Only parse config if no help/version flags
  const config = createConfig();
  await startServer(config);
}

if (process.argv[1] === new URL(import.meta.url).pathname) {
  main();
}
