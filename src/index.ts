import { handleVersionHelp } from './setup/version-help.ts';
import type { ServerConfig } from './types.ts';

export { GOOGLE_SCOPE } from './constants.ts';
export * as mcp from './mcp/index.ts';
export * as schemas from './schemas/index.ts';
export * as setup from './setup/index.ts';
export * from './types.ts';

export async function startServer(config: ServerConfig): Promise<void> {
  // createHTTPServer/createStdioServer pull in @modelcontextprotocol/sdk and every mcp/tools/*.ts
  // (googleapis included); deferred so a --version/--help run never reaches them.
  const { createHTTPServer } = await import('./setup/http.ts');
  const { createStdioServer } = await import('./setup/stdio.ts');
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

  // Only parse config if no help/version flags. config.ts's own heavy imports (@mcp-z/oauth-google,
  // @mcp-z/server) are deferred internally, so this static-looking call stays cheap until here.
  const { createConfig } = await import('./setup/config.ts');
  const config = createConfig();
  await startServer(config);
}

if (process.argv[1] === new URL(import.meta.url).pathname) {
  main();
}
