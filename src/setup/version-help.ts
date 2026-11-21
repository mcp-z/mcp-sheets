import * as fs from 'fs';
import moduleRoot from 'module-root-sync';
import * as path from 'path';
import * as url from 'url';
import { parseArgs } from 'util';

// Kept dependency-free (fs/path/url/module-root-sync only) so `--version`/`--help`
// resolve nothing beyond Node startup: config.ts's parseConfig pulls in @mcp-z/oauth-google and
// @mcp-z/server, which this path must not touch.
const pkg = JSON.parse(fs.readFileSync(path.join(moduleRoot(url.fileURLToPath(import.meta.url)), 'package.json'), 'utf-8'));

const HELP_TEXT = `
Usage: mcp-sheets [options]

MCP server for Google Sheets spreadsheet management with OAuth authentication.

Options:
  --version, -v          Show version number
  --help, -h             Show this help message
  --auth=<mode>          Authentication mode (default: loopback-oauth)
                         Modes: loopback-oauth, service-account, dcr
  --headless             Disable browser auto-open, return auth URL instead
  --redirect-uri=<uri>   OAuth redirect URI (default: ephemeral loopback)
  --dcr-store-uri=<uri>  DCR client storage URI (default: .mcp-z/dcr.json)
  --port=<port>          Enable HTTP transport on specified port
  --stdio                Enable stdio transport (default if no port)
  --log-level=<level>    Logging level (default: info)
  --resource-store-uri=<uri>    Resource store URI for CSV file storage (default: .mcp-z/mcp-sheets/files beside the nearest .mcp.json, else ~/.mcp-z/mcp-sheets/files)
  --base-url=<url>       Base URL for HTTP file serving (optional)

Environment Variables:
  GOOGLE_CLIENT_ID       OAuth client ID (REQUIRED)
  GOOGLE_CLIENT_SECRET   OAuth client secret (optional)
  AUTH_MODE              Default authentication mode (optional)
  HEADLESS               Disable browser auto-open (optional)
  DCR_STORE_URI          DCR storage URI (optional, same as --dcr-store-uri)
  TOKEN_STORE_URI        Token storage URI (optional)
  PORT                   Default HTTP port (optional)
  LOG_LEVEL              Default logging level (optional)
  RESOURCE_STORE_URI            Resource store URI (optional, file://)
  BASE_URL               Base URL for HTTP file serving (optional)

Configuration Files:
  Variables may also be set in the first of these files that exists:
    1. .mcp-z/mcp-sheets/.env beside the nearest .mcp.json (else ~/.mcp-z/mcp-sheets/.env)
    2. ~/.config/mcp-z/mcp-sheets/.env   (~ is your home directory, on every OS)
  Values already in the environment, including .mcp.json "env", take precedence.
  On macOS and Linux the file must be private (chmod 600), or the server will not start.

Storage Backends:
  TOKEN_STORE_URI and DCR_STORE_URI accept any keyv-registry protocol.
  file+shared:// (the default), file:// and memory:// work out of the box. Any other backend
  needs its adapter installed alongside this server:
    npm install -g @keyv/redis
    TOKEN_STORE_URI=redis://localhost:6379 mcp-sheets

OAuth Scopes:
  openid https://www.googleapis.com/auth/userinfo.profile https://www.googleapis.com/auth/userinfo.email https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive

Examples:
  mcp-sheets                           # Use default settings
  mcp-sheets --auth=service-account    # Use service account auth
  mcp-sheets --port=3000               # HTTP transport on port 3000
  mcp-sheets --resource-store-uri=file:///tmp/sheets    # Custom resource store URI
  GOOGLE_CLIENT_ID=xxx mcp-sheets      # Set client ID via env var
`.trim();

/** Package metadata read from package.json, resolved relative to this file. */
export function readPkg(): { name: string; version: string; repository?: string | { url?: string } } {
  return pkg;
}

/**
 * Handle --version/--help flags before config parsing.
 * These must work without requiring any configuration or heavy dependency.
 */
export function handleVersionHelp(args: string[]): { handled: boolean; output?: string } {
  const { values } = parseArgs({
    args,
    options: {
      version: { type: 'boolean', short: 'v' },
      help: { type: 'boolean', short: 'h' },
    },
    strict: false,
    allowPositionals: true,
  });

  if (values.version) return { handled: true, output: pkg.version };
  if (values.help) return { handled: true, output: HELP_TEXT };
  return { handled: false };
}
