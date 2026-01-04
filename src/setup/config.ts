import { parseDcrConfig, parseConfig as parseOAuthConfig } from '@mcp-z/oauth-google';
import { findConfigPath, parseConfig as parseTransportConfig } from '@mcp-z/server';
import * as fs from 'fs';
import moduleRoot from 'module-root-sync';
import { homedir } from 'os';
import * as path from 'path';
import * as url from 'url';
import { parseArgs } from 'util';
import { GOOGLE_SCOPE } from '../constants.ts';
import type { ServerConfig } from '../types.ts';

const pkg = JSON.parse(fs.readFileSync(path.join(moduleRoot(url.fileURLToPath(import.meta.url)), 'package.json'), 'utf-8'));

const HELP_TEXT = `
Usage: mcp-sheets [options]

MCP server for Google Sheets spreadsheet management with OAuth authentication.

Options:
  --version              Show version number
  --help                 Show this help message
  --auth=<mode>          Authentication mode (default: loopback-oauth)
                         Modes: loopback-oauth, service-account, dcr
  --headless             Disable browser auto-open, return auth URL instead
  --redirect-uri=<uri>   OAuth redirect URI (default: ephemeral loopback)
  --dcr-mode=<mode>      DCR mode (self-hosted or external, default: self-hosted)
  --dcr-verify-url=<url> External verification endpoint (required for external mode)
  --dcr-store-uri=<uri>  DCR client storage URI (required for self-hosted mode)
  --port=<port>          Enable HTTP transport on specified port
  --stdio                Enable stdio transport (default if no port)
  --log-level=<level>    Logging level (default: info)
  --resource-store-uri=<uri>    Resource store URI for CSV file storage (default: file://~/.mcp-z/mcp-sheets/files)
  --base-url=<url>       Base URL for HTTP file serving (optional)

Environment Variables:
  GOOGLE_CLIENT_ID       OAuth client ID (REQUIRED)
  GOOGLE_CLIENT_SECRET   OAuth client secret (optional)
  AUTH_MODE              Default authentication mode (optional)
  HEADLESS               Disable browser auto-open (optional)
  DCR_MODE               DCR mode (optional, same format as --dcr-mode)
  DCR_VERIFY_URL         External verification URL (optional, same as --dcr-verify-url)
  DCR_STORE_URI          DCR storage URI (optional, same as --dcr-store-uri)
  TOKEN_STORE_URI        Token storage URI (optional)
  PORT                   Default HTTP port (optional)
  LOG_LEVEL              Default logging level (optional)
  RESOURCE_STORE_URI            Resource store URI (optional, file://)
  BASE_URL               Base URL for HTTP file serving (optional)

OAuth Scopes:
  openid https://www.googleapis.com/auth/userinfo.profile https://www.googleapis.com/auth/userinfo.email https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive

Examples:
  mcp-sheets                           # Use default settings
  mcp-sheets --auth=service-account    # Use service account auth
  mcp-sheets --port=3000               # HTTP transport on port 3000
  mcp-sheets --resource-store-uri=file:///tmp/sheets    # Custom resource store URI
  GOOGLE_CLIENT_ID=xxx mcp-sheets      # Set client ID via env var
`.trim();

/**
 * Handle --version and --help flags before config parsing.
 * These should work without requiring any configuration.
 */
export function handleVersionHelp(args: string[]): { handled: boolean; output?: string } {
  const { values } = parseArgs({
    args,
    options: {
      version: { type: 'boolean' },
      help: { type: 'boolean' },
    },
    strict: false,
  });

  if (values.version) return { handled: true, output: pkg.version };
  if (values.help) return { handled: true, output: HELP_TEXT };
  return { handled: false };
}

/**
 * Parse Sheets server configuration from CLI arguments and environment.
 *
 * CLI Arguments (all optional):
 * - --auth=<mode>          Authentication mode (default: loopback-oauth)
 *                          Modes: loopback-oauth, service-account, dcr
 * - --headless             Disable browser auto-open, return auth URL instead
 * - --redirect-uri=<uri>   OAuth redirect URI (default: ephemeral loopback)
 * - --dcr-mode=<mode>      DCR mode (self-hosted or external, default: self-hosted)
 * - --dcr-verify-url=<url> External verification endpoint (required for external mode)
 * - --dcr-store-uri=<uri>  DCR client storage URI (required for self-hosted mode)
 * - --port=<port>          Enable HTTP transport on specified port
 * - --stdio                Enable stdio transport (default if no port)
 * - --log-level=<level>    Logging level (default: info)
 * - --resource-store-uri=<uri>    Resource store URI for CSV file storage (default: file://~/.mcp-z/mcp-sheets/files)
 * - --base-url=<url>       Base URL for HTTP file serving (optional)
 *
 * Environment Variables:
 * - GOOGLE_CLIENT_ID       OAuth client ID (REQUIRED)
 * - GOOGLE_CLIENT_SECRET   OAuth client secret (optional)
 * - AUTH_MODE              Default authentication mode (optional)
 * - HEADLESS               Disable browser auto-open (optional)
 * - DCR_MODE               DCR mode (optional, same format as --dcr-mode)
 * - DCR_VERIFY_URL         External verification URL (optional, same as --dcr-verify-url)
 * - DCR_STORE_URI          DCR storage URI (optional, same as --dcr-store-uri)
 * - TOKEN_STORE_URI        Token storage URI (optional)
 * - PORT                   Default HTTP port (optional)
 * - LOG_LEVEL              Default logging level (optional)
 * - RESOURCE_STORE_URI            Resource store URI (optional, file://)
 * - BASE_URL               Base URL for HTTP file serving (optional)
 *
 * OAuth Scopes (from constants.ts):
 * openid https://www.googleapis.com/auth/userinfo.profile https://www.googleapis.com/auth/userinfo.email https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive
 */
export function parseConfig(args: string[], env: Record<string, string | undefined>): ServerConfig {
  const transportConfig = parseTransportConfig(args, env);
  const oauthConfig = parseOAuthConfig(args, env);

  // Parse DCR configuration if DCR mode is enabled
  const dcrConfig = oauthConfig.auth === 'dcr' ? parseDcrConfig(args, env, GOOGLE_SCOPE) : undefined;

  // Parse application-level config (LOG_LEVEL, RESOURCE_STORE_URI, BASE_URL)
  const { values } = parseArgs({
    args,
    options: {
      'log-level': { type: 'string' },
      'base-url': { type: 'string' },
      'resource-store-uri': { type: 'string' },
    },
    strict: false, // Allow other arguments
    allowPositionals: true,
  });

  const name = pkg.name.replace(/^@[^/]+\//, '');
  // Parse repository URL from package.json, stripping git+ prefix and .git suffix
  const rawRepoUrl = typeof pkg.repository === 'object' ? pkg.repository.url : pkg.repository;
  const repositoryUrl = rawRepoUrl?.replace(/^git\+/, '').replace(/\.git$/, '') ?? `https://github.com/mcp-z/${name}`;
  let rootDir = homedir();
  try {
    const configPath = findConfigPath({ config: '.mcp.json', cwd: process.cwd(), stopDir: homedir() });
    rootDir = path.dirname(configPath);
  } catch {
    rootDir = homedir();
  }
  const baseDir = path.join(rootDir, '.mcp-z');
  const cliLogLevel = typeof values['log-level'] === 'string' ? values['log-level'] : undefined;
  const envLogLevel = env.LOG_LEVEL;
  const logLevel = cliLogLevel ?? envLogLevel ?? 'info';

  // Parse storage configuration
  const cliResourceStoreUri = typeof values['resource-store-uri'] === 'string' ? values['resource-store-uri'] : undefined;
  const envResourceStoreUri = env.RESOURCE_STORE_URI;
  const defaultResourceStorePath = path.join(baseDir, name, 'files');
  const resourceStoreUri = normalizeResourceStoreUri(cliResourceStoreUri ?? envResourceStoreUri ?? defaultResourceStorePath);

  const cliBaseUrl = typeof values['base-url'] === 'string' ? values['base-url'] : undefined;
  const envBaseUrl = env.BASE_URL;
  const baseUrl = cliBaseUrl ?? envBaseUrl;

  // Combine configs
  return {
    ...oauthConfig, // Includes clientId, auth, headless, redirectUri
    transport: transportConfig.transport,
    logLevel,
    baseDir,
    name,
    version: pkg.version,
    repositoryUrl,
    resourceStoreUri,
    ...(baseUrl && { baseUrl }),
    ...(dcrConfig && { dcrConfig }),
  };
}

function normalizeResourceStoreUri(resourceStoreUri: string): string {
  const filePrefix = 'file://';
  if (resourceStoreUri.startsWith(filePrefix)) {
    const rawPath = resourceStoreUri.slice(filePrefix.length);
    const expandedPath = rawPath.startsWith('~') ? rawPath.replace(/^~/, homedir()) : rawPath;
    return `${filePrefix}${path.resolve(expandedPath)}`;
  }

  if (resourceStoreUri.includes('://')) return resourceStoreUri;

  const expandedPath = resourceStoreUri.startsWith('~') ? resourceStoreUri.replace(/^~/, homedir()) : resourceStoreUri;
  return `${filePrefix}${path.resolve(expandedPath)}`;
}

/**
 * Build production configuration from process globals.
 * Entry point for production server.
 */
export function createConfig(): ServerConfig {
  return parseConfig(process.argv, process.env);
}
