import type { EnrichedExtra } from '@mcp-z/oauth-google';
import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js';
import pino from 'pino';

/**
 * Typed handler signature for test files
 * Use with tool's Input type: `let handler: TypedHandler<Input>;`
 */
export type TypedHandler<I> = (input: I, extra: EnrichedExtra) => Promise<CallToolResult>;

/**
 * Create EnrichedExtra for testing
 *
 * In production, the middleware automatically creates and injects authContext and logger.
 * In tests, we call handlers directly, so we need to provide it ourselves.
 *
 * Note: The auth and logger here are just placeholders - the real auth/logger come from
 * the middleware wrapper created in create-middleware-context.ts
 */
export function createExtra(): EnrichedExtra {
  return {
    signal: new AbortController().signal,
    requestId: 'test-request-id',
    sendNotification: async () => {},
    sendRequest: async () => ({}) as unknown,
    // Middleware injects these - placeholders for type compatibility
    authContext: {
      auth: {} as unknown, // Placeholder auth client
      accountId: 'test-account',
    },
    logger: pino({ level: 'silent' }),
  } as EnrichedExtra;
}
