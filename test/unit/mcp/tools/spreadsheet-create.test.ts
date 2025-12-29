import type { Logger, LoopbackOAuthProvider } from '@mcp-z/oauth-google';
import assert from 'assert';
import createTool, { type Input, type Output } from '../../../../src/mcp/tools/spreadsheet-create.js';
import { createExtra, type TypedHandler } from '../../../lib/create-extra.js';
import createMiddlewareContext from '../../../lib/create-middleware-context.js';
import { deleteTestSpreadsheet } from '../../../lib/spreadsheet-helpers.js';

describe('spreadsheet-create', () => {
  // Shared instances for all tests
  let authProvider: LoopbackOAuthProvider;
  let accountId: string;
  let logger: Logger;
  let handler: TypedHandler<Input>;

  before(async () => {
    // Get middleware for tool creation and auth for close operations
    const middlewareContext = await createMiddlewareContext();
    authProvider = middlewareContext.authProvider;
    accountId = middlewareContext.accountId;
    logger = middlewareContext.logger;
    const middleware = middlewareContext.middleware;
    const tool = createTool();
    const wrappedTool = middleware.withToolAuth(tool);
    handler = wrappedTool.handler;
  });

  it('spreadsheet_create creates a spreadsheet and is deletable', async () => {
    let createdFileId: string | null = null;
    try {
      const title = `ci-create-${Date.now()}`;
      const res = await handler({ title: title }, createExtra());

      // Validate complete response structure according to outputSchema
      assert.ok(res, 'Handler returned no result');

      // Validate structuredContent.result matches outputSchema
      const structured = res.structuredContent?.result as Output | undefined;
      assert.ok(structured, 'Response missing structuredContent.result');
      assert.strictEqual(structured.type, 'success', 'Expected success result');
      if (structured.type === 'success') {
        assert.ok(typeof structured.id === 'string' && structured.id.length > 0, 'Item missing valid id');
        createdFileId = structured.id;
      }

      // Validate content array matches outputSchema requirements
      const content = res.content;
      assert.ok(Array.isArray(content), 'Response missing content array');
      assert.ok(content.length > 0, 'Content array is empty');
      const firstContent = content[0];
      assert.ok(firstContent, 'First content item missing');
      assert.strictEqual(firstContent.type, 'text', 'Content item missing text type');
      assert.ok(typeof firstContent.text === 'string', 'Content item missing text field');

      // leave createdFileId for deterministic close in finally
    } finally {
      if (createdFileId) {
        try {
          const accessToken = await authProvider.getAccessToken(accountId);
          await deleteTestSpreadsheet(accessToken, createdFileId, logger);
        } catch (cleanupError) {
          logger.error({ err: cleanupError }, 'Failed to close test spreadsheet');
        }
      }
    }
  });
});
