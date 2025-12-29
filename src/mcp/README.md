# MCP Components: Unified Authentication Pattern

Docs: https://mcp-z.github.io/mcp-sheets
This directory contains MCP component implementations (tools, resources, prompts) for Google Sheets.

## Unified Middleware Pattern

All MCP components (tools, resources, and prompts) receive the `RequestHandlerExtra` parameter from the MCP SDK, allowing consistent middleware-based authentication across all component types.

### Tools

**Handler Signature:**
```typescript
type ToolHandler<A extends AnyArgs = AnyArgs> = (
  argsOrExtra: A | RequestHandlerExtra,
  maybeExtra?: RequestHandlerExtra
) => Promise<CallToolResult>;
```

**Implementation:**
```typescript
export default function createTool() {
  const handler = async (args: In, extra: EnrichedExtra): Promise<CallToolResult> => {
    // Middleware enriches extra with authContext and logger
    const { auth } = extra.authContext;
    const { logger } = extra;

    const sheets = google.sheets({ version: 'v4', auth });
    // ...
  };

  return {
    name: 'rows-get',
    config,
    handler,
  } satisfies ToolModule;
}
```

### Resources

**Handler Signature:**
```typescript
type ReadResourceTemplateCallback = (
  uri: URL,
  variables: Variables,
  extra: RequestHandlerExtra  // ✅ Third parameter
) => ReadResourceResult | Promise<ReadResourceResult>;
```

**Implementation:**
```typescript
export default function createResource(): ResourceModule {
  const handler = async (uri: URL, vars: Record<string, string | string[]>, extra: EnrichedExtra): Promise<ReadResourceResult> => {
    // Middleware enriches extra with authContext and logger
    const { auth } = extra.authContext;
    const { logger } = extra;

    const sheets = google.sheets({ version: 'v4', auth });
    const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId });

    return {
      contents: [{
        uri: uri.href,
        mimeType: 'application/json',
        text: JSON.stringify(spreadsheet.data)
      }]
    };
  };

  return {
    name: 'spreadsheet',
    template,
    config,
    handler,
  };
}
```

### Prompts

**Handler Signature:**
```typescript
type PromptHandler = (
  args: { [x: string]: unknown },
  extra: RequestHandlerExtra  // ✅ Second parameter
) => Promise<GetPromptResult>;
```

**Implementation:**
```typescript
export default function createPrompt(): PromptModule {
  const handler = async (args: { [x: string]: unknown }, extra: RequestHandlerExtra) => {
    const { data, goal } = argsSchema.parse(args);
    return {
      messages: [
        { role: 'system', content: { type: 'text', text: '...' } },
        { role: 'user', content: { type: 'text', text: `...` } },
      ],
    };
  };

  return { name: 'data-analyze-data', config, handler };
}
```

## Registration Pattern

All components follow the same pattern:

```typescript
const { middleware: authMiddleware } = oauthAdapters;

// All components wrapped with auth middleware using same pattern
const tools = Object.values(toolFactories)
  .map((f) => f())
  .map(authMiddleware.withToolAuth);

const resources = Object.values(resourceFactories)
  .map((x) => x())
  .map(authMiddleware.withResourceAuth);

const prompts = Object.values(promptFactories)
  .map((x) => x())
  .map(authMiddleware.withPromptAuth);

registerTools(mcpServer, tools);
registerResources(mcpServer, resources);
registerPrompts(mcpServer, prompts);
```

## Key Benefits

1. **Consistent Pattern**: All components use middleware for cross-cutting concerns
2. **Type Safety**: Middleware enriches `extra` with `authContext` and `logger`
3. **Lazy Authentication**: Auth only happens when requests come in
4. **DCR Support**: Middleware handles both stored accounts and DCR bearer tokens
5. **Error Handling**: Middleware converts auth errors to component-specific error formats

## See Also

- @mcp-z/oauth-google
