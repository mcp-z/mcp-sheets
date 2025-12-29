import type { PromptModule } from '@mcp-z/server';
import type { RequestHandlerExtra } from '@modelcontextprotocol/sdk/shared/protocol.js';
import type { ServerNotification, ServerRequest } from '@modelcontextprotocol/sdk/types.js';

export default function createPrompt() {
  const config = {
    description: 'Reference guide for Google Sheets A1 notation syntax',
  };

  const handler = async (_args: { [x: string]: unknown }, _extra: RequestHandlerExtra<ServerRequest, ServerNotification>) => {
    return {
      messages: [
        {
          role: 'user' as const,
          content: {
            type: 'text' as const,
            text: `# Google Sheets A1 Notation Reference

A1 notation identifies cells and ranges in Google Sheets.

## Cell References
- \`B5\` - Single cell at column B, row 5
- \`A1\` - Top-left cell

## Range References
- \`A1:D10\` - Rectangle from A1 to D10
- \`A5:D5\` - Single row (columns A-D of row 5)
- \`B1:B10\` - Single column segment (rows 1-10 of column B)

## Full Row/Column References
- \`5:5\` - Entire row 5
- \`B:B\` - Entire column B
- \`A:D\` - Columns A through D (all rows)
- \`1:10\` - Rows 1 through 10 (all columns)

## Common Patterns
| Goal | Notation |
|------|----------|
| Get one cell | \`B5\` |
| Get a row | \`A5:Z5\` or \`5:5\` |
| Get a column | \`B:B\` or \`B1:B1000\` |
| Get a data table | \`A1:F100\` |
| Get headers | \`1:1\` or \`A1:Z1\` |`,
          },
        },
      ],
    };
  };

  return {
    name: 'a1-notation',
    config,
    handler,
  } satisfies PromptModule;
}
