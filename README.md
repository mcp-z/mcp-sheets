# @mcp-z/mcp-sheets

Docs: https://mcp-z.github.io/mcp-sheets
Google Sheets MCP server for reading, writing, and formatting spreadsheets.

## Common uses

- Find spreadsheets and sheets
- Append and update data
- Apply formatting, validation, and charts

## Transports

MCP supports stdio and HTTP.

**Stdio**
```json
{
  "mcpServers": {
    "sheets": {
      "command": "npx",
      "args": ["-y", "@mcp-z/mcp-sheets"]
    }
  }
}
```

**HTTP**
```json
{
  "mcpServers": {
    "sheets": {
      "type": "http",
      "url": "http://localhost:9004/mcp",
      "start": {
        "command": "npx",
        "args": ["-y", "@mcp-z/mcp-sheets", "--port=9004"]
      }
    }
  }
}
```

`start` is an extension used by `npx @mcp-z/cli up` to launch HTTP servers for you.

## Create a Google Cloud app

1. Go to [Google Cloud Console](https://console.cloud.google.com/).
2. Create or select a project.
3. Enable the Google Sheets API.
4. Create OAuth 2.0 credentials (Desktop app).
5. Copy the Client ID and Client Secret.

## OAuth modes

Configure via environment variables or the `env` block in `.mcp.json`. See `server.json` for the full list of options.

### Loopback OAuth (default)

Environment variables:

```bash
GOOGLE_CLIENT_ID=your-client-id
GOOGLE_CLIENT_SECRET=your-client-secret
```

Example:
```json
{
  "mcpServers": {
    "sheets": {
      "command": "npx",
      "args": ["-y", "@mcp-z/mcp-sheets"],
      "env": {
        "GOOGLE_CLIENT_ID": "your-client-id",
        "GOOGLE_CLIENT_SECRET": "your-client-secret"
      }
    }
  }
}
```

### Service account

Environment variables:

```bash
AUTH_MODE=service-account
GOOGLE_SERVICE_ACCOUNT_KEY_FILE=/path/to/service-account.json
```

Example:
```json
{
  "mcpServers": {
    "sheets": {
      "command": "npx",
      "args": ["-y", "@mcp-z/mcp-sheets", "--auth=service-account"],
      "env": {
        "GOOGLE_SERVICE_ACCOUNT_KEY_FILE": "/path/to/service-account.json"
      }
    }
  }
}
```

### DCR (self-hosted)

HTTP only. Requires a public base URL.

```json
{
  "mcpServers": {
    "sheets-dcr": {
      "command": "npx",
      "args": [
        "-y",
        "@mcp-z/mcp-sheets",
        "--auth=dcr",
        "--port=3456",
        "--base-url=https://oauth.example.com"
      ],
      "env": {
        "GOOGLE_CLIENT_ID": "your-client-id",
        "GOOGLE_CLIENT_SECRET": "your-client-secret"
      }
    }
  }
}
```

## How to use

```bash
# List tools
mcp-z inspect --servers sheets --tools

# Find a spreadsheet
mcp-z call sheets spreadsheet-find '{"spreadsheetRef":"Quarterly Report"}'
```

## Tools

1. cells-format
2. chart-create
3. columns-get
4. columns-update
5. csv-get-columns
6. dimensions-batch-update
7. dimensions-move
8. rows-append
9. rows-csv-append
10. rows-get
11. sheet-copy
12. sheet-copy-to
13. sheet-create
14. sheet-delete
15. sheet-find
16. sheet-rename
17. spreadsheet-copy
18. spreadsheet-create
19. spreadsheet-find
20. spreadsheet-rename
21. validation-set
22. values-batch-update
23. values-clear
24. values-csv-update
25. values-replace
26. values-search

## Resources

1. spreadsheet

## Prompts

1. a1-notation

## Configuration reference

See `server.json` for all supported environment variables, CLI arguments, and defaults.
