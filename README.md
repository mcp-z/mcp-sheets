# @mcp-z/mcp-sheets

MCP server for Google Sheets integration with OAuth authentication, spreadsheet management, batch operations, and advanced formatting

Requires Node.js >=20. The examples use `npx`, included with npm, to run this server and `@mcp-z/cli`.

## Common uses

- Find spreadsheets and sheets
- Append and update data
- Apply formatting, validation, and charts

## Transports

MCP supports stdio and HTTP.

Both the 2025 and 2026-07-28 protocol revisions are served, over either transport, from the same
server. Your client negotiates whichever it speaks. A 2025 client keeps working with no change,
and support for it is not being dropped. The 2026-07-28 revision is stateless, so a client speaking
it sends no `initialize` handshake and carries no session id.

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

`start` is an extension used by `npx @mcp-z/cli up` to launch HTTP servers for you. The HTTP endpoint is `/mcp`.

## Create a Google Cloud app

1. Go to [Google Cloud Console](https://console.cloud.google.com/).
2. Create or select a project.
3. Enable the Google Sheets API.
4. Create OAuth 2.0 credentials (Desktop app).
5. Copy the Client ID and Client Secret.
6. Select the credential type that matches your transport:
   - For stdio, choose "Desktop app" under APIs & Services.
   - For HTTP, choose "Web application" and add your public `/oauth/callback` URL. Local HTTP uses the port configured with `--port` or `PORT`.
   - For local hosting, add `http://127.0.0.1` for the [ephemeral redirect URL](https://en.wikipedia.org/wiki/Ephemeral_port).
7. Enable OAuth2 [scopes](https://console.cloud.google.com/auth/scopes): openid https://www.googleapis.com/auth/userinfo.profile https://www.googleapis.com/auth/userinfo.email https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive
8. Add [test emails](https://console.cloud.google.com/auth/audience)

## OAuth modes

Configure via environment variables or the `env` block in `.mcp.json`. See `server.json` for the full list of options.

### Loopback OAuth (default)

Environment variables:

```bash
GOOGLE_CLIENT_ID=your-client-id
GOOGLE_CLIENT_SECRET=your-client-secret
```

Example (stdio) - Create .mcp.json:
```json
{
  "mcpServers": {
    "sheets": {
      "command": "npx",
      "args": ["-y", "@mcp-z/mcp-sheets"],
      "env": {
        "GOOGLE_CLIENT_ID": "your-client-id"
      }
    }
  }
}
```

Example (http) - Create .mcp.json:
```json
{
  "mcpServers": {
    "sheets": {
      "type": "http",
      "url": "http://localhost:3000/mcp",
      "start": {
        "command": "npx",
        "args": ["-y", "@mcp-z/mcp-sheets", "--port=3000"],
        "env": {
          "GOOGLE_CLIENT_ID": "your-client-id"
        }
      }
    }
  }
}
```

Local (default): omit REDIRECT_URI → ephemeral loopback. Cloud: set REDIRECT_URI to your public /oauth/callback and expose the service publicly.

Note: the `start` block is a helper in `npx @mcp-z/cli up` for starting an HTTP server from your `.mcp.json`. See [@mcp-z/cli](https://github.com/mcp-z/cli) for details.

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
npx -y @mcp-z/cli inspect --servers sheets --tools

# Find a spreadsheet
npx -y @mcp-z/cli call-tool sheets spreadsheet-find '{"spreadsheetRef":"Quarterly Report"}'
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
25. values-markdown-update
26. values-replace
27. values-search

## Resources

1. spreadsheet

## Prompts

1. a1-notation

## Configuration reference

See [`server.json`](https://github.com/mcp-z/mcp-sheets/blob/master/server.json) for all supported environment variables, CLI arguments, and defaults.

## Storage backends

OAuth tokens (`TOKEN_STORE_URI`) and DCR registrations (`DCR_STORE_URI`) are stored through [keyv-registry](https://www.npmjs.com/package/keyv-registry), which picks an adapter from the URI protocol.

`file://` (the default, under `~/.mcp-z/`) and `memory://` work with no extra setup.

Any other backend needs its adapter installed alongside this server. Adapters are resolved with `require()`, so a globally installed server finds a globally installed adapter:

```bash
npm install -g @mcp-z/mcp-sheets @keyv/redis

TOKEN_STORE_URI=redis://localhost:6379 mcp-sheets
```

A protocol whose adapter is missing fails at startup naming the package to install.

## Documentation

[API Docs](https://mcp-z.github.io/mcp-sheets)
