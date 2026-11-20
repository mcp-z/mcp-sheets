# Changelog

## [2.6.3] - 2026-11-20

### Security

- In `--auth=dcr` mode every MCP request now needs a valid bearer token. Before, `/mcp` answered `initialize`, `tools/list` and the other methods with 200 and no token, and only tool calls were refused. Requests without a valid token now get 401 with a `WWW-Authenticate` challenge naming the server's resource metadata; this changes behaviour for anyone who relied on the open endpoint. Clients that sign in through the server's OAuth flow are unaffected. Requires `@mcp-z/server` 2.4.0 and `@mcp-z/oauth-google` 2.4.3.

### Removed

- The external DCR mode settings `--dcr-mode=external`, `DCR_MODE=external`, `--dcr-verify-url` and `DCR_VERIFY_URL` are gone from the help text and `server.json`. They never had an effect: the server always ran its own sign-in endpoints. With `--auth=dcr`, a server started with external mode, or with a verify URL, now stops at start with an error saying so instead of ignoring the setting; in other auth modes the settings are still ignored and nothing stops. `--dcr-mode=self-hosted` and `DCR_MODE=self-hosted` still work.

### Fixed

- The `DCR_STORE_URI` entry in `server.json` no longer carries a fixed default with a wrong path, so `mcp-z` manifest generation does not write one into a configuration. The server's own default is `.mcp-z/dcr.json`.
- `--dcr-store-uri` is honoured. It was documented but ignored, so only `DCR_STORE_URI` chose where DCR registrations were stored; the flag now wins over the variable.
- Two mcp-z servers, or two sessions of one server, sharing `.mcp-z/tokens.json` no longer undo each other's sign-ins and sign-outs, and each sees the others' changes without restarting. Existing sign-ins are kept. The file is now written with mode 0600 on macOS and Linux. Upgrade every mcp-z server and the cli together: an older one running beside this one still rewrites the file. A damaged `tokens.json` or `dcr.json` now stops the server at start with an error naming the file, instead of being overwritten. Requires `keyv-registry` 3.1.0.
- A failed start in stdio mode no longer prints a second `sonic boom is not ready yet` error after the real one, and the log keeps its last lines.
- `cells-format`, `validation-set` and `values-markdown-update` report an unusable range or cell (for example `Sheet1!A1:B2`, `a1:b2`, `$A$1`, `A2:B`, or a column or row where a single cell is needed) as invalid parameters when every one failed, instead of an internal error. `cells-format` and `validation-set` do the same for a sheet that does not exist. Partial failures still return `success` with the failures listed.

## [2.6.2] - 2026-11-18

### Fixed

- `sheet-copy` fails with an error naming each failed copy when every copy failed, instead of returning `success` with only a `failures` list. Partial failures still return `success` with their failures.
- `sheet-delete` fails with an error naming each failed delete when every delete failed, instead of returning `success` with only a `failures` list. Partial failures still return `success` with their failures.
- `cells-format` fails with an error naming each failed range when every range failed, instead of returning `success` with `successCount` 0. Partial failures still return `success` with `failedRanges`.
- `validation-set` fails with an error naming each failed range when every range failed, instead of returning `success` with `successCount` 0. Partial failures still return `success` with `failedRanges`.
- `values-markdown-update` fails with an error naming each failed cell when every cell failed, instead of returning `success` with `successCount` 0. Partial failures still return `success` with `failedCells`.

## [2.6.1]

### Fixed

- `columns-update` writes updates to existing rows on the tab you named. Before, they could land on the first tab of the spreadsheet.
- `columns-update` updates to existing rows honour `valueInputOption`: `USER_ENTERED` parses numbers, dates and formulas, and `RAW` stores the text as given. Before, every updated cell was stored as text.
- `columns-update` updates of existing rows leave a cell unchanged when its value is `null` or missing, as the per-row path already did. Before, the batch path cleared it.
- `columns-update` fails with an error naming the failures when every write failed, instead of returning `success` with only an `errors` list. Partial failures still return `success` with their errors.

## [2.6.0]

### Changed

- On first start, sign-ins saved by earlier versions (`TOKEN_STORE_URI` and `DCR_STORE_URI`, default namespace only) are moved to the keyv 6 storage format, and the server logs "Moved N saved sign-in entries to the keyv 6 format" when it moves any. Nothing is moved on later starts.
- Upgrade servers that share one token file together: once one has moved the entries, an earlier version no longer finds them and asks you to sign in again.
- With a `sqlite://` token store opened by several servers at once on the first start after upgrading, one of them may fail once with "database is locked". Start it again.
- Old `duckdb://` data in `~/.keyv-duckdb/store.duckdb` is not moved (keyv 5 wrote it there whatever the URI said), so those sign-ins are lost and you sign in again. That file holds tokens; delete it.
- Storage adapters you install yourself for `TOKEN_STORE_URI` / `DCR_STORE_URI` (for example `@keyv/redis`) should be keyv 6 releases.

## [2.5.1] - 2026-11-16

### Changed

- The write, format and read tool descriptions explain native dates: a date is a serial number with a date format, `RAW` keeps a date-looking string as text, `USER_ENTERED` may parse it, and `values-batch-update` shows the write, format and read-back steps (read back with `UNFORMATTED_VALUE`).

### Fixed

- A failed header read (quota, sign-in, service or network error) is reported as an error instead of being treated as an empty sheet, so `rows-append`, `rows-csv-append` and `columns-update` no longer add a second header row after the data or append duplicate rows and still return success.

## [2.5.0] - 2026-11-15

### Added

- `cells-copy` copies a finite cell range to another range in the same spreadsheet (same or different tab) using native Sheets paste-special: `pasteType` (PASTE_NORMAL default, PASTE_VALUES, PASTE_FORMAT, PASTE_NO_BORDERS, PASTE_FORMULA, PASTE_DATA_VALIDATION, PASTE_CONDITIONAL_FORMATTING) and `pasteOrientation` (NORMAL or TRANSPOSE). The destination is one anchor cell or an exact multiple of the source; pastes outside the destination grid are rejected before any write.
- Private settings such as the OAuth client ID and secret can live in a file instead of the shell: the first file found of `<baseDir>/mcp-sheets/.env` (the `.mcp-z` folder beside the nearest `.mcp.json`, else in the home folder) and `~/.config/mcp-z/mcp-sheets/.env`. Non-empty environment values win over the file. On macOS and Linux a file readable by other accounts is refused. Requires `@mcp-z/server` 2.3.0.

### Fixed

- Duplicate sheet keys are reported as an error, and the key scan covers every row of the grid instead of stopping at the first gap; a failed key read is reported instead of ending the scan.

## [2.4.2] - 2026-10-12

### Fixed

- CSV column inspection waits for the input and parser to close before returning, allowing local files to be removed and recreated immediately after first-row inspection or malformed-input failures. Original CSV errors remain visible when cleanup also reports an error.

## [2.4.1] - 2026-10-11

### Fixed

- CSV column inspection now closes the input after reading the first row, allowing local files to be removed immediately and stopping downloads of the remaining HTTP body. Source and malformed CSV errors continue to report through the tool's existing error response.

## [2.4.0] - 2026-10-09

### Added

- `spreadsheet-metadata-get` inspects workbook properties and structure across all sheets before making changes, without fetching cell grid data.
- `cells-metadata-get` reads selected cell metadata to inspect notes and validation before editing. Reads require finite ranges on one sheet and exclude cell values, formulas, and direct formatting.

## [2.3.0] - 2026-09-07

### Changed

- Clients speaking the 2026-07-28 protocol revision now receive cache hints on list results: `tools/list`, `prompts/list`, `resources/templates/list` and `server/discover` carry a five-minute TTL and `cacheScope: 'public'`, while `resources/list` and `resources/read` stay `private` with no TTL because they vary by account. Previously every cacheable result used the SDK's conservative `ttlMs: 0` / `private` default, which caches nothing. 2025-era clients are unaffected — the fields do not exist on that revision.
- Tools, resources and prompts are now registered in name order, so `tools/list` returns the same order from every connection and a client can keep a cached catalog valid across a reconnect. The listed order differs from previous releases; no tool is added, removed or renamed.

## [2.2.0] - 2026-09-06

### Added

- Serves the 2026-07-28 MCP protocol revision alongside the existing 2025 revision, over both HTTP and stdio. A client speaking either one reaches the same tools; the 2026 revision is stateless, so such a client sends no `initialize` handshake. Legacy 2025 clients are unaffected.

### Fixed

- The MCP server is now built per request (HTTP) and per connection (stdio) rather than shared. The SDK caches the negotiated protocol revision on the server instance, so a shared one pinned itself to whichever revision arrived first and answered the other with `-32601 Method not found`.

## [2.1.1] - 2026-09-06

### Changed

- Depends on `@googleapis/sheets` and `@googleapis/drive` instead of the `googleapis` meta-package. Same generated client and the same `*_v*` types, from the same source; `googleapis` ships every Google API, and this package uses one or two of them. The installed SDK drops from 206 MB to 4 MB.

## [2.1.0] - 2026-09-06

### Fixed

- Works with `@mcp-z/oauth-google` 2.0.1, which replaced `toAuth()` with a token provider. Version 2.0.0 of this package resolves that release through its `^2.0.0` range and fails at runtime on any Google API call. Upgrade.

## [2.0.0] - 2026-09-06

### Changed

- Migrated to the v2 MCP SDK. `McpError`/`ErrorCode` are `ProtocolError`/`ProtocolErrorCode`, reached through `@mcp-z/server`; wire codes are unchanged.
- The 1.x line is maintained on `support/1.x` and published under the `support-1` dist-tag.

## [1.2.3] - 2026-09-05

### Fixed

- Origin validation and loopback bind for the HTTP transport (DNS rebinding).

## [1.2.0] - 2026-08-29

### Changed

- Exports smoke tests added.

## [1.1.0] - 2026-08-09

### Added

- `values-markdown-update`: nested hyperlinks, bold and italics.

## [1.0.0] - 2025-12-29

Initial release.
