# Changelog

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
