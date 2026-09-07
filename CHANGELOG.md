# Changelog

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
