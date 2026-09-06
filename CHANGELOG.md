# Changelog

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
