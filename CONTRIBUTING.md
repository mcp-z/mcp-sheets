# Contributing to @mcp-z/mcp-sheets

MCP server for Google Sheets integration with OAuth authentication, spreadsheet management, batch operations, and advanced formatting

## Before Starting

A few conventions here differ from what you might expect:

- **Breaking changes over compatibility.** This project has no compatibility burden yet. Do not add back-compat layers, migration utilities, or wrappers for deprecated APIs - change the API cleanly and bump the major.
- **Keep it approachable.** This is a small community project, not an enterprise codebase. Prefer the simplest solution that fits in the existing files over new abstractions, frameworks, or shared infrastructure.
- **Tests run against real services, not mocks.** Suites call live provider APIs with real credentials, so you need your own test account configured (copy `.env.test.example` to `.env.test`). A test that fails on credentials is reported, not skipped or loosened.
- **Never write to stdout in server code.** MCP speaks JSON-RPC over stdio; a stray `console.log` corrupts the protocol stream. Use the injected logger, which writes to stderr.
- **Test scratch goes in the package's gitignored `.tmp/`**, never `os.tmpdir()`.

## Branches

`master` is the only maintained release line. The 1.x line is retired and receives no fixes. All changes target `master`.

## Pre-Commit Commands

Install ts-dev-stack globally if not already installed:

```bash
npm install -g ts-dev-stack
```

Run before committing to format, build, type-check, lint, check dependencies, and generate API documentation:

```bash
tsds validate
npm test
```

`tsds validate` does not run tests. Run `npm test` separately. Validation also runs automatically on `npm publish` via the `prepublishOnly` hook; a failure blocks the publish.

## Testing

```bash
npm run test:setup    # Generate OAuth tokens (interactive, run once)
npm test              # Run the test suite
npm run test:engines  # Run the suite across every supported Node version
```

Specs live in `test/unit/`, mirroring `src/`. Cross-service specs live in `test/integration/`. Both run under `npm test`.

Run test commands from the package root. The test setup optionally loads the package root's `.env.test`, else `~/.config/mcp-z/mcp-sheets/.env.test` (the first one found is used) through `portable-env`; file values override matching inherited values. CI or the shell can supply values without a file. Enabled live tests require their necessary values, not the file itself.

Live integration specs share one spreadsheet per test process. The first spec that calls `sharedSpreadsheet()` (`test/lib/spreadsheet-helpers.ts`) creates `ci-shared-<run id>`. The process that created it deletes it at the end of the run, and also on SIGINT or SIGTERM. Each spec adds its own tabs with `addTestTabs`, addresses them only by gid or by the range helper the tabs come with, and removes every tab it created (tabs a tool created are removed too), in its own `after`; `cells-metadata-get` removes its extra tab in a `finally`, and `sheet-delete` removes only the tabs its tests did not already delete. Never write to or rename the shared default tab. A spec gets its own spreadsheet only when the spreadsheet itself is the subject: creating, renaming, the default tab's gid, or a second spreadsheet as a destination. The three metadata specs share one `sharedMetadataFixture()` instead. The run ends by deleting any spreadsheet no delete was attempted on and failing, naming the id and title of each one left behind, if any remains; a delete that already failed is never tried again. A killed run can leave `ci-shared-*` or `ci-metadata-*` spreadsheets behind; find them in Drive by that prefix and delete them by hand.

## Package Development

See `README.md` for package overview and usage.

## GitHub Actions

Trusted master pushes and manual CI runs call this commit's existing **Tests** workflow. Ubuntu runs first; Windows follows only after Ubuntu succeeds. Each platform installs and validates the package, seeds real credentials, runs full `npm test -- --bail`, then full `npm run test:engines -- --bail`. Both commands discover unit, integration and exports tests without runner-level environment preloads or file exclusions. The existing `test:ci` and `test:ci:engines` aliases also retain full discovery. Environment-aware test and setup entrypoints import the existing optional loader before dependent imports.

Pull requests receive no provider secrets. They run `npm test -- "test/exports/*.test.*" "test/unit/**/*.test.ts" "test/integration/server.test.ts" "test/integration/sheets-server-spawn.test.ts" "test/integration/csv-stream-lifecycle.test.ts"` and the same positive selection through `npm run test:engines`. This includes the six local CSV handler unit cases, ten stdio integration cases using the real built server with isolated fixture state, and one local HTTP CSV stream lifecycle case. The live metadata invocation lives in `test/integration/server-metadata.test.ts` and remains in full trusted discovery. Its child shares the existing token-owner store to preserve refresh rotations; its logs and other stores remain private. A green PR check does not certify live-provider behavior.

### Trusted tests

The existing workflow remains manually dispatchable from master for either platform. Automatic and manual credential jobs share one non-cancelling queued `live-provider-tests` job lock through seed, tests, persistence and cleanup. The protected `live-test` environment supplies current secrets when each job starts. Keep this repository's existing token implementation and environment ownership.

Configure these existing environment secrets:

- `CI_SECRETS_TOKEN`
- `TEST_ACCOUNT_ID`
- `TEST_REFRESH_TOKEN`
- `TEST_SCOPE`
- `GOOGLE_CLIENT_ID`
- `GOOGLE_CLIENT_SECRET`, optional for a public OAuth client

`CI_SECRETS_TOKEN` is a fine-grained GitHub PAT with this repository selected and **Environments: Read and write**. GitHub's default workflow token cannot update environment secrets; see the [environment-secret API permissions](https://docs.github.com/en/rest/actions/secrets#create-or-update-an-environment-secret). Keep its expiry visible to the maintainer. Provider refresh tokens and GitHub authorization have separate lifetimes.

After `npm run test:setup` signs in the test account, `npm run test:ci-setup` sets every secret above except `CI_SECRETS_TOKEN` in the `live-test` environment from local state, so a new test account or an expired refresh token needs no hand-set secrets. It reads the refresh tokens and `TEST_ACCOUNT_ID` from `.tokens/`, the scopes from `GOOGLE_SCOPE` in `src/constants.ts`, and the client values from `.env.test`. Refresh tokens are set last, and nothing is set while a required secret is missing. It runs `gh`, so sign in with `gh auth login` first; values go to `gh secret set` on stdin and are never printed.

`npm run test:ci-setup -- --dry` lists each secret as present or missing without setting anything, and reports whether `CI_SECRETS_TOKEN` exists. That token is never uploaded.

The existing seeder refuses pre-existing private files, validates required configuration, refreshes the real account, writes `.env.test` and `.tokens/store.json`, and persists the replacement refresh token before tests. Always-run persistence saves later rotations on test failure; cleanup removes only this run's owned private files. Missing configuration, failed renewal and failed persistence fail the job. Private credentials are never cached or uploaded. Forced cancellation or runner loss can prevent finalization; inspect both finalizers after execution. Reauthorize locally with `npm run test:setup` when consent or credential repair is needed. No Sheets test requires interactive browser consent; CI retains `TEST_INCLUDE_MANUAL=false` and `HEADLESS=true`.

Google external apps in Testing can issue seven-day refresh tokens for these scopes; successful access-token renewal does not remove that policy. Confirm the existing project's consent publishing configuration before relying on unattended runs. See [Google's token lifecycle documentation](https://developers.google.com/identity/protocols/oauth2).

Live tests share the existing process-local 1.2-second throttle across paced auth clients and raw-request reservations. Reported authorization counts are `getAccessToken` invocations, not OAuth refresh exchanges or exact HTTP request counts. The throttle does not coordinate child processes or repositories. Sheets normally permits 60 reads and 60 writes per user per project per minute; see [usage limits](https://developers.google.com/workspace/sheets/api/limits). Coordinate shared accounts, refresh grants and quotas across repositories and local runs. Provider failures remain failures. `--bail` stops the affected runtime; canonical nvu still attempts later engines after an engine failure, so this does not claim whole-sweep fail-fast behavior.

The 90-minute platform budget accommodates the retained historical estimate of about 14 minutes per full suite, roughly 70 minutes for current Node plus four admitted engines, before install/validation/credential overhead. This is a planning estimate, not a fresh performance measurement or a provider-run authorization.
