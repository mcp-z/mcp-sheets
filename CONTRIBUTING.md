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

Two lines. `master` is the current major and where all new work goes; `support/1.x` maintains the 1.x line for consumers who have not migrated.

    master            2.x    current    the v2 MCP SDK, both protocol eras
    support/1.x       1.x    security fixes only, cut at v1.2.3

Check which one you are on before editing:

```bash
git rev-parse --abbrev-ref HEAD
```

Features, dependency migrations and API changes go to `master` only. A security fix that also affects 1.x is **cherry-picked** to `support/1.x` — never merge the branches into each other, in either direction.

This file is the 1.x line's guide too. It lives only on `master` so it cannot drift between the lines; from `support/1.x`, read it with `git show master:CONTRIBUTING.md`.

Releasing `support/1.x` carries one trap. `npm publish` moves `latest` to the highest version published, so a 1.x release made after the next major exists must name its dist-tag or every bare `npm install @mcp-z/mcp-sheets` serves the old line:

```bash
npm publish --tag support-1
```

`prepublishOnly` refuses a bare publish from `support/*`, so forgetting the flag fails the publish rather than moving `latest`. `npm dist-tag add @mcp-z/mcp-sheets@<version> latest` reverses a mistake at any time.

## Pre-Commit Commands

Install ts-dev-stack globally if not already installed:

```bash
npm install -g ts-dev-stack
```

Run before committing - this builds, type-checks, lints, and tests:

```bash
tsds validate
```

`tsds validate` also runs automatically on `npm publish` via the `prepublishOnly` hook; a failure blocks the publish.

## Testing

```bash
npm run test:setup    # Generate OAuth tokens (interactive, run once)
npm test              # Run the test suite
npm run test:engines  # Run the suite across every supported Node version
```

Specs live in `test/unit/`, mirroring `src/`. Cross-service specs live in `test/integration/`. Both run under `npm test`.

## Package Development

See `README.md` for package overview and usage.

## GitHub Actions

CI follows the Linux/Windows template used by each-package: Node 26, `npm ci`, `prepublishOnly`, a current-runtime test run, and the supported-engine sweep. macOS coverage runs locally. Pull requests receive no provider credentials.

`npm run test:ci` and `npm run test:ci:engines` run the credential-free selection. They exclude `test/integration/**`, `test/unit/mcp/tools/csv-get-columns.test.ts`. These files require provider configuration, live services, or interactive consent; some also contain local checks. Their exclusion is a coverage gap until the separate live-service automation is provisioned.

`npm test` and `npm run test:engines` retain full discovery. CI sets `TEST_INCLUDE_MANUAL=false`; consent tests require a person and run locally with `TEST_INCLUDE_MANUAL=true`. A green credential-free check does not certify live-provider behavior. Release evidence must include the configured live suites and relevant manual OAuth flows.

### Live provider tests

Run the **Live provider tests** workflow from master, selecting Linux or Windows. It runs the full non-interactive suite once with `--bail`, without an engine matrix or automatic retries. Live runs are manual while request usage and service quotas are being measured. A green PR check covers only the credential-free selection above.

This repository owns its `live-test` GitHub environment, restricted to master, and its own live-job concurrency group. There is no cross-repository coordinator. Concurrent runs in different repositories can still share provider quotas; avoid starting several against the same account at once. Browser consent tests remain local and opt-in.

Configure these individual environment secrets from the existing test configuration and token stores:

- `CI_SECRETS_TOKEN`
- `TEST_ACCOUNT_ID`
- `TEST_REFRESH_TOKEN`
- `TEST_SCOPE`
- `GOOGLE_CLIENT_ID`
- `GOOGLE_CLIENT_SECRET`

`CI_SECRETS_TOKEN` is a fine-grained GitHub PAT with this repository selected and **Environments: Read and write**. GitHub's default workflow token cannot update environment secrets; see the [environment-secret API permissions](https://docs.github.com/en/rest/actions/secrets#create-or-update-an-environment-secret). Keep the PAT's expiry visible to the maintainer; do not use the local broad GitHub login as a CI secret. Provider refresh tokens and GitHub authorization have separate lifetimes.

The seeder validates configuration, renews provider credentials, saves private runner files, and writes replacement refresh tokens back to environment secrets before tests run. An always-run finalizer saves subsequent replacements even when tests fail, then removes runner credential files. Credentials are not cached or uploaded as artifacts. Missing configuration, failed renewal and failed persistence fail the job; CI never opens a consent screen. After provider revocation or expiry, reauthorize locally with the existing setup command and reseed that repository's refresh-token secrets.

Google external apps in Testing can issue seven-day refresh tokens for these scopes; successful access-token renewal does not remove that policy. Confirm the existing project's consent publishing configuration before relying on long-term unattended runs. See [Google's token lifecycle documentation](https://developers.google.com/identity/protocols/oauth2).

Live tests share a process-local authorization throttle across their real OAuth clients. It spaces authorization starts by 1.2 seconds, below Sheets' 60 reads or writes per minute per user. Both direct fixture clients and tool middleware use it; the production clients keep their normal behavior. The test runner reports authorization counts, not exact HTTP request counts. Provider failures still fail the run, and the throttle does not coordinate separate repositories or processes.
