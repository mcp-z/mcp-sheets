# Contributing to MCP-Z

## Before Starting

**MUST READ**:
- [QUALITY.MD](QUALITY.md) - Quality principles (summarize before starting work)

## Pre-Commit Commands

Install ts-dev-stack globally if not already installed:
```bash
npm install -g ts-dev-stack
```

Run before committing:
```bash
tsds validate
```

## Pre-Publish Verification

Before publishing packages to npm, verify each package builds and runs correctly:

```bash
# Verify specific package
cd servers/mcp-gmail && npm run verify

# Verify all server packages
for server in gmail outlook sheets drive workflows; do
  echo "Verifying @mcp-z/mcp-$server..."
  cd servers/$server && npm run verify && cd ../.. || exit 1
done
```

**Two-Phase Verification:**

**Phase 1: Built Package Verification** (fast)
- ✅ Package builds successfully (TypeScript compilation)
- ✅ Required files exist (bin/, dist/)
- ✅ Server starts using @mcp-z/cli infrastructure
- ✅ Tools respond correctly

**Phase 2: npm pack + Install Verification** (comprehensive)
- ✅ npm pack creates valid tarball
- ✅ Installation from tarball succeeds (production dependencies only)
- ✅ Installed package structure valid
- ✅ Excluded files not present (scripts/, test/, src/)
- ✅ Server starts from installed package
- ✅ Runtime dependencies complete (no missing deps)

**Production-Grade Error Handling:**
Uses @mcp-z/cli infrastructure for:
- ✅ Timeout protection (5s graceful shutdown)
- ✅ SIGKILL fallback (prevents hanging)
- ✅ No orphaned processes
- ✅ Comprehensive cleanup

**Automatic verification:**
Verification runs automatically during `npm publish` via the `prepublishOnly` hook. If **either phase** fails, publish is blocked.

**Execution time:** ~30-60 seconds per package (Phase 1: ~15s, Phase 2: ~30s)

## Package Development

See package documentation:
- `README.md` - Package overview and usage
- `QUALITY.md` - Quality principles and standards
- `CLAUDE.md` - Development patterns and architecture guidance

