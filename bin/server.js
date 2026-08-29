#!/usr/bin/env node

// Checks --version/--help/`version` via the dependency-free version-help module before ever
// touching index.js, which statically re-exports the mcp/setup/schemas namespaces (googleapis,
// @modelcontextprotocol/sdk, @mcp-z/oauth-google, @mcp-z/server et al.) -- importing index.js at
// all, even without calling anything in it, evaluates that whole graph.
if (typeof require === 'undefined') {
  // biome-ignore lint/security/noGlobalEval: dual esm and cjs
  eval(
    "import('../dist/esm/setup/version-help.js').then(({ handleVersionHelp }) => { const result = handleVersionHelp(process.argv.slice(2)); if (result.handled) { console.log(result.output); process.exit(0); } return import('../dist/esm/index.js').then((cli) => cli.default(process.argv.slice(2), 'mcp-sheets')); }).catch((err) => { console.error(err); process.exit(-1); });"
  );
} else {
  const { handleVersionHelp } = require('../dist/cjs/setup/version-help.js');
  const result = handleVersionHelp(process.argv.slice(2));
  if (result.handled) {
    console.log(result.output);
    process.exit(0);
  }
  require('../dist/cjs/index.js')(process.argv.slice(2), 'mcp-sheets');
}
