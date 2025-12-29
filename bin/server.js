#!/usr/bin/env node

// biome-ignore lint/security/noGlobalEval: dual esm and cjs
if (typeof require === 'undefined') eval("import('../dist/esm/index.js').then((cli) => cli.default(process.argv.slice(2), 'mcp-sheets')).catch((err) => { console.log(err); process.exit(-1); });");
else require('../dist/cjs/index.js')(process.argv.slice(2), 'mcp-sheets');
