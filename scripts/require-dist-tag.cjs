// npm publish moves `latest` to the highest version published, so a support-line
// release made after a higher major exists must name its dist-tag. Refuses the
// publish rather than letting a bare one silently move `latest` to the old line.
'use strict';

const { execSync } = require('node:child_process');

let branch;
try {
  branch = execSync('git rev-parse --abbrev-ref HEAD', { stdio: ['ignore', 'pipe', 'ignore'] })
    .toString()
    .trim();
} catch {
  // No branch to read (publishing from an unpacked tarball, a detached build): the
  // support-branch rule cannot apply, and a publish here is not the case being guarded.
  process.exit(0);
}

const match = /^support\/(\d+)\.x$/.exec(branch);
if (!match) process.exit(0);

const expected = `support-${match[1]}`;
let argv = [];
try {
  argv = JSON.parse(process.env.npm_config_argv || '{}').original || [];
} catch {}
const named = argv.some((arg, i) => arg === '--tag' && argv[i + 1] === expected) || argv.includes(`--tag=${expected}`);

if (!named) {
  console.error(`Publishing from ${branch} requires --tag ${expected}.`);
  console.error("A bare publish moves the 'latest' dist-tag to this line, so every plain install serves it.");
  process.exit(1);
}
