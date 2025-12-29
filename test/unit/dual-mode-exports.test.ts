import assert from 'assert';
import { createRequire } from 'module';
import * as path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const require = createRequire(import.meta.url);

describe('Sheets Server Dual-Mode Exports', () => {
  it('should load successfully in CommonJS mode', () => {
    // Tests dual ESM/CJS exports for cross-platform compatibility
    const serverPath = path.resolve(__dirname, '../../dist/cjs/index.js');

    // This will throw if there are any import/export issues
    const server = require(serverPath);

    // Verify expected exports
    assert.ok(server, 'sheets server module should export something');
    assert.strictEqual(typeof server.setup.createStdioServer, 'function', 'Should export createStdioServer function');
    assert.strictEqual(typeof server.setup.createHTTPServer, 'function', 'Should export createHTTPServer function');
    assert.strictEqual(typeof server.default, 'function', 'Should export default main function for bin script');

    // Each server may have different scope exports
    const expectedScope = 'GOOGLE_SCOPE' as string;
    if (expectedScope !== 'NONE') {
      assert.ok(server[expectedScope as keyof typeof server], `Should export ${expectedScope}`);
    }
  });

  it('should have working dual-mode __filename pattern', () => {
    // Test your dual-mode pattern works
    const filename = typeof __filename !== 'undefined' ? __filename : fileURLToPath(import.meta.url);
    const dirname = path.dirname(filename);

    assert.ok(filename.endsWith('.test.ts'), 'Should resolve correct filename');
    assert.ok(dirname.includes('test/unit'), 'Should resolve correct dirname');
  });

  it('should work with Node.js built-in namespace imports', () => {
    // Test that Node.js built-ins work as expected
    const fs = require('fs');
    const path = require('path');
    const os = require('os');
    const url = require('url');

    assert.strictEqual(typeof fs.readFileSync, 'function');
    assert.strictEqual(typeof path.join, 'function');
    assert.strictEqual(typeof os.homedir, 'function');
    assert.strictEqual(typeof url.fileURLToPath, 'function');
  });
});
