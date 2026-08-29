const assert = require('assert');
const { default: main, setup } = require('@mcp-z/mcp-sheets');

describe('exports .cjs', () => {
  it('named exports resolve', () => {
    assert.equal(typeof main, 'function');
    for (const fn of [setup.createStdioServer, setup.createHTTPServer]) assert.equal(typeof fn, 'function');
  });
});
