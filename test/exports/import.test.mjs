import main, { setup } from '@mcp-z/mcp-sheets';
import assert from 'assert';

describe('exports .mjs', () => {
  it('named exports resolve', () => {
    assert.equal(typeof main, 'function');
    for (const fn of [setup.createStdioServer, setup.createHTTPServer]) assert.equal(typeof fn, 'function');
  });
});
