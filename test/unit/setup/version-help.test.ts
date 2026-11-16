import assert from 'assert';
import { handleVersionHelp } from '../../../src/setup/version-help.ts';

describe('setup/version-help', () => {
  it('help names the private configuration file location', () => {
    const { handled, output } = handleVersionHelp(['--help']);
    assert.strictEqual(handled, true);
    assert.ok(output?.includes('~/.config/mcp-z/mcp-sheets/.env'));
  });
});
