import assert from 'assert';
import { handleVersionHelp } from '../../../src/setup/version-help.ts';

describe('setup/version-help', () => {
  it('help names the private configuration file location', () => {
    const { handled, output } = handleVersionHelp(['--help']);
    assert.strictEqual(handled, true);
    assert.ok(output?.includes('~/.config/mcp-z/mcp-sheets/.env'));
  });

  it('help names the full resource store default', () => {
    assert.ok(handleVersionHelp(['--help']).output?.includes('(default: .mcp-z/mcp-sheets/files beside the nearest .mcp.json, else ~/.mcp-z/mcp-sheets/files)'));
  });
});
