import assert from 'assert';
import { NATIVE_DATE_EXAMPLE_FORMAT, NATIVE_DATE_EXAMPLE_READ, NATIVE_DATE_EXAMPLE_WRITE, NATIVE_DATE_GUIDANCE } from '../../src/constants.ts';
import { toolFactories } from '../../src/mcp/index.ts';

const ids = { id: 'SID', gid: '0' };

describe('native date example', () => {
  it('parses with the real tool input schemas', () => {
    const write = toolFactories.valuesBatchUpdate().config.inputSchema.parse({ ...ids, ...NATIVE_DATE_EXAMPLE_WRITE, includeData: false });
    assert.strictEqual(write.valueInputOption, 'RAW');
    toolFactories.cellsFormat().config.inputSchema.parse({ ...ids, ...NATIVE_DATE_EXAMPLE_FORMAT });
    for (const factory of [toolFactories.rowsGet, toolFactories.columnsGet, toolFactories.valuesSearch]) {
      const { shape } = factory().config.inputSchema;
      assert.ok(shape.render, 'reader exposes render');
      assert.strictEqual(shape.render.parse(NATIVE_DATE_EXAMPLE_READ.render), 'UNFORMATTED_VALUE');
    }
    toolFactories.rowsGet().config.inputSchema.parse({ ...ids, ...NATIVE_DATE_EXAMPLE_READ });
  });

  it('builds the guidance example from the constants', () => {
    assert.ok(NATIVE_DATE_GUIDANCE.includes(JSON.stringify(NATIVE_DATE_EXAMPLE_WRITE)));
    assert.ok(NATIVE_DATE_GUIDANCE.includes(JSON.stringify(NATIVE_DATE_EXAMPLE_FORMAT)));
    assert.ok(NATIVE_DATE_GUIDANCE.includes(JSON.stringify(NATIVE_DATE_EXAMPLE_READ)));
  });
});
