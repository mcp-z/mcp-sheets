import assert from 'assert';
import * as schemas from '../../../src/schemas/index.ts';

describe('schemas.StrictSheetGidSchema', () => {
  it('accepts non-negative integers and digit strings, normalising to a digit string', () => {
    for (const [input, expected] of [
      [0, '0'],
      ['0', '0'],
      ['123', '123'],
      [456, '456'],
    ] as const) {
      const result = schemas.StrictSheetGidSchema.safeParse(input);
      assert.ok(result.success, `Should accept ${String(input)}`);
      assert.strictEqual(result.data, expected);
    }
  });

  it('rejects null, undefined, booleans, negatives, decimals and non-digit strings', () => {
    for (const input of [null, undefined, true, false, -1, 1.5, 'abc', '', '-1', '1.5', ' 1', '1 ', NaN, Infinity, {}, []]) {
      assert.strictEqual(schemas.StrictSheetGidSchema.safeParse(input).success, false, `Should reject ${String(input)}`);
    }
  });

  it('caps gids at int32 for numbers and digit strings', () => {
    assert.strictEqual(schemas.StrictSheetGidSchema.parse(2147483647), '2147483647');
    assert.strictEqual(schemas.StrictSheetGidSchema.parse('2147483647'), '2147483647');
    assert.strictEqual(schemas.StrictSheetGidSchema.parse('007'), '7');
    for (const input of [2147483648, '2147483648', '99999999999999999999', Number.MAX_SAFE_INTEGER]) {
      assert.strictEqual(schemas.StrictSheetGidSchema.safeParse(input).success, false, `Should reject ${String(input)}`);
    }
  });
});
