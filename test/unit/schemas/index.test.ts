import { schemas } from '@mcp-z/mcp-sheets';
import assert from 'assert';

/**
 * TDD Tests for schemas.SheetGidSchema coercion
 *
 * These tests verify that gid values are properly coerced to strings,
 * especially for edge cases like gid=0 which is the default sheet ID.
 *
 * Issue: When MCP clients pass gid: 0 (number) instead of gid: "0" (string),
 * the comparison `String(sheetId) === gid` fails because "0" !== 0.
 */
describe('schemas.SheetGidSchema coercion', () => {
  it('should accept string "0" as valid gid', () => {
    const result = schemas.SheetGidSchema.safeParse('0');
    assert.ok(result.success, 'Should accept string "0"');
    assert.equal(result.data, '0', 'Should preserve string "0"');
  });

  it('should accept string "12345" as valid gid', () => {
    const result = schemas.SheetGidSchema.safeParse('12345');
    assert.ok(result.success, 'Should accept string "12345"');
    assert.equal(result.data, '12345', 'Should preserve string "12345"');
  });

  it('should coerce number 0 to string "0"', () => {
    // This is the key test - when MCP passes gid: 0 as a number
    const result = schemas.SheetGidSchema.safeParse(0);
    assert.ok(result.success, 'Should accept number 0 and coerce to string');
    assert.equal(result.data, '0', 'Should coerce number 0 to string "0"');
  });

  it('should coerce number 12345 to string "12345"', () => {
    const result = schemas.SheetGidSchema.safeParse(12345);
    assert.ok(result.success, 'Should accept number 12345 and coerce to string');
    assert.equal(result.data, '12345', 'Should coerce number 12345 to string "12345"');
  });

  it('should reject empty string', () => {
    const result = schemas.SheetGidSchema.safeParse('');
    assert.ok(!result.success, 'Should reject empty string');
  });

  // Note: null and undefined get coerced to "null" and "undefined" strings by z.coerce.string()
  // In practice, these are rejected at the parent object level when the field is required.
  // The critical fix is ensuring numbers (especially 0) are coerced to strings.
});
