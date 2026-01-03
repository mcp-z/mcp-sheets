import assert from 'assert';
import { DEFAULT_APPEND_COUNT, DEFAULT_COLUMN_COUNT, DEFAULT_ROW_COUNT, type DimensionOperation, type DimensionRequest, type DimensionType, MAX_COLUMN_COUNT, MAX_ROW_COUNT, sortOperations } from '../../../../../src/mcp/tools/lib/dimension-operations.ts';

describe('dimension-operations', () => {
  describe('constants', () => {
    it('defines valid default constants', () => {
      assert.strictEqual(DEFAULT_APPEND_COUNT, 1);
      assert.strictEqual(DEFAULT_ROW_COUNT, 1000);
      assert.strictEqual(DEFAULT_COLUMN_COUNT, 26);
      assert.strictEqual(MAX_ROW_COUNT, 10000000);
      assert.strictEqual(MAX_COLUMN_COUNT, 18278);
    });

    it('ensures max values are greater than defaults', () => {
      assert.ok(MAX_ROW_COUNT > DEFAULT_ROW_COUNT);
      assert.ok(MAX_COLUMN_COUNT > DEFAULT_COLUMN_COUNT);
    });
  });

  describe('sortOperations', () => {
    it('sorts operations by priority: delete -> insert -> append', () => {
      const operations: DimensionRequest[] = [
        {
          operation: 'appendDimension',
          dimension: 'ROWS',
          startIndex: 5,
        },
        {
          operation: 'insertDimension',
          dimension: 'ROWS',
          startIndex: 2,
        },
        {
          operation: 'deleteDimension',
          dimension: 'ROWS',
          startIndex: 1,
          endIndex: 2,
        },
      ];

      const sorted = sortOperations(operations);

      assert.strictEqual(sorted[0]?.operation, 'deleteDimension');
      assert.strictEqual(sorted[1]?.operation, 'insertDimension');
      assert.strictEqual(sorted[2]?.operation, 'appendDimension');
    });

    it('maintains original array and returns new sorted array', () => {
      const operations: DimensionRequest[] = [
        {
          operation: 'appendDimension',
          dimension: 'ROWS',
          startIndex: 5,
        },
        {
          operation: 'deleteDimension',
          dimension: 'ROWS',
          startIndex: 1,
          endIndex: 2,
        },
      ];

      const sorted = sortOperations(operations);

      // Original array unchanged
      assert.strictEqual(operations[0]?.operation, 'appendDimension');
      assert.strictEqual(operations[1]?.operation, 'deleteDimension');

      // New array is sorted
      assert.strictEqual(sorted[0]?.operation, 'deleteDimension');
      assert.strictEqual(sorted[1]?.operation, 'appendDimension');
      assert.notStrictEqual(sorted, operations);
    });

    it('sorts delete operations in descending order by startIndex', () => {
      const operations: DimensionRequest[] = [
        {
          operation: 'deleteDimension',
          dimension: 'ROWS',
          startIndex: 1,
          endIndex: 2,
        },
        {
          operation: 'deleteDimension',
          dimension: 'ROWS',
          startIndex: 5,
          endIndex: 6,
        },
        {
          operation: 'deleteDimension',
          dimension: 'ROWS',
          startIndex: 3,
          endIndex: 4,
        },
      ];

      const sorted = sortOperations(operations);

      assert.strictEqual(sorted[0]?.startIndex, 5);
      assert.strictEqual(sorted[1]?.startIndex, 3);
      assert.strictEqual(sorted[2]?.startIndex, 1);
    });

    it('sorts insert operations in ascending order by startIndex', () => {
      const operations: DimensionRequest[] = [
        {
          operation: 'insertDimension',
          dimension: 'ROWS',
          startIndex: 5,
        },
        {
          operation: 'insertDimension',
          dimension: 'ROWS',
          startIndex: 1,
        },
        {
          operation: 'insertDimension',
          dimension: 'ROWS',
          startIndex: 3,
        },
      ];

      const sorted = sortOperations(operations);

      assert.strictEqual(sorted[0]?.startIndex, 1);
      assert.strictEqual(sorted[1]?.startIndex, 3);
      assert.strictEqual(sorted[2]?.startIndex, 5);
    });

    it('handles mixed dimension types correctly', () => {
      const operations: DimensionRequest[] = [
        {
          operation: 'deleteDimension',
          dimension: 'COLUMNS',
          startIndex: 2,
          endIndex: 3,
        },
        {
          operation: 'insertDimension',
          dimension: 'ROWS',
          startIndex: 1,
        },
        {
          operation: 'appendDimension',
          dimension: 'COLUMNS',
          startIndex: 5,
        },
      ];

      const sorted = sortOperations(operations);

      assert.strictEqual(sorted[0]?.operation, 'deleteDimension');
      assert.strictEqual(sorted[0]?.dimension, 'COLUMNS');
      assert.strictEqual(sorted[1]?.operation, 'insertDimension');
      assert.strictEqual(sorted[1]?.dimension, 'ROWS');
      assert.strictEqual(sorted[2]?.operation, 'appendDimension');
      assert.strictEqual(sorted[2]?.dimension, 'COLUMNS');
    });

    it('handles empty operations array', () => {
      const operations: DimensionRequest[] = [];
      const sorted = sortOperations(operations);

      assert.strictEqual(sorted.length, 0);
      assert.notStrictEqual(sorted, operations);
    });

    it('handles single operation', () => {
      const operations: DimensionRequest[] = [
        {
          operation: 'insertDimension',
          dimension: 'ROWS',
          startIndex: 1,
        },
      ];

      const sorted = sortOperations(operations);

      assert.strictEqual(sorted.length, 1);
      assert.strictEqual(sorted[0]?.operation, 'insertDimension');
      assert.notStrictEqual(sorted, operations);
    });

    it('preserves all properties of dimension requests', () => {
      const operations: DimensionRequest[] = [
        {
          operation: 'insertDimension',
          dimension: 'ROWS',
          startIndex: 1,
          endIndex: 2,
          inheritFromBefore: true,
        },
      ];

      const sorted = sortOperations(operations);

      assert.strictEqual(sorted[0]?.operation, 'insertDimension');
      assert.strictEqual(sorted[0]?.dimension, 'ROWS');
      assert.strictEqual(sorted[0]?.startIndex, 1);
      assert.strictEqual(sorted[0]?.endIndex, 2);
      assert.strictEqual(sorted[0]?.inheritFromBefore, true);
    });
  });

  describe('type definitions', () => {
    it('defines correct operation types', () => {
      const validOperations: DimensionOperation[] = ['insertDimension', 'deleteDimension', 'appendDimension'];

      // This test ensures TypeScript types are working correctly
      validOperations.forEach((op) => {
        assert.ok(typeof op === 'string');
      });
    });

    it('defines correct dimension types', () => {
      const validDimensions: DimensionType[] = ['ROWS', 'COLUMNS'];

      // This test ensures TypeScript types are working correctly
      validDimensions.forEach((dim) => {
        assert.ok(typeof dim === 'string');
      });
    });
  });
});
