import '../../lib/env-loader.ts';
import assert from 'assert';
import {
  // Cell count calculations
  calculateFiniteRangeCellCount,
  calculateRangeDimensions,
  columnIndexToString,
  columnStringToIndex,
  // Constants and types
  GOOGLE_SHEETS_LIMITS,
  // A1 notation validation
  isValidA1Notation,
  // Range parsing functions
  parseA1Notation,
  parseCellReference,
  type RangeReference,
  rangeToA1Notation,
  validateA1Notation,
} from '../../../src/spreadsheet/range-operations.ts';

describe('calculateFiniteRangeCellCount', () => {
  it('counts a single cell and finite rectangles', () => {
    assert.equal(calculateFiniteRangeCellCount('D9'), 1);
    assert.equal(calculateFiniteRangeCellCount('B2:D5'), 12);
    assert.equal(calculateFiniteRangeCellCount('A1:J25000'), 250000);
  });

  it('rejects unbounded, reversed and invalid ranges', () => {
    for (const range of ['A:A', '1:2', 'A5:B2', 'C1:A5', 'C5:A2', '', 'A0', 'Sheet1!A1']) {
      assert.throws(() => calculateFiniteRangeCellCount(range), Error, range);
    }
  });
});

describe('A1 Notation Validation', () => {
  describe('isValidA1Notation', () => {
    it('should validate single cell references', () => {
      assert.strictEqual(isValidA1Notation('A1'), true);
      assert.strictEqual(isValidA1Notation('Z99'), true);
      assert.strictEqual(isValidA1Notation('AA1'), true);
      assert.strictEqual(isValidA1Notation('AAA1'), true);
      assert.strictEqual(isValidA1Notation('B5'), true);
    });

    it('should validate cell ranges', () => {
      assert.strictEqual(isValidA1Notation('A1:B2'), true);
      assert.strictEqual(isValidA1Notation('A1:Z99'), true);
      assert.strictEqual(isValidA1Notation('AA1:BB10'), true);
      assert.strictEqual(isValidA1Notation('A1:AAA1000'), true);
    });

    it('should validate column ranges', () => {
      assert.strictEqual(isValidA1Notation('A:A'), true);
      assert.strictEqual(isValidA1Notation('A:Z'), true);
      assert.strictEqual(isValidA1Notation('AA:BB'), true);
      assert.strictEqual(isValidA1Notation('A:AAA'), true);
    });

    it('should validate row ranges', () => {
      assert.strictEqual(isValidA1Notation('1:1'), true);
      assert.strictEqual(isValidA1Notation('1:100'), true);
      assert.strictEqual(isValidA1Notation('5:10'), true);
      assert.strictEqual(isValidA1Notation('1:10000000'), true);
    });

    it('should reject invalid formats', () => {
      assert.strictEqual(isValidA1Notation(''), false);
      assert.strictEqual(isValidA1Notation('A'), false);
      assert.strictEqual(isValidA1Notation('1'), false);
      assert.strictEqual(isValidA1Notation('A0'), false);
      assert.strictEqual(isValidA1Notation('0A'), false);
      assert.strictEqual(isValidA1Notation('A1:'), false);
      assert.strictEqual(isValidA1Notation(':A1'), false);
      assert.strictEqual(isValidA1Notation('A1:B'), false);
      assert.strictEqual(isValidA1Notation('INVALID'), false);
      assert.strictEqual(isValidA1Notation('A1:B2:C3'), false);
    });

    it('should reject values exceeding Google Sheets limits', () => {
      // Test row limits (max 10,000,000)
      assert.strictEqual(isValidA1Notation('A10000001'), false);
      assert.strictEqual(isValidA1Notation('A1:A10000001'), false);
      assert.strictEqual(isValidA1Notation('10000001:10000001'), false);

      // Test column limits (max ZZZ which is 18278)
      assert.strictEqual(isValidA1Notation('AAAA1'), false);
    });

    it('should handle edge cases and boundary values', () => {
      // Maximum valid values
      assert.strictEqual(isValidA1Notation('ZZZ10000000'), true);
      assert.strictEqual(isValidA1Notation('A10000000'), true);
      assert.strictEqual(isValidA1Notation('10000000:10000000'), true);

      // Non-string inputs
      assert.strictEqual(isValidA1Notation(null as unknown as string), false);
      assert.strictEqual(isValidA1Notation(undefined as unknown as string), false);
      assert.strictEqual(isValidA1Notation(123 as unknown as string), false);
      assert.strictEqual(isValidA1Notation({} as unknown as string), false);
    });
  });

  describe('validateA1Notation', () => {
    it('should not throw for valid notation', () => {
      assert.doesNotThrow(() => validateA1Notation('A1'));
      assert.doesNotThrow(() => validateA1Notation('A1:B2'));
      assert.doesNotThrow(() => validateA1Notation('A:B'));
      assert.doesNotThrow(() => validateA1Notation('1:2'));
    });

    it('should throw descriptive errors for invalid notation', () => {
      assert.throws(() => validateA1Notation('INVALID'), /Invalid A1 notation: "INVALID". Valid formats: A1, A1:B2, A:B, 1:2/);
      assert.throws(() => validateA1Notation(''), /Invalid A1 notation: "". Valid formats: A1, A1:B2, A:B, 1:2/);
    });
  });
});

describe('Column Utilities', () => {
  describe('columnStringToIndex', () => {
    it('should convert single letter columns correctly', () => {
      assert.strictEqual(columnStringToIndex('A'), 1);
      assert.strictEqual(columnStringToIndex('B'), 2);
      assert.strictEqual(columnStringToIndex('Z'), 26);
    });

    it('should convert two letter columns correctly', () => {
      assert.strictEqual(columnStringToIndex('AA'), 27);
      assert.strictEqual(columnStringToIndex('AB'), 28);
      assert.strictEqual(columnStringToIndex('AZ'), 52);
      assert.strictEqual(columnStringToIndex('BA'), 53);
    });

    it('should convert three letter columns correctly', () => {
      assert.strictEqual(columnStringToIndex('AAA'), 703);
      assert.strictEqual(columnStringToIndex('ZZZ'), 18278);
    });
  });

  describe('columnIndexToString', () => {
    it('should convert single digit indices correctly', () => {
      assert.strictEqual(columnIndexToString(1), 'A');
      assert.strictEqual(columnIndexToString(2), 'B');
      assert.strictEqual(columnIndexToString(26), 'Z');
    });

    it('should convert two digit indices correctly', () => {
      assert.strictEqual(columnIndexToString(27), 'AA');
      assert.strictEqual(columnIndexToString(28), 'AB');
      assert.strictEqual(columnIndexToString(52), 'AZ');
      assert.strictEqual(columnIndexToString(53), 'BA');
    });

    it('should convert three digit indices correctly', () => {
      assert.strictEqual(columnIndexToString(703), 'AAA');
      assert.strictEqual(columnIndexToString(18278), 'ZZZ');
    });
  });

  it('should be reversible (string->index->string)', () => {
    const testCols = ['A', 'Z', 'AA', 'AZ', 'BA', 'ZZ', 'AAA', 'ZZZ'];
    for (const col of testCols) {
      const index = columnStringToIndex(col);
      const backToString = columnIndexToString(index);
      assert.strictEqual(backToString, col);
    }
  });

  it('should be reversible (index->string->index)', () => {
    const testIndices = [1, 26, 27, 52, 53, 702, 703, 18278];
    for (const index of testIndices) {
      const colString = columnIndexToString(index);
      const backToIndex = columnStringToIndex(colString);
      assert.strictEqual(backToIndex, index);
    }
  });
});

describe('Range Parsing Functions', () => {
  describe('parseCellReference', () => {
    it('should parse valid cell references', () => {
      const result1 = parseCellReference('A1');
      assert.strictEqual(result1.column, 'A');
      assert.strictEqual(result1.columnIndex, 1);
      assert.strictEqual(result1.row, 1);

      const result2 = parseCellReference('Z99');
      assert.strictEqual(result2.column, 'Z');
      assert.strictEqual(result2.columnIndex, 26);
      assert.strictEqual(result2.row, 99);

      const result3 = parseCellReference('AA100');
      assert.strictEqual(result3.column, 'AA');
      assert.strictEqual(result3.columnIndex, 27);
      assert.strictEqual(result3.row, 100);
    });

    it('should throw for invalid cell references', () => {
      assert.throws(() => parseCellReference('A'), /Invalid cell reference: A/);
      assert.throws(() => parseCellReference('1'), /Invalid cell reference: 1/);
      assert.throws(() => parseCellReference('A0'), /Invalid cell reference: A0/);
      assert.throws(() => parseCellReference('INVALID'), /Invalid cell reference: INVALID/);
    });
  });

  describe('parseA1Notation', () => {
    it('should parse single cell references', () => {
      const result = parseA1Notation('A1');
      assert.strictEqual(result.type, 'cell');
      assert.ok(result.startCell);
      assert.strictEqual(result.startCell.column, 'A');
      assert.strictEqual(result.startCell.row, 1);
      assert.strictEqual(result.endCell, result.startCell);
    });

    it('should parse cell ranges', () => {
      const result = parseA1Notation('A1:B2');
      assert.strictEqual(result.type, 'range');
      assert.ok(result.startCell);
      assert.ok(result.endCell);
      assert.strictEqual(result.startCell.column, 'A');
      assert.strictEqual(result.startCell.row, 1);
      assert.strictEqual(result.endCell.column, 'B');
      assert.strictEqual(result.endCell.row, 2);
    });

    it('should parse column ranges', () => {
      const result1 = parseA1Notation('A:A');
      assert.strictEqual(result1.type, 'column');
      assert.strictEqual(result1.startColumn, 'A');
      assert.strictEqual(result1.endColumn, 'A');

      const result2 = parseA1Notation('A:C');
      assert.strictEqual(result2.type, 'column');
      assert.strictEqual(result2.startColumn, 'A');
      assert.strictEqual(result2.endColumn, 'C');
    });

    it('should parse row ranges', () => {
      const result1 = parseA1Notation('1:1');
      assert.strictEqual(result1.type, 'row');
      assert.strictEqual(result1.startRow, 1);
      assert.strictEqual(result1.endRow, 1);

      const result2 = parseA1Notation('1:5');
      assert.strictEqual(result2.type, 'row');
      assert.strictEqual(result2.startRow, 1);
      assert.strictEqual(result2.endRow, 5);
    });

    it('should parse single column references', () => {
      const result = parseA1Notation('A:A');
      assert.strictEqual(result.type, 'column');
      assert.strictEqual(result.startColumn, 'A');
      assert.strictEqual(result.endColumn, 'A');
      assert.strictEqual(result.startColumnIndex, 1);
      assert.strictEqual(result.endColumnIndex, 1);
    });

    it('should parse single row references', () => {
      const result = parseA1Notation('5:5');
      assert.strictEqual(result.type, 'row');
      assert.strictEqual(result.startRow, 5);
      assert.strictEqual(result.endRow, 5);
    });

    it('should throw for invalid notation', () => {
      assert.throws(() => parseA1Notation('INVALID'), /Invalid A1 notation: "INVALID"/);
    });
  });

  describe('rangeToA1Notation', () => {
    it('should convert cell ranges back to notation', () => {
      const cellRange: RangeReference = {
        type: 'cell',
        startCell: { column: 'A', columnIndex: 1, row: 1 },
        endCell: { column: 'A', columnIndex: 1, row: 1 },
      };
      assert.strictEqual(rangeToA1Notation(cellRange), 'A1');
    });

    it('should convert cell ranges back to notation', () => {
      const range: RangeReference = {
        type: 'range',
        startCell: { column: 'A', columnIndex: 1, row: 1 },
        endCell: { column: 'B', columnIndex: 2, row: 2 },
      };
      assert.strictEqual(rangeToA1Notation(range), 'A1:B2');
    });

    it('should convert single column ranges back to notation', () => {
      const columnRange: RangeReference = {
        type: 'column',
        startColumn: 'A',
        endColumn: 'A',
        startColumnIndex: 1,
        endColumnIndex: 1,
      };
      assert.strictEqual(rangeToA1Notation(columnRange), 'A');
    });

    it('should convert multi-column ranges back to notation', () => {
      const columnRange: RangeReference = {
        type: 'column',
        startColumn: 'A',
        endColumn: 'C',
        startColumnIndex: 1,
        endColumnIndex: 3,
      };
      assert.strictEqual(rangeToA1Notation(columnRange), 'A:C');
    });

    it('should convert single row ranges back to notation', () => {
      const rowRange: RangeReference = {
        type: 'row',
        startRow: 5,
        endRow: 5,
      };
      assert.strictEqual(rangeToA1Notation(rowRange), '5');
    });

    it('should convert multi-row ranges back to notation', () => {
      const rowRange: RangeReference = {
        type: 'row',
        startRow: 1,
        endRow: 5,
      };
      assert.strictEqual(rangeToA1Notation(rowRange), '1:5');
    });

    it('should throw for invalid ranges', () => {
      assert.throws(() => rangeToA1Notation({ type: 'cell' } as RangeReference), /Invalid cell range: missing startCell/);
      assert.throws(() => rangeToA1Notation({ type: 'range' } as RangeReference), /Invalid range: missing start or end cell/);
      assert.throws(() => rangeToA1Notation({ type: 'column' } as RangeReference), /Invalid column range: missing start or end column/);
      assert.throws(() => rangeToA1Notation({ type: 'row' } as RangeReference), /Invalid row range: missing start or end row/);
      assert.throws(() => rangeToA1Notation({ type: 'unknown' } as unknown as RangeReference), /Unknown range type: unknown/);
    });
  });

  it('should roundtrip parse and convert correctly', () => {
    const testNotations = ['A1', 'A1:B2', 'A:C', '1:5'];
    for (const notation of testNotations) {
      const parsed = parseA1Notation(notation);
      const converted = rangeToA1Notation(parsed);
      assert.strictEqual(converted, notation);
    }
  });
});

describe('Cell Count and Dimension Calculations', () => {
  describe('calculateRangeDimensions', () => {
    it('should calculate dimensions for single cells', () => {
      const result = calculateRangeDimensions('A1');
      assert.strictEqual(result.rows, 1);
      assert.strictEqual(result.columns, 1);
      assert.strictEqual(result.cells, 1);
    });

    it('should calculate dimensions for cell ranges', () => {
      const result1 = calculateRangeDimensions('A1:B2');
      assert.strictEqual(result1.rows, 2);
      assert.strictEqual(result1.columns, 2);
      assert.strictEqual(result1.cells, 4);

      const result2 = calculateRangeDimensions('A1:C3');
      assert.strictEqual(result2.rows, 3);
      assert.strictEqual(result2.columns, 3);
      assert.strictEqual(result2.cells, 9);

      const result3 = calculateRangeDimensions('A1:E10');
      assert.strictEqual(result3.rows, 10);
      assert.strictEqual(result3.columns, 5);
      assert.strictEqual(result3.cells, 50);
    });

    it('should calculate dimensions for column ranges', () => {
      const result1 = calculateRangeDimensions('A:A');
      assert.strictEqual(result1.rows, GOOGLE_SHEETS_LIMITS.MAX_ROWS);
      assert.strictEqual(result1.columns, 1);
      assert.strictEqual(result1.cells, GOOGLE_SHEETS_LIMITS.MAX_ROWS);

      const result2 = calculateRangeDimensions('A:C');
      assert.strictEqual(result2.rows, GOOGLE_SHEETS_LIMITS.MAX_ROWS);
      assert.strictEqual(result2.columns, 3);
      assert.strictEqual(result2.cells, GOOGLE_SHEETS_LIMITS.MAX_ROWS * 3);
    });

    it('should calculate dimensions for row ranges', () => {
      const result1 = calculateRangeDimensions('1:1');
      assert.strictEqual(result1.rows, 1);
      assert.strictEqual(result1.columns, GOOGLE_SHEETS_LIMITS.MAX_COLUMNS);
      assert.strictEqual(result1.cells, GOOGLE_SHEETS_LIMITS.MAX_COLUMNS);

      const result2 = calculateRangeDimensions('1:5');
      assert.strictEqual(result2.rows, 5);
      assert.strictEqual(result2.columns, GOOGLE_SHEETS_LIMITS.MAX_COLUMNS);
      assert.strictEqual(result2.cells, 5 * GOOGLE_SHEETS_LIMITS.MAX_COLUMNS);
    });

    it('should handle large ranges correctly', () => {
      const result = calculateRangeDimensions('A1:Z1000');
      assert.strictEqual(result.rows, 1000);
      assert.strictEqual(result.columns, 26);
      assert.strictEqual(result.cells, 26000);
    });
  });
});

describe('Edge Cases and Boundary Conditions', () => {
  it('should handle maximum valid ranges', () => {
    // Test maximum row
    assert.doesNotThrow(() => parseA1Notation('A10000000'));
    assert.doesNotThrow(() => calculateRangeDimensions('A10000000'));

    // Test maximum column
    assert.doesNotThrow(() => parseA1Notation('ZZZ1'));
    assert.doesNotThrow(() => calculateRangeDimensions('ZZZ1'));

    // Test maximum range
    assert.doesNotThrow(() => parseA1Notation('A1:ZZZ10000000'));
  });

  it('should handle empty and single-cell ranges', () => {
    const singleCell = calculateRangeDimensions('A1');
    assert.strictEqual(singleCell.rows, 1);
    assert.strictEqual(singleCell.columns, 1);
    assert.strictEqual(singleCell.cells, 1);
  });

  it('should handle large range calculations without overflow', () => {
    // Test column range that would have many cells
    const columnRange = calculateRangeDimensions('A:A');
    assert.strictEqual(columnRange.cells, GOOGLE_SHEETS_LIMITS.MAX_ROWS);

    // Test row range that would have many cells
    const rowRange = calculateRangeDimensions('1:1');
    assert.strictEqual(rowRange.cells, GOOGLE_SHEETS_LIMITS.MAX_COLUMNS);
  });

  it('should handle various input types gracefully', () => {
    // Test error handling for various invalid inputs
    assert.throws(() => parseCellReference(''));
    assert.throws(() => parseA1Notation(''));
    assert.throws(() => calculateRangeDimensions('INVALID'));
  });

  it('should maintain type safety', () => {
    // Verify that the function returns the correct types
    const cellRef = parseCellReference('A1');
    assert.strictEqual(typeof cellRef.column, 'string');
    assert.strictEqual(typeof cellRef.columnIndex, 'number');
    assert.strictEqual(typeof cellRef.row, 'number');

    const rangeRef = parseA1Notation('A1:B2');
    assert.strictEqual(typeof rangeRef.type, 'string');
    assert.ok(['cell', 'range', 'column', 'row'].includes(rangeRef.type));

    const dimensions = calculateRangeDimensions('A1:B2');
    assert.strictEqual(typeof dimensions.rows, 'number');
    assert.strictEqual(typeof dimensions.columns, 'number');
    assert.strictEqual(typeof dimensions.cells, 'number');
  });

  it('should validate Google Sheets constants', () => {
    assert.strictEqual(typeof GOOGLE_SHEETS_LIMITS.MAX_ROWS, 'number');
    assert.strictEqual(typeof GOOGLE_SHEETS_LIMITS.MAX_COLUMNS, 'number');
    assert.strictEqual(typeof GOOGLE_SHEETS_LIMITS.MAX_CELLS, 'number');
    assert.strictEqual(typeof GOOGLE_SHEETS_LIMITS.MAX_BATCH_REQUESTS, 'number');
    assert.strictEqual(typeof GOOGLE_SHEETS_LIMITS.MAX_DIMENSION_BATCH_REQUESTS, 'number');

    assert.strictEqual(GOOGLE_SHEETS_LIMITS.MAX_ROWS, 10_000_000);
    assert.strictEqual(GOOGLE_SHEETS_LIMITS.MAX_COLUMNS, 18_278);
    assert.strictEqual(GOOGLE_SHEETS_LIMITS.MAX_CELLS, 10_000_000);
    assert.strictEqual(GOOGLE_SHEETS_LIMITS.MAX_BATCH_REQUESTS, 1000);
    assert.strictEqual(GOOGLE_SHEETS_LIMITS.MAX_DIMENSION_BATCH_REQUESTS, 100);
  });
});

describe('Mathematical Correctness', () => {
  it('should calculate range dimensions correctly for complex cases', () => {
    // Test various mathematical calculations
    const testCases = [
      { range: 'A1:B2', expectedRows: 2, expectedCols: 2, expectedCells: 4 },
      { range: 'A1:E10', expectedRows: 10, expectedCols: 5, expectedCells: 50 },
      { range: 'B2:F6', expectedRows: 5, expectedCols: 5, expectedCells: 25 },
      { range: 'AA1:AB10', expectedRows: 10, expectedCols: 2, expectedCells: 20 },
    ];

    for (const testCase of testCases) {
      const result = calculateRangeDimensions(testCase.range);
      assert.strictEqual(result.rows, testCase.expectedRows, `Rows mismatch for ${testCase.range}`);
      assert.strictEqual(result.columns, testCase.expectedCols, `Columns mismatch for ${testCase.range}`);
      assert.strictEqual(result.cells, testCase.expectedCells, `Cells mismatch for ${testCase.range}`);
    }
  });

  it('should handle column arithmetic correctly', () => {
    // Verify that column string/index conversion is mathematically correct
    const testCases = [
      { col: 'A', index: 1 },
      { col: 'Z', index: 26 },
      { col: 'AA', index: 27 },
      { col: 'AB', index: 28 },
      { col: 'AZ', index: 52 },
      { col: 'BA', index: 53 },
      { col: 'ZZ', index: 702 },
      { col: 'AAA', index: 703 },
    ];

    for (const testCase of testCases) {
      assert.strictEqual(columnStringToIndex(testCase.col), testCase.index);
      assert.strictEqual(columnIndexToString(testCase.index), testCase.col);
    }
  });
});
