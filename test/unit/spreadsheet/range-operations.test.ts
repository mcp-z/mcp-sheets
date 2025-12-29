import assert from 'assert';
import {
  // Batch operation builders
  buildValuesBatchUpdateRequest,
  // Cell count calculations
  calculateRangeDimensions,
  calculateTotalCells,
  columnIndexToString,
  columnStringToIndex,
  detectRangeConflicts,
  // Additional utilities
  expandRange,
  // Constants and types
  GOOGLE_SHEETS_LIMITS,
  getRangeIntersection,
  // A1 notation validation
  isValidA1Notation,
  // Range parsing functions
  parseA1Notation,
  parseCellReference,
  type RangeReference,
  // Range conflict detection
  rangesOverlap,
  rangeToA1Notation,
  splitRangeIntoChunks,
  validateA1Notation,
  validateBatchRanges,
} from '../../../src/spreadsheet/range-operations.js';

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

  describe('calculateTotalCells', () => {
    it('should calculate total cells for single range', () => {
      const total = calculateTotalCells(['A1:B2']);
      assert.strictEqual(total, 4);
    });

    it('should calculate total cells for multiple ranges', () => {
      const total = calculateTotalCells(['A1:B2', 'C1:C3', 'E1']);
      // A1:B2 = 4, C1:C3 = 3, E1 = 1 => total = 8
      assert.strictEqual(total, 8);
    });

    it('should handle empty array', () => {
      const total = calculateTotalCells([]);
      assert.strictEqual(total, 0);
    });

    it('should handle mix of range types', () => {
      const total = calculateTotalCells(['A1', 'B1:C2', 'D:D']);
      // A1 = 1, B1:C2 = 4, D:D = MAX_ROWS
      assert.strictEqual(total, 1 + 4 + GOOGLE_SHEETS_LIMITS.MAX_ROWS);
    });
  });
});

describe('Batch Operation Builders', () => {
  describe('buildValuesBatchUpdateRequest', () => {
    it('should build basic batch update request', () => {
      const requests = [
        {
          range: 'A1:B2',
          values: [
            ['a', 'b'],
            ['c', 'd'],
          ],
        },
      ];

      const result = buildValuesBatchUpdateRequest(requests, 'Sheet1');

      assert.strictEqual(result.valueInputOption, 'USER_ENTERED');
      assert.strictEqual(result.includeValuesInResponse, false);
      assert.strictEqual(result.responseDateTimeRenderOption, 'FORMATTED_STRING');
      assert.strictEqual(result.responseValueRenderOption, 'FORMATTED_VALUE');
      assert.strictEqual(result.data.length, 1);
      assert.ok(result.data[0], 'Expected data[0] to exist');
      assert.strictEqual(result.data[0].range, 'Sheet1!A1:B2');
      assert.strictEqual(result.data[0].majorDimension, 'ROWS');
      assert.deepStrictEqual(result.data[0].values, [
        ['a', 'b'],
        ['c', 'd'],
      ]);
    });

    it('should respect custom options', () => {
      const requests = [
        {
          range: 'A1:B2',
          values: [
            ['a', 'b'],
            ['c', 'd'],
          ],
          majorDimension: 'COLUMNS' as const,
        },
      ];

      const options = {
        valueInputOption: 'RAW' as const,
        includeValuesInResponse: true,
        responseDateTimeRenderOption: 'SERIAL_NUMBER' as const,
        responseValueRenderOption: 'UNFORMATTED_VALUE' as const,
      };

      const result = buildValuesBatchUpdateRequest(requests, 'MySheet', options);

      assert.strictEqual(result.valueInputOption, 'RAW');
      assert.strictEqual(result.includeValuesInResponse, true);
      assert.strictEqual(result.responseDateTimeRenderOption, 'SERIAL_NUMBER');
      assert.strictEqual(result.responseValueRenderOption, 'UNFORMATTED_VALUE');
      assert.ok(result.data[0], 'Expected data[0] to exist');
      assert.strictEqual(result.data[0].majorDimension, 'COLUMNS');
    });

    it('should handle multiple requests', () => {
      const requests = [
        {
          range: 'A1:B2',
          values: [
            ['a', 'b'],
            ['c', 'd'],
          ],
        },
        {
          range: 'D1:E2',
          values: [
            ['e', 'f'],
            ['g', 'h'],
          ],
        },
      ];

      const result = buildValuesBatchUpdateRequest(requests, 'TestSheet');

      assert.strictEqual(result.data.length, 2);
      assert.ok(result.data[0], 'Expected data[0] to exist');
      assert.ok(result.data[1], 'Expected data[1] to exist');
      assert.strictEqual(result.data[0].range, 'TestSheet!A1:B2');
      assert.strictEqual(result.data[1].range, 'TestSheet!D1:E2');
    });

    it('should validate ranges and throw for invalid ones', () => {
      const requests = [
        {
          range: 'INVALID_RANGE',
          values: [['a']],
        },
      ];

      assert.throws(() => buildValuesBatchUpdateRequest(requests, 'Sheet1'), /Invalid range in request 0:/);
    });

    it('should validate total cell limits', () => {
      // Create a request that exceeds the cell limit
      const hugeValues = Array(5000).fill(Array(5000).fill('data'));
      const requests = [
        {
          range: 'A1:NTP5000', // This would be 5000 * 5000 = 25M cells
          values: hugeValues,
        },
      ];

      assert.throws(() => buildValuesBatchUpdateRequest(requests, 'Sheet1'), /Batch update exceeds maximum cells limit/);
    });
  });
});

describe('Range Conflict Detection', () => {
  describe('rangesOverlap', () => {
    it('should detect overlapping cell ranges', () => {
      assert.strictEqual(rangesOverlap('A1:B2', 'B1:C2'), true);
      assert.strictEqual(rangesOverlap('A1:C3', 'B2:D4'), true);
      assert.strictEqual(rangesOverlap('A1:B2', 'A1:B2'), true); // identical ranges
    });

    it('should detect non-overlapping cell ranges', () => {
      assert.strictEqual(rangesOverlap('A1:B2', 'C3:D4'), false);
      assert.strictEqual(rangesOverlap('A1:A1', 'B1:B1'), false);
      assert.strictEqual(rangesOverlap('A1:B1', 'A2:B2'), false);
    });

    it('should handle single cells', () => {
      assert.strictEqual(rangesOverlap('A1', 'A1'), true);
      assert.strictEqual(rangesOverlap('A1', 'B1'), false);
      assert.strictEqual(rangesOverlap('A1', 'A1:B2'), true);
    });

    it('should handle column ranges', () => {
      assert.strictEqual(rangesOverlap('A:A', 'A:A'), true);
      assert.strictEqual(rangesOverlap('A:B', 'B:C'), true);
      assert.strictEqual(rangesOverlap('A:A', 'B:B'), false);
      assert.strictEqual(rangesOverlap('A:A', 'A1:A100'), true);
    });

    it('should handle row ranges', () => {
      assert.strictEqual(rangesOverlap('1:1', '1:1'), true);
      assert.strictEqual(rangesOverlap('1:2', '2:3'), true);
      assert.strictEqual(rangesOverlap('1:1', '2:2'), false);
      assert.strictEqual(rangesOverlap('1:1', 'A1:Z1'), true);
    });

    it('should handle mixed range types', () => {
      assert.strictEqual(rangesOverlap('A:A', '1:1'), true); // Column A intersects row 1 at A1
      assert.strictEqual(rangesOverlap('A1:B2', 'A:A'), true);
      assert.strictEqual(rangesOverlap('A1:B2', '1:1'), true);
    });

    it('should handle invalid ranges gracefully', () => {
      assert.strictEqual(rangesOverlap('INVALID', 'A1:B2'), false);
      assert.strictEqual(rangesOverlap('A1:B2', 'INVALID'), false);
      assert.strictEqual(rangesOverlap('INVALID', 'INVALID'), false);
    });
  });

  describe('detectRangeConflicts', () => {
    it('should detect no conflicts in non-overlapping ranges', () => {
      const ranges = ['A1:A2', 'B1:B2', 'C1:C2'];
      const conflicts = detectRangeConflicts(ranges);
      assert.strictEqual(conflicts.length, 0);
    });

    it('should detect conflicts in overlapping ranges', () => {
      const ranges = ['A1:B2', 'B1:C2', 'D1:D1'];
      const conflicts = detectRangeConflicts(ranges);
      assert.strictEqual(conflicts.length, 1);
      assert.ok(conflicts[0], 'Expected conflicts[0] to exist');
      assert.strictEqual(conflicts[0].range1, 'A1:B2');
      assert.strictEqual(conflicts[0].range2, 'B1:C2');
      assert.strictEqual(conflicts[0].conflictType, 'overlap');
    });

    it('should detect multiple conflicts', () => {
      const ranges = ['A1:B2', 'B1:C2', 'A1:A1', 'B2:B2'];
      const conflicts = detectRangeConflicts(ranges);
      // A1:B2 overlaps with B1:C2, A1:A1, and B2:B2
      // B1:C2 overlaps with B2:B2
      // A1:A1 is contained in A1:B2
      // B2:B2 overlaps with multiple ranges
      assert.ok(conflicts.length >= 4);
    });

    it('should handle empty ranges array', () => {
      const conflicts = detectRangeConflicts([]);
      assert.strictEqual(conflicts.length, 0);
    });

    it('should handle single range', () => {
      const conflicts = detectRangeConflicts(['A1:B2']);
      assert.strictEqual(conflicts.length, 0);
    });
  });

  describe('getRangeIntersection', () => {
    it('should find intersection of overlapping ranges', () => {
      assert.strictEqual(getRangeIntersection('A1:C3', 'B2:D4'), 'B2:C3');
      assert.strictEqual(getRangeIntersection('A1:B2', 'A1:B2'), 'A1:B2');
      assert.strictEqual(getRangeIntersection('A1:C1', 'B1:D1'), 'B1:C1');
    });

    it('should return single cell intersection', () => {
      assert.strictEqual(getRangeIntersection('A1:A1', 'A1:A1'), 'A1');
      assert.strictEqual(getRangeIntersection('A1:B2', 'B2:C3'), 'B2');
    });

    it('should return null for non-overlapping ranges', () => {
      assert.strictEqual(getRangeIntersection('A1:A2', 'B1:B2'), null);
      assert.strictEqual(getRangeIntersection('A1:A1', 'B2:B2'), null);
    });

    it('should handle column and row ranges', () => {
      assert.strictEqual(getRangeIntersection('A:A', 'B:B'), null);
      assert.strictEqual(getRangeIntersection('A:C', 'B:D'), 'B1:C10000000');
      assert.strictEqual(getRangeIntersection('1:1', '2:2'), null);
      assert.strictEqual(getRangeIntersection('1:3', '2:4'), 'A2:ZZZ3');
    });

    it('should handle invalid ranges gracefully', () => {
      assert.strictEqual(getRangeIntersection('INVALID', 'A1:B2'), null);
      assert.strictEqual(getRangeIntersection('A1:B2', 'INVALID'), null);
    });
  });

  describe('validateBatchRanges', () => {
    it('should validate correct batch of ranges', () => {
      const ranges = ['A1:A2', 'B1:B2', 'C1:C2'];
      const result = validateBatchRanges(ranges);
      assert.strictEqual(result.valid, true);
      assert.strictEqual(result.errors.length, 0);
      assert.strictEqual(result.totalCells, 6);
    });

    it('should detect invalid ranges', () => {
      const ranges = ['A1:A2', 'INVALID', 'C1:C2'];
      const result = validateBatchRanges(ranges);
      assert.strictEqual(result.valid, false);
      assert.ok(result.errors.length > 0);
      assert.ok(result.errors[0], 'Expected errors[0] to exist');
      assert.ok(result.errors[0].includes('INVALID'));
    });

    it('should warn about large ranges', () => {
      const ranges = ['A1:ZZZ1000']; // Very large range
      const result = validateBatchRanges(ranges);
      assert.ok(result.warnings.length > 0);
      assert.ok(result.warnings[0], 'Expected warnings[0] to exist');
      assert.ok(result.warnings[0].includes('may impact performance'));
    });

    it('should detect too many ranges', () => {
      const ranges = Array(1001).fill('A1:A1'); // Exceed MAX_BATCH_REQUESTS
      const result = validateBatchRanges(ranges);
      assert.strictEqual(result.valid, false);
      assert.ok(result.errors.some((e) => e.includes('Too many ranges')));
    });

    it('should detect cell limit exceeded', () => {
      // This would create a range with too many cells
      const ranges = ['A:ZZZ']; // Full sheet would exceed cell limit
      const result = validateBatchRanges(ranges);
      assert.strictEqual(result.valid, false);
      assert.ok(result.errors.some((e) => e.includes('Total cells exceed limit')));
    });

    it('should detect range conflicts in warnings', () => {
      const ranges = ['A1:B2', 'B1:C2']; // Overlapping ranges
      const result = validateBatchRanges(ranges);
      assert.ok(result.warnings.some((w) => w.includes('overlap')));
    });
  });
});

describe('Additional Range Utilities', () => {
  describe('expandRange', () => {
    it('should expand single cell range', () => {
      assert.strictEqual(expandRange('A1', 2, 2), 'A1:C3');
      assert.strictEqual(expandRange('B2', 1, 1), 'B2:C3');
      assert.strictEqual(expandRange('A1', 0, 0), 'A1');
    });

    it('should expand existing range', () => {
      assert.strictEqual(expandRange('A1:B2', 1, 1), 'A1:C3');
      assert.strictEqual(expandRange('A1:C3', 2, 2), 'A1:E5');
    });

    it('should respect Google Sheets limits', () => {
      const result = expandRange('ZZZ10000000', 1, 1);
      // Should not exceed limits, so expansion should be capped
      assert.strictEqual(result, 'ZZZ10000000:ZZZ10000000');
    });

    it('should not expand non-cell ranges', () => {
      assert.strictEqual(expandRange('A:A', 1, 1), 'A:A');
      assert.strictEqual(expandRange('1:1', 1, 1), '1:1');
      assert.strictEqual(expandRange('A:C', 1, 1), 'A:C');
    });
  });

  describe('splitRangeIntoChunks', () => {
    it('should not split small ranges', () => {
      const chunks = splitRangeIntoChunks('A1:B2', 10);
      assert.deepStrictEqual(chunks, ['A1:B2']);
    });

    it('should split large ranges by rows', () => {
      const chunks = splitRangeIntoChunks('A1:B1000', 100); // 2000 cells, max 100 per chunk
      assert.ok(chunks.length > 1);
      assert.ok(
        chunks.every((chunk) => {
          const dims = calculateRangeDimensions(chunk);
          return dims.cells <= 100;
        })
      );
    });

    it('should split very wide ranges by columns', () => {
      // Create a range that's very wide but not tall
      const chunks = splitRangeIntoChunks('A1:ZZ2', 10); // 52*2=104 cells, max 10 per chunk
      assert.ok(chunks.length > 1);
    });

    it('should handle non-range types', () => {
      const chunks1 = splitRangeIntoChunks('A:A', 1000);
      assert.deepStrictEqual(chunks1, ['A:A']);

      const chunks2 = splitRangeIntoChunks('1:1', 1000);
      assert.deepStrictEqual(chunks2, ['1:1']);
    });

    it('should use default max cells per chunk', () => {
      const chunks = splitRangeIntoChunks('A1:A1'); // Use default maxCellsPerChunk
      assert.deepStrictEqual(chunks, ['A1:A1']);
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

  it('should handle range overlap logic correctly', () => {
    // Test boundary conditions for overlap detection
    const testCases = [
      { range1: 'A1:B2', range2: 'B2:C3', shouldOverlap: true }, // Touch at corner
      { range1: 'A1:B2', range2: 'C1:D2', shouldOverlap: false }, // Adjacent columns
      { range1: 'A1:B2', range2: 'A3:B4', shouldOverlap: false }, // Adjacent rows
      { range1: 'A1:C3', range2: 'B2:B2', shouldOverlap: true }, // Point inside range
      { range1: 'A1:A1', range2: 'A1:A1', shouldOverlap: true }, // Identical single cells
    ];

    for (const testCase of testCases) {
      const result = rangesOverlap(testCase.range1, testCase.range2);
      assert.strictEqual(result, testCase.shouldOverlap, `Overlap test failed for ${testCase.range1} and ${testCase.range2}`);
    }
  });
});
