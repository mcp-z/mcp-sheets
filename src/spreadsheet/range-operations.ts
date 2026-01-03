/**
 * Range Operations Utilities for Google Sheets
 *
 * This module provides comprehensive utilities for working with A1 notation,
 * range parsing, batch operations, and range validation in Google Sheets.
 *
 * Key features:
 * - A1 notation validation and parsing
 * - Range manipulation and calculation utilities
 * - Batch operation builders for Google Sheets API
 * - Cell count and dimension calculations
 * - Range conflict detection for batch operations
 * - Google Sheets limits validation and enforcement
 */

import { a1Col } from './column-utilities.ts';

// Google Sheets constants and limits
export const GOOGLE_SHEETS_LIMITS = {
  MAX_ROWS: 10_000_000,
  MAX_COLUMNS: 18_278, // ZZZ in base-26
  MAX_CELLS: 10_000_000, // Approximate limit for total cells
  MAX_BATCH_REQUESTS: 1000,
  MAX_DIMENSION_BATCH_REQUESTS: 100,
} as const;

/**
 * Represents a parsed cell reference (e.g., A1, B5)
 */
export interface CellReference {
  column: string;
  columnIndex: number; // 1-based
  row: number; // 1-based
}

/**
 * Represents a parsed range (e.g., A1:B5, A:B, 1:2)
 */
export interface RangeReference {
  type: 'cell' | 'row' | 'column' | 'range';
  startCell?: CellReference;
  endCell?: CellReference;
  startRow?: number;
  endRow?: number;
  startColumn?: string;
  endColumn?: string;
  startColumnIndex?: number;
  endColumnIndex?: number;
}

/**
 * Represents dimensions of a range
 */
export interface RangeDimensions {
  rows: number;
  columns: number;
  cells: number;
}

/**
 * Represents a range conflict between two ranges
 */
export interface RangeConflict {
  range1: string;
  range2: string;
  conflictType: 'overlap' | 'contains' | 'contained' | 'adjacent';
  description: string;
}

/**
 * A1 Notation Validation Utilities
 */

/**
 * Validates if a string is a valid A1 notation
 */
export function isValidA1Notation(notation: string): boolean {
  if (!notation || typeof notation !== 'string') {
    return false;
  }

  // Use the same regex pattern as in the schema
  const a1Pattern = /^(?:[A-Z]{1,3}(?:[1-9]\d{0,6}|10000000)(?::[A-Z]{1,3}(?:[1-9]\d{0,6}|10000000))?|[A-Z]{1,3}:[A-Z]{1,3}|(?:[1-9]\d{0,6}|10000000):(?:[1-9]\d{0,6}|10000000))$/;

  if (!a1Pattern.test(notation)) {
    return false;
  }

  // Additional validation for Google Sheets limits
  const parts = notation.split(':');

  for (const part of parts) {
    // Check cell reference
    const cellMatch = part.match(/^([A-Z]{1,3})([1-9]\d{0,6}|10000000)$/);
    if (cellMatch && cellMatch[1] && cellMatch[2]) {
      const [, colStr, rowStr] = cellMatch;

      // Validate column
      const colNum = columnStringToIndex(colStr);
      if (colNum > GOOGLE_SHEETS_LIMITS.MAX_COLUMNS) {
        return false;
      }

      // Validate row
      const rowNum = parseInt(rowStr, 10);
      if (rowNum > GOOGLE_SHEETS_LIMITS.MAX_ROWS) {
        return false;
      }
    }

    // Check column reference
    const colMatch = part.match(/^([A-Z]{1,3})$/);
    if (colMatch && colMatch[1]) {
      const colNum = columnStringToIndex(colMatch[1]);
      if (colNum > GOOGLE_SHEETS_LIMITS.MAX_COLUMNS) {
        return false;
      }
    }

    // Check row reference
    const rowMatch = part.match(/^([1-9]\d{0,6}|10000000)$/);
    if (rowMatch && rowMatch[1]) {
      const rowNum = parseInt(rowMatch[1], 10);
      if (rowNum > GOOGLE_SHEETS_LIMITS.MAX_ROWS) {
        return false;
      }
    }
  }

  return true;
}

/**
 * Validates A1 notation and throws detailed error if invalid
 */
export function validateA1Notation(notation: string): void {
  if (!isValidA1Notation(notation)) {
    throw new Error(`Invalid A1 notation: "${notation}". Valid formats: A1, A1:B2, A:B, 1:2`);
  }
}

/**
 * Range Parsing and Manipulation Functions
 */

/**
 * Converts column string (A, B, AA, etc.) to 1-based index
 */
export function columnStringToIndex(colStr: string): number {
  let colNum = 0;
  for (let i = 0; i < colStr.length; i++) {
    colNum = colNum * 26 + (colStr.charCodeAt(i) - 64);
  }
  return colNum;
}

/**
 * Converts 1-based column index to column string (A, B, AA, etc.)
 */
export function columnIndexToString(colIndex: number): string {
  return a1Col(colIndex);
}

/**
 * Parses a cell reference (e.g., "A1", "Z999") into its components
 */
export function parseCellReference(cellRef: string): CellReference {
  const match = cellRef.match(/^([A-Z]{1,3})([1-9]\d{0,6}|10000000)$/);
  if (!match || !match[1] || !match[2]) {
    throw new Error(`Invalid cell reference: ${cellRef}`);
  }

  const [, column, rowStr] = match;
  const row = parseInt(rowStr, 10);
  const columnIndex = columnStringToIndex(column);

  return {
    column,
    columnIndex,
    row,
  };
}

/**
 * Parses A1 notation into a structured range reference
 */
export function parseA1Notation(notation: string): RangeReference {
  validateA1Notation(notation);

  const parts = notation.split(':');

  if (parts.length === 1) {
    const part = parts[0];
    if (!part) {
      throw new Error('Invalid A1 notation: empty part after split');
    }

    // Single cell reference (e.g., A1)
    const cellMatch = part.match(/^([A-Z]{1,3})([1-9]\d{0,6}|10000000)$/);
    if (cellMatch && cellMatch[1] && cellMatch[2]) {
      const startCell = parseCellReference(part);
      return {
        type: 'cell',
        startCell,
        endCell: startCell,
      };
    }

    // Single column reference (e.g., A)
    const colMatch = part.match(/^([A-Z]{1,3})$/);
    if (colMatch && colMatch[1]) {
      const column = colMatch[1];
      const columnIndex = columnStringToIndex(column);
      return {
        type: 'column',
        startColumn: column,
        endColumn: column,
        startColumnIndex: columnIndex,
        endColumnIndex: columnIndex,
      };
    }

    // Single row reference (e.g., 1)
    const rowMatch = part.match(/^([1-9]\d{0,6}|10000000)$/);
    if (rowMatch && rowMatch[1]) {
      const row = parseInt(rowMatch[1], 10);
      return {
        type: 'row',
        startRow: row,
        endRow: row,
      };
    }
  }

  if (parts.length === 2 && parts[0] && parts[1]) {
    const [start, end] = parts;

    // Cell range (e.g., A1:B5)
    const startCellMatch = start.match(/^([A-Z]{1,3})([1-9]\d{0,6}|10000000)$/);
    const endCellMatch = end.match(/^([A-Z]{1,3})([1-9]\d{0,6}|10000000)$/);

    if (startCellMatch && endCellMatch && startCellMatch[1] && startCellMatch[2] && endCellMatch[1] && endCellMatch[2]) {
      const startCell = parseCellReference(start);
      const endCell = parseCellReference(end);
      return {
        type: 'range',
        startCell,
        endCell,
      };
    }

    // Column range (e.g., A:B)
    const startColMatch = start.match(/^([A-Z]{1,3})$/);
    const endColMatch = end.match(/^([A-Z]{1,3})$/);

    if (startColMatch && endColMatch && startColMatch[1] && endColMatch[1]) {
      const startColumn = startColMatch[1];
      const endColumn = endColMatch[1];
      const startColumnIndex = columnStringToIndex(startColumn);
      const endColumnIndex = columnStringToIndex(endColumn);
      return {
        type: 'column',
        startColumn,
        endColumn,
        startColumnIndex,
        endColumnIndex,
      };
    }

    // Row range (e.g., 1:5)
    const startRowMatch = start.match(/^([1-9]\d{0,6}|10000000)$/);
    const endRowMatch = end.match(/^([1-9]\d{0,6}|10000000)$/);

    if (startRowMatch && endRowMatch && startRowMatch[1] && endRowMatch[1]) {
      const startRow = parseInt(startRowMatch[1], 10);
      const endRow = parseInt(endRowMatch[1], 10);
      return {
        type: 'row',
        startRow,
        endRow,
      };
    }
  }

  throw new Error(`Unable to parse A1 notation: ${notation}`);
}

/**
 * Converts a range reference back to A1 notation
 */
export function rangeToA1Notation(range: RangeReference): string {
  switch (range.type) {
    case 'cell':
      if (!range.startCell) throw new Error('Invalid cell range: missing startCell');
      return `${range.startCell.column}${range.startCell.row}`;

    case 'range':
      if (!range.startCell || !range.endCell) throw new Error('Invalid range: missing start or end cell');
      return `${range.startCell.column}${range.startCell.row}:${range.endCell.column}${range.endCell.row}`;

    case 'column':
      if (!range.startColumn || !range.endColumn) throw new Error('Invalid column range: missing start or end column');
      if (range.startColumn === range.endColumn) {
        return range.startColumn;
      }
      return `${range.startColumn}:${range.endColumn}`;

    case 'row':
      if (range.startRow === undefined || range.endRow === undefined) throw new Error('Invalid row range: missing start or end row');
      if (range.startRow === range.endRow) {
        return range.startRow.toString();
      }
      return `${range.startRow}:${range.endRow}`;

    default:
      throw new Error(`Unknown range type: ${(range as { type: unknown }).type}`);
  }
}

/**
 * Cell Count and Dimension Calculation Utilities
 */

/**
 * Calculates the dimensions of a range
 */
export function calculateRangeDimensions(notation: string): RangeDimensions {
  const range = parseA1Notation(notation);

  switch (range.type) {
    case 'cell':
      return { rows: 1, columns: 1, cells: 1 };

    case 'range': {
      if (!range.startCell || !range.endCell) throw new Error('Invalid range: missing cells');
      const rows = range.endCell.row - range.startCell.row + 1;
      const columns = range.endCell.columnIndex - range.startCell.columnIndex + 1;
      return { rows, columns, cells: rows * columns };
    }

    case 'column': {
      if (range.startColumnIndex === undefined || range.endColumnIndex === undefined) {
        throw new Error('Invalid column range: missing column indices');
      }
      const columnCount = range.endColumnIndex - range.startColumnIndex + 1;
      return {
        rows: GOOGLE_SHEETS_LIMITS.MAX_ROWS,
        columns: columnCount,
        cells: GOOGLE_SHEETS_LIMITS.MAX_ROWS * columnCount,
      };
    }

    case 'row': {
      if (range.startRow === undefined || range.endRow === undefined) {
        throw new Error('Invalid row range: missing row numbers');
      }
      const rowCount = range.endRow - range.startRow + 1;
      return {
        rows: rowCount,
        columns: GOOGLE_SHEETS_LIMITS.MAX_COLUMNS,
        cells: rowCount * GOOGLE_SHEETS_LIMITS.MAX_COLUMNS,
      };
    }

    default:
      throw new Error(`Unknown range type: ${(range as { type: string }).type}`);
  }
}

/**
 * Calculates total cells affected by multiple ranges
 */
export function calculateTotalCells(ranges: string[]): number {
  return ranges.reduce((total, range) => {
    const dimensions = calculateRangeDimensions(range);
    return total + dimensions.cells;
  }, 0);
}

/**
 * Batch Operation Builders for Google API
 */

/**
 * Builds a values batch update request for Google Sheets API
 */
export function buildValuesBatchUpdateRequest(
  requests: Array<{
    range: string;
    values: (string | number | boolean | null | undefined)[][];
    majorDimension?: 'ROWS' | 'COLUMNS';
  }>,
  sheetTitle: string,
  options: {
    valueInputOption?: 'RAW' | 'USER_ENTERED';
    includeValuesInResponse?: boolean;
    responseDateTimeRenderOption?: 'FORMATTED_STRING' | 'SERIAL_NUMBER';
    responseValueRenderOption?: 'FORMATTED_VALUE' | 'UNFORMATTED_VALUE' | 'FORMULA';
  } = {}
) {
  // Validate all ranges first
  requests.forEach((req, index) => {
    try {
      validateA1Notation(req.range);
    } catch (error) {
      throw new Error(`Invalid range in request ${index}: ${error}`);
    }
  });

  // Calculate total cells for validation
  const totalCells = calculateTotalCells(requests.map((r) => r.range));
  if (totalCells > GOOGLE_SHEETS_LIMITS.MAX_CELLS) {
    throw new Error(`Batch update exceeds maximum cells limit: ${totalCells} > ${GOOGLE_SHEETS_LIMITS.MAX_CELLS}`);
  }

  // Build the request
  const data = requests.map((req) => ({
    range: `${sheetTitle}!${req.range}`,
    values: req.values,
    majorDimension: req.majorDimension || 'ROWS',
  }));

  return {
    valueInputOption: options.valueInputOption || 'USER_ENTERED',
    data,
    includeValuesInResponse: options.includeValuesInResponse || false,
    responseDateTimeRenderOption: options.responseDateTimeRenderOption || 'FORMATTED_STRING',
    responseValueRenderOption: options.responseValueRenderOption || 'FORMATTED_VALUE',
  };
}

/**
 * Range Conflict Detection
 */

/**
 * Checks if two ranges overlap
 */
export function rangesOverlap(range1: string, range2: string): boolean {
  try {
    const parsed1 = parseA1Notation(range1);
    const parsed2 = parseA1Notation(range2);

    // Different types might still overlap, so we need to normalize to cell ranges
    const normalized1 = normalizeRangeToBounds(parsed1);
    const normalized2 = normalizeRangeToBounds(parsed2);

    // Check for overlap
    return !(normalized1.endRow < normalized2.startRow || normalized2.endRow < normalized1.startRow || normalized1.endCol < normalized2.startCol || normalized2.endCol < normalized1.startCol);
  } catch {
    // If we can't parse the ranges, assume no overlap
    return false;
  }
}

/**
 * Normalizes a range to bounds for overlap checking
 */
function normalizeRangeToBounds(range: RangeReference): {
  startRow: number;
  endRow: number;
  startCol: number;
  endCol: number;
} {
  switch (range.type) {
    case 'cell':
      if (!range.startCell) throw new Error('Invalid cell range');
      return {
        startRow: range.startCell.row,
        endRow: range.startCell.row,
        startCol: range.startCell.columnIndex,
        endCol: range.startCell.columnIndex,
      };

    case 'range':
      if (!range.startCell || !range.endCell) throw new Error('Invalid range');
      return {
        startRow: range.startCell.row,
        endRow: range.endCell.row,
        startCol: range.startCell.columnIndex,
        endCol: range.endCell.columnIndex,
      };

    case 'column':
      if (range.startColumnIndex === undefined || range.endColumnIndex === undefined) {
        throw new Error('Invalid column range');
      }
      return {
        startRow: 1,
        endRow: GOOGLE_SHEETS_LIMITS.MAX_ROWS,
        startCol: range.startColumnIndex,
        endCol: range.endColumnIndex,
      };

    case 'row':
      if (range.startRow === undefined || range.endRow === undefined) {
        throw new Error('Invalid row range');
      }
      return {
        startRow: range.startRow,
        endRow: range.endRow,
        startCol: 1,
        endCol: GOOGLE_SHEETS_LIMITS.MAX_COLUMNS,
      };

    default:
      throw new Error(`Unknown range type: ${(range as { type: unknown }).type}`);
  }
}

/**
 * Detects conflicts between multiple ranges
 */
export function detectRangeConflicts(ranges: string[]): RangeConflict[] {
  const conflicts: RangeConflict[] = [];

  for (let i = 0; i < ranges.length; i++) {
    for (let j = i + 1; j < ranges.length; j++) {
      const range1 = ranges[i];
      const range2 = ranges[j];

      if (!range1 || !range2) continue;

      if (rangesOverlap(range1, range2)) {
        conflicts.push({
          range1,
          range2,
          conflictType: 'overlap',
          description: `Ranges ${range1} and ${range2} overlap`,
        });
      }
    }
  }

  return conflicts;
}

/**
 * Validates that a batch of ranges doesn't exceed Google Sheets limits
 */
export function validateBatchRanges(ranges: string[]): {
  valid: boolean;
  errors: string[];
  warnings: string[];
  totalCells: number;
} {
  const errors: string[] = [];
  const warnings: string[] = [];
  let totalCells = 0;

  // Validate each range
  ranges.forEach((range, index) => {
    try {
      validateA1Notation(range);
      const dimensions = calculateRangeDimensions(range);
      totalCells += dimensions.cells;

      // Check for very large ranges that might cause performance issues
      if (dimensions.cells > 1_000_000) {
        warnings.push(`Range ${index + 1} (${range}) affects ${dimensions.cells.toLocaleString()} cells, which may impact performance`);
      }
    } catch (error) {
      errors.push(`Range ${index + 1} (${range}): ${error}`);
    }
  });

  // Check batch size limits
  if (ranges.length > GOOGLE_SHEETS_LIMITS.MAX_BATCH_REQUESTS) {
    errors.push(`Too many ranges: ${ranges.length} > ${GOOGLE_SHEETS_LIMITS.MAX_BATCH_REQUESTS}`);
  }

  // Check total cells limit
  if (totalCells > GOOGLE_SHEETS_LIMITS.MAX_CELLS) {
    errors.push(`Total cells exceed limit: ${totalCells.toLocaleString()} > ${GOOGLE_SHEETS_LIMITS.MAX_CELLS.toLocaleString()}`);
  }

  // Check for conflicts
  const conflicts = detectRangeConflicts(ranges);
  conflicts.forEach((conflict) => {
    warnings.push(conflict.description);
  });

  return {
    valid: errors.length === 0,
    errors,
    warnings,
    totalCells,
  };
}

/**
 * Utility Functions for Common Range Operations
 */

/**
 * Expands a range by a specified number of rows and columns
 */
export function expandRange(notation: string, expandRows: number, expandCols: number): string {
  const range = parseA1Notation(notation);

  if (range.type === 'cell' && range.startCell) {
    const newEndRow = Math.min(range.startCell.row + expandRows, GOOGLE_SHEETS_LIMITS.MAX_ROWS);
    const newEndCol = Math.min(range.startCell.columnIndex + expandCols, GOOGLE_SHEETS_LIMITS.MAX_COLUMNS);
    const newEndColStr = columnIndexToString(newEndCol);

    if (expandRows === 0 && expandCols === 0) {
      return notation;
    }

    return `${range.startCell.column}${range.startCell.row}:${newEndColStr}${newEndRow}`;
  }

  if (range.type === 'range' && range.startCell && range.endCell) {
    const newEndRow = Math.min(range.endCell.row + expandRows, GOOGLE_SHEETS_LIMITS.MAX_ROWS);
    const newEndCol = Math.min(range.endCell.columnIndex + expandCols, GOOGLE_SHEETS_LIMITS.MAX_COLUMNS);
    const newEndColStr = columnIndexToString(newEndCol);

    return `${range.startCell.column}${range.startCell.row}:${newEndColStr}${newEndRow}`;
  }

  // For row and column ranges, expansion doesn't make sense in the same way
  return notation;
}

/**
 * Gets the intersection of two ranges
 */
export function getRangeIntersection(range1: string, range2: string): string | null {
  try {
    const parsed1 = parseA1Notation(range1);
    const parsed2 = parseA1Notation(range2);

    const bounds1 = normalizeRangeToBounds(parsed1);
    const bounds2 = normalizeRangeToBounds(parsed2);

    // Check if they overlap
    if (!rangesOverlap(range1, range2)) {
      return null;
    }

    // Calculate intersection bounds
    const startRow = Math.max(bounds1.startRow, bounds2.startRow);
    const endRow = Math.min(bounds1.endRow, bounds2.endRow);
    const startCol = Math.max(bounds1.startCol, bounds2.startCol);
    const endCol = Math.min(bounds1.endCol, bounds2.endCol);

    const startColStr = columnIndexToString(startCol);
    const endColStr = columnIndexToString(endCol);

    if (startRow === endRow && startCol === endCol) {
      return `${startColStr}${startRow}`;
    }

    return `${startColStr}${startRow}:${endColStr}${endRow}`;
  } catch {
    return null;
  }
}

/**
 * Splits a large range into smaller chunks for batch processing
 */
export function splitRangeIntoChunks(notation: string, maxCellsPerChunk = 100_000): string[] {
  const range = parseA1Notation(notation);
  const dimensions = calculateRangeDimensions(notation);

  if (dimensions.cells <= maxCellsPerChunk) {
    return [notation];
  }

  const chunks: string[] = [];

  if (range.type === 'range' && range.startCell && range.endCell) {
    const totalRows = range.endCell.row - range.startCell.row + 1;
    const totalCols = range.endCell.columnIndex - range.startCell.columnIndex + 1;

    // Calculate chunk size based on max cells
    const cellsPerRow = totalCols;
    const maxRowsPerChunk = Math.floor(maxCellsPerChunk / cellsPerRow);

    if (maxRowsPerChunk >= 1) {
      // Split by rows
      for (let currentRow = range.startCell.row; currentRow <= range.endCell.row; currentRow += maxRowsPerChunk) {
        const chunkEndRow = Math.min(currentRow + maxRowsPerChunk - 1, range.endCell.row);
        const chunkRange = `${range.startCell.column}${currentRow}:${range.endCell.column}${chunkEndRow}`;
        chunks.push(chunkRange);
      }
    } else {
      // If even one row is too big, split by columns
      const maxColsPerChunk = Math.floor(maxCellsPerChunk / totalRows);

      for (let currentCol = range.startCell.columnIndex; currentCol <= range.endCell.columnIndex; currentCol += maxColsPerChunk) {
        const chunkEndCol = Math.min(currentCol + maxColsPerChunk - 1, range.endCell.columnIndex);
        const startColStr = columnIndexToString(currentCol);
        const endColStr = columnIndexToString(chunkEndCol);
        const chunkRange = `${startColStr}${range.startCell.row}:${endColStr}${range.endCell.row}`;
        chunks.push(chunkRange);
      }
    }
  } else {
    // For other range types, just return the original range
    // (column and row ranges are already at their limits)
    chunks.push(notation);
  }

  return chunks;
}

/**
 * Converts a RangeReference to Google Sheets API GridRange format
 *
 * This handles the different range types correctly:
 * - 'cell' and 'range' types: Extract indices from startCell/endCell
 * - 'row' types: Use startRow/endRow, omit column indices for full row
 * - 'column' types: Use startColumnIndex/endColumnIndex, omit row indices for full column
 *
 * Important: Google Sheets API uses 0-based indices, while A1 notation uses 1-based indices.
 * This function handles the conversion properly for each range type.
 */
export function rangeReferenceToGridRange(
  rangeRef: RangeReference,
  sheetId: number
): {
  sheetId: number;
  startRowIndex?: number;
  endRowIndex?: number;
  startColumnIndex?: number;
  endColumnIndex?: number;
} {
  if (rangeRef.type === 'cell' || rangeRef.type === 'range') {
    // For cell and range types, extract from startCell/endCell
    if (!rangeRef.startCell || !rangeRef.endCell) {
      throw new Error(`Invalid ${rangeRef.type} range: missing start or end cell`);
    }

    return {
      sheetId,
      // Convert 1-based row to 0-based startRowIndex
      startRowIndex: rangeRef.startCell.row - 1,
      // endRowIndex is exclusive (so row 5 becomes endRowIndex 5)
      endRowIndex: rangeRef.endCell.row,
      // Convert 1-based column index to 0-based startColumnIndex
      startColumnIndex: rangeRef.startCell.columnIndex - 1,
      // endColumnIndex is exclusive
      endColumnIndex: rangeRef.endCell.columnIndex,
    };
  }
  if (rangeRef.type === 'row') {
    // For row types, use startRow/endRow and omit column indices for full row width
    if (rangeRef.startRow === undefined || rangeRef.endRow === undefined) {
      throw new Error('Invalid row range: missing start or end row');
    }

    return {
      sheetId,
      // Convert 1-based row to 0-based startRowIndex
      startRowIndex: rangeRef.startRow - 1,
      // endRowIndex is exclusive
      endRowIndex: rangeRef.endRow,
      // Omit startColumnIndex and endColumnIndex to apply to all columns
    };
  }
  if (rangeRef.type === 'column') {
    // For column types, use startColumnIndex/endColumnIndex and omit row indices for full column height
    if (rangeRef.startColumnIndex === undefined || rangeRef.endColumnIndex === undefined) {
      throw new Error('Invalid column range: missing start or end column index');
    }

    return {
      sheetId,
      // Omit startRowIndex and endRowIndex to apply to all rows
      // Convert 1-based column index to 0-based startColumnIndex
      startColumnIndex: rangeRef.startColumnIndex - 1,
      // endColumnIndex is exclusive
      endColumnIndex: rangeRef.endColumnIndex,
    };
  }

  throw new Error(`Unknown range type: ${(rangeRef as { type: unknown }).type}`);
}
