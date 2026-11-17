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
 * Counts cells in a finite sheet-local range, rejecting reversed bounds.
 */
export function calculateFiniteRangeCellCount(notation: string): number {
  const range = parseA1Notation(notation);
  if (range.type !== 'cell' && range.type !== 'range') {
    throw new Error(`Range must have finite start and end rows and columns: ${notation}`);
  }
  const dimensions = calculateRangeDimensions(notation);
  if (dimensions.rows < 1 || dimensions.columns < 1) throw new Error(`Range end must not precede its start: ${notation}`);
  return dimensions.cells;
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

/**
 * A finite, sheet-local block of cells with 1-based inclusive bounds.
 */
export interface CellBlock {
  startRow: number;
  endRow: number;
  startColumn: number;
  endColumn: number;
}

/**
 * Parses a finite sheet-local A1 range (cell or range) into 1-based inclusive bounds.
 * Throws for invalid syntax, whole rows/columns and reversed bounds.
 */
export function parseFiniteCellBlock(notation: string): CellBlock {
  const range = parseA1Notation(notation);
  if (range.type !== 'cell' && range.type !== 'range') {
    throw new Error(`Range must have finite start and end rows and columns: ${notation}`);
  }
  calculateFiniteRangeCellCount(notation);
  if (!range.startCell || !range.endCell) throw new Error(`Invalid range: ${notation}`);
  return { startRow: range.startCell.row, endRow: range.endCell.row, startColumn: range.startCell.columnIndex, endColumn: range.endCell.columnIndex };
}

/**
 * Formats a block as A1 notation (a single cell when it spans one cell).
 */
export function cellBlockToA1(block: CellBlock): string {
  const start = `${columnIndexToString(block.startColumn)}${block.startRow}`;
  if (block.startRow === block.endRow && block.startColumn === block.endColumn) return start;
  return `${start}:${columnIndexToString(block.endColumn)}${block.endRow}`;
}

/**
 * Footprint of a copy-paste: the cells written in the destination.
 *
 * - destination is one cell: the (oriented) source size anchored at that cell
 * - destination is a range whose height and width are whole multiples of the oriented source: the destination itself (the source is tiled)
 * - anything else throws
 */
export function calculatePasteRange(source: string, destination: string, orientation: 'NORMAL' | 'TRANSPOSE'): CellBlock {
  const src = parseFiniteCellBlock(source);
  const dest = parseFiniteCellBlock(destination);
  const srcRows = src.endRow - src.startRow + 1;
  const srcColumns = src.endColumn - src.startColumn + 1;
  const rows = orientation === 'TRANSPOSE' ? srcColumns : srcRows;
  const columns = orientation === 'TRANSPOSE' ? srcRows : srcColumns;
  const destRows = dest.endRow - dest.startRow + 1;
  const destColumns = dest.endColumn - dest.startColumn + 1;

  if (destRows === 1 && destColumns === 1) {
    return { startRow: dest.startRow, endRow: dest.startRow + rows - 1, startColumn: dest.startColumn, endColumn: dest.startColumn + columns - 1 };
  }
  if (destRows % rows === 0 && destColumns % columns === 0) return dest;
  throw new Error(
    `Destination ${destination} is ${destRows}x${destColumns} cells; the ${orientation === 'TRANSPOSE' ? 'transposed ' : ''}source ${source} is ${rows}x${columns}. Use a single anchor cell (e.g., ${columnIndexToString(dest.startColumn)}${dest.startRow}) or a range whose height and width are whole multiples of ${rows}x${columns}.`
  );
}
