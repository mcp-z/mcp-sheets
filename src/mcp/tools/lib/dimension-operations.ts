// Google Sheets dimension constants
export const DEFAULT_APPEND_COUNT = 1; // Default number of rows/columns to append
export const DEFAULT_ROW_COUNT = 1000; // Google Sheets default for new sheets
export const DEFAULT_COLUMN_COUNT = 26; // Google Sheets default for new sheets (A-Z)
export const MAX_ROW_COUNT = 10000000; // Google Sheets maximum
export const MAX_COLUMN_COUNT = 18278; // Google Sheets maximum (ZZZ in base-26)

// Type definitions for dimension operations
export type DimensionOperation = 'insertDimension' | 'deleteDimension' | 'appendDimension';
export type DimensionType = 'ROWS' | 'COLUMNS';

export interface DimensionRequest {
  operation: DimensionOperation;
  dimension: DimensionType;
  startIndex: number;
  endIndex?: number;
  inheritFromBefore?: boolean;
}

// Helper function to sort operations for optimal execution order
export function sortOperations(requests: DimensionRequest[]): DimensionRequest[] {
  // Sort by: 1) operation type (delete -> insert -> append), 2) proper index ordering to prevent conflicts
  return [...requests].sort((a, b) => {
    // Operation type priority: deleteDimension (0), insertDimension (1), appendDimension (2)
    const operationPriority = {
      deleteDimension: 0,
      insertDimension: 1,
      appendDimension: 2,
    };

    const aPriority = operationPriority[a.operation];
    const bPriority = operationPriority[b.operation];

    if (aPriority !== bPriority) {
      return aPriority - bPriority;
    }

    // Within same operation type, sort by appropriate index
    if (a.operation === 'deleteDimension') {
      // For deletes, process higher START indices first to avoid index shifting issues
      // This ensures we delete from the end towards the beginning
      const aStart = a.startIndex;
      const bStart = b.startIndex;
      if (aStart !== bStart) {
        return bStart - aStart; // Higher start index first
      }
      // If start indices are equal, process higher end index first for consistency
      const aEnd = a.endIndex ?? a.startIndex + 1;
      const bEnd = b.endIndex ?? b.startIndex + 1;
      return bEnd - aEnd;
    }
    // For inserts and appends, process lower indices first
    // This ensures we modify from the beginning towards the end
    return a.startIndex - b.startIndex;
  });
}

// Helper function to build Google Sheets API dimension request
export function buildDimensionRequest(operation: DimensionRequest, sheetId: number) {
  const { operation: operationType, dimension, startIndex, endIndex, inheritFromBefore } = operation;

  const dimensionRange = {
    sheetId,
    dimension,
    startIndex,
    ...(endIndex !== undefined && { endIndex }),
  };

  switch (operationType) {
    case 'insertDimension':
      return {
        insertDimension: {
          range: dimensionRange,
          inheritFromBefore: inheritFromBefore ?? false,
        },
      };
    case 'deleteDimension':
      return {
        deleteDimension: {
          range: dimensionRange,
        },
      };
    case 'appendDimension':
      return {
        appendDimension: {
          sheetId,
          dimension,
          length: DEFAULT_APPEND_COUNT,
        },
      };
    default:
      throw new Error(`Unsupported operation type: ${operationType}`);
  }
}

// Helper function to calculate affected count for each operation
export function calculateAffectedCount(operation: DimensionRequest): number {
  const { operation: operationType, startIndex, endIndex } = operation;

  switch (operationType) {
    case 'insertDimension':
    case 'deleteDimension':
      return endIndex !== undefined ? endIndex - startIndex : 1;
    case 'appendDimension':
      return DEFAULT_APPEND_COUNT;
    default:
      return 0;
  }
}
