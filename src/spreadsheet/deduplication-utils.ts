/** Deduplication utilities for sheet data operations */

/**
 * Build deduplication key from row values with support for both named columns and indices
 *
 * @param row - Row data (array aligned with headers or raw data)
 * @param keyColumns - Column references: strings (header names) or numbers (0-based indices)
 * @param headers - Header names (for column name lookup, null when using indices only)
 * @param hasHeaders - Whether headers are present (affects validation)
 * @returns Composite key string (null-safe, joined with '::')
 *
 * @example
 * // With headers
 * buildDeduplicationKey(
 *   ['john@example.com', 'John', '555-1234'],
 *   ['email', 'phone'],
 *   ['email', 'name', 'phone'],
 *   true
 * )
 * // Returns: 'john@example.com::555-1234'
 *
 * @example
 * // Without headers (using indices)
 * buildDeduplicationKey(
 *   ['value1', 'value2', 'value3'],
 *   [0, 2],
 *   null,
 *   false
 * )
 * // Returns: 'value1::value3'
 */
export function buildDeduplicationKey(row: (string | number | boolean | null)[], keyColumns: (string | number)[], headers: string[] | null, hasHeaders: boolean): string {
  const keyParts = keyColumns.map((colRef) => {
    let colIndex: number;

    // Resolve column reference to index
    if (typeof colRef === 'number') {
      // Direct numeric index
      colIndex = colRef;
    } else {
      // String header name - requires headers
      if (!hasHeaders || !headers) {
        throw new Error(`buildDeduplicationKey: String column reference "${colRef}" requires hasHeaders=true and headers array`);
      }
      colIndex = headers.indexOf(colRef);
      if (colIndex === -1) {
        throw new Error(`buildDeduplicationKey: Header "${colRef}" not found in [${headers.join(', ')}]`);
      }
    }

    // Get value at resolved index
    if (colIndex >= row.length) {
      return '';
    }
    const value = row[colIndex];
    return value === null || value === undefined ? '' : String(value);
  });

  return keyParts.join('::');
}
