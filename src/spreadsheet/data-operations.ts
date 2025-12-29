import type { sheets_v4 } from 'googleapis';
import type { Logger } from '../types.js';
import { a1Col } from './column-utilities.js';
import { findSheetByRef } from './sheet-operations.js';

// Core data types
export type Cell = string | number | boolean | null | undefined;
export type Row = Cell[];

// Enhanced type definitions for shared data operations
export interface ColumnMapping {
  canonical: string;
  sheet: string;
  index: number;
}

export interface HeaderValidationResult {
  valid: boolean;
  missingColumns: string[];
  extraColumns: string[];
  mappings: ColumnMapping[];
}

export interface KeyGenerationStrategy {
  keyColumns: string[];
  useProviderIdLogic: boolean;
  separator?: string;
}

export interface DataPartition {
  toAppend: Row[];
  toUpdate: Array<{ row: Row; existingRowIndex: number }>;
  skippedKeys: string[];
}

export interface BatchOperationResult {
  updatedRows: number;
  inserted: string[];
  rowsSkipped: number;
  errors?: string[];
}

export interface UpsertOptions {
  keyStrategy: KeyGenerationStrategy;
  allowUpdates: boolean;
  batchSize?: number;
  valueInputOption?: 'RAW' | 'USER_ENTERED';
}

export async function discoverHeader(sheets: sheets_v4.Sheets, spreadsheetId: string, sheetTitle: string): Promise<string[]> {
  try {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${sheetTitle}!1:1`,
      majorDimension: 'ROWS',
    });
    return ((response.data?.values || [])[0] || []) as string[];
  } catch {
    return [];
  }
}

export function validateAndMapHeaders(sheetHeaders: string[], canonicalHeaders: string[], requiredColumns: string[] = []): HeaderValidationResult {
  const normalizeColumn = (col: string) =>
    String(col ?? '')
      .toLowerCase()
      .replace(/[^a-z0-9]/g, '');

  const sheetNormalized = sheetHeaders.map(normalizeColumn);
  const canonicalNormalized = canonicalHeaders.map(normalizeColumn);
  const requiredNormalized = requiredColumns.map(normalizeColumn);

  const mappings: ColumnMapping[] = [];
  const missingColumns: string[] = [];
  const extraColumns: string[] = [...sheetHeaders];

  // Create mappings for canonical columns
  canonicalHeaders.forEach((canonical, canonicalIndex) => {
    const canonicalNorm = canonicalNormalized[canonicalIndex];
    if (!canonicalNorm) return; // Skip if normalization resulted in empty string

    let sheetIndex = sheetNormalized.indexOf(canonicalNorm);

    // Special handling for 'id' -> 'messageid' mapping
    if (sheetIndex === -1 && canonicalNorm === 'id') {
      sheetIndex = sheetNormalized.indexOf('messageid');
    }

    if (sheetIndex !== -1) {
      const sheetHeader = sheetHeaders[sheetIndex];
      if (sheetHeader !== undefined) {
        mappings.push({
          canonical,
          sheet: sheetHeader,
          index: sheetIndex,
        });
        // Remove from extra columns
        const extraIndex = extraColumns.indexOf(sheetHeader);
        if (extraIndex !== -1) {
          extraColumns.splice(extraIndex, 1);
        }
      }
    } else if (requiredNormalized.includes(canonicalNorm)) {
      missingColumns.push(canonical);
    }
  });

  return {
    valid: missingColumns.length === 0,
    missingColumns,
    extraColumns,
    mappings,
  };
}

/** Ensures consistent key generation across all functions */
export function generateRowKey(row: Row, header: string[], strategy: KeyGenerationStrategy): string {
  const { keyColumns, useProviderIdLogic, separator = '\\' } = strategy;

  // Normalize header for consistent lookups
  const lowerHeader = header.map((h) => String(h ?? '').toLowerCase());

  if (useProviderIdLogic) {
    const providerIndex = lowerHeader.indexOf('provider');
    let idIndex = lowerHeader.indexOf('messageid');
    if (idIndex === -1) idIndex = lowerHeader.indexOf('id');

    if (providerIndex >= 0 && idIndex >= 0) {
      const providerVal = String(row[providerIndex] ?? '').trim();
      const idVal = String(row[idIndex] ?? '').trim();
      if (providerVal || idVal) {
        // Only return key if at least one component exists
        return [providerVal, idVal].join(separator);
      }
    }
  }

  // Standard key generation using specified columns
  const keyIndices = keyColumns
    .map((name) => {
      const normalizedName = String(name ?? '').toLowerCase();
      let index = lowerHeader.indexOf(normalizedName);
      // Consistent fallback: id -> messageid
      if (index === -1 && normalizedName === 'id') {
        index = lowerHeader.indexOf('messageid');
      }
      return index;
    })
    .filter((index) => index >= 0);

  if (keyIndices.length === 0) {
    return ''; // Return empty string for invalid key configurations
  }

  const components = keyIndices.map((index) => String(row[index] ?? '').trim());
  // Only return key if all components are non-empty
  if (components.every((comp) => comp.length > 0)) {
    return components.join(separator);
  }

  return '';
}

export function validateRowKeys(rows: Row[], header: string[], strategy: KeyGenerationStrategy): { valid: boolean; duplicateKeys: string[]; keyMap: Map<string, number[]> } {
  const keyMap = new Map<string, number[]>();
  const duplicateKeys: string[] = [];

  rows.forEach((row, index) => {
    const key = generateRowKey(row, header, strategy);
    if (key.replace(/\\+/g, '') === '') return; // Skip empty keys

    if (!keyMap.has(key)) {
      keyMap.set(key, []);
    }
    keyMap.get(key)?.push(index);
  });

  // Find duplicates
  keyMap.forEach((indices, key) => {
    if (indices.length > 1) {
      duplicateKeys.push(key);
    }
  });

  return {
    valid: duplicateKeys.length === 0,
    duplicateKeys,
    keyMap,
  };
}

// Overloaded function signatures for appendRows
export async function appendRows(sheets: sheets_v4.Sheets, params: { spreadsheetId: string; sheetTitle: string; rows?: unknown[]; keySet?: Set<string> | null; keyColumns?: string[]; header?: string[]; logger: Logger }): Promise<{ updatedRows: number; inserted: string[]; rowsSkipped?: number }>;

export async function appendRows(sheets: sheets_v4.Sheets, params: { spreadsheetId: string; sheetRef: string; rows?: unknown[]; keySet?: Set<string> | null; keyColumns?: string[]; header?: string[]; logger: Logger }): Promise<{ updatedRows: number; inserted: string[]; rowsSkipped?: number }>;

export async function appendRows(
  sheets: sheets_v4.Sheets,
  {
    spreadsheetId,
    sheetTitle,
    sheetRef,
    rows = [],
    keySet = null as Set<string> | null,
    keyColumns = ['id'] as string[],
    header = [] as string[],
    logger,
  }: {
    spreadsheetId: string;
    sheetTitle?: string;
    sheetRef?: string;
    rows?: unknown[];
    keySet?: Set<string> | null;
    keyColumns?: string[];
    header?: string[];
    logger: Logger;
  }
) {
  if (!sheets) throw new Error('appendRows: sheets is required');
  if (!spreadsheetId) throw new Error('appendRows: spreadsheetId is required');
  if (!sheetTitle && !sheetRef) throw new Error('appendRows: either sheetTitle or sheetRef is required');

  // Resolve the actual sheet title from sheetRef if provided
  let resolvedSheetTitle = sheetTitle;
  if (sheetRef) {
    const sheet = await findSheetByRef(sheets, spreadsheetId, sheetRef, logger);
    if (sheet) {
      resolvedSheetTitle = sheet.properties?.title || sheetRef;
    } else {
      // Sheet doesn't exist, use the sheetRef as the title
      resolvedSheetTitle = sheetRef;
    }
  }

  if (!resolvedSheetTitle) throw new Error('appendRows: could not resolve sheet title');
  if (!Array.isArray(rows) || rows.length === 0) return { updatedRows: 0, inserted: [] as string[] };

  if (!Array.isArray(header) || header.length === 0) {
    const respHeader = await sheets.spreadsheets.values.get({ spreadsheetId, range: `${resolvedSheetTitle}!1:1`, majorDimension: 'ROWS' });
    const vr = respHeader.data as sheets_v4.Schema$ValueRange;
    header = ((vr?.values || [])[0] || []) as string[];
  }

  const resolveKeyIndices = (kc: string[], hdr: string[]) => {
    if (!Array.isArray(kc) || kc.length === 0) return [] as number[];
    for (const c of kc) {
      if (typeof c !== 'string' || c.length === 0) throw new Error('appendRows: keyColumns must be an array of non-empty strings');
    }
    if (!Array.isArray(hdr) || hdr.length === 0) {
      throw new Error('appendRows: header array is required when keyColumns are header names');
    }
    return kc.map((name) => {
      const idx = hdr.indexOf(name);
      if (idx === -1) throw new Error(`appendRows: header name "${name}" not found in header`);
      return idx;
    });
  };

  const keyColsIdx = resolveKeyIndices(keyColumns, header);

  // Internal batch size to avoid Google Sheets limits
  const batchSize = 50;

  if (!keyColsIdx || keyColsIdx.length === 0) {
    let totalUpdated = 0;
    for (let i = 0; i < rows.length; i += batchSize) {
      const batch = rows.slice(i, i + batchSize);
      const resp = await sheets.spreadsheets.values.append({ spreadsheetId, range: `${resolvedSheetTitle}!A1`, valueInputOption: 'RAW', insertDataOption: 'INSERT_ROWS', requestBody: { values: batch as unknown[][] } });
      const app = resp.data as sheets_v4.Schema$AppendValuesResponse;
      const updatedRows = Number(app?.updates?.updatedRows || batch.length);
      totalUpdated += updatedRows;
    }
    return { updatedRows: totalUpdated, inserted: [] as string[], rowsSkipped: 0 };
  }

  const toKey = (r: unknown[]) => {
    const row = r as (string | number | boolean | null | undefined)[];

    // Use consistent key generation strategy
    const strategy: KeyGenerationStrategy = {
      keyColumns,
      useProviderIdLogic: true,
      separator: '\\',
    };

    return generateRowKey(row, header, strategy);
  };

  const rowsToInsert: unknown[][] = [];
  const insertedKeys: string[] = [];
  for (const r of rows as unknown[][]) {
    const key = toKey(r as unknown[]);
    if (keySet && keySet.has(key)) continue;
    rowsToInsert.push(r as unknown[]);
    insertedKeys.push(key);
  }

  if (rowsToInsert.length === 0) {
    const rowsSkipped = rows.length - rowsToInsert.length; // Should be rows.length when all are skipped
    return { updatedRows: 0, inserted: [] as string[], rowsSkipped };
  }

  // Use smaller internal batch size to avoid Google Sheets character limit errors
  // Especially important when dealing with emails that may have long bodies
  const INTERNAL_BATCH_SIZE = 50;

  let totalUpdated = 0;
  for (let i = 0; i < rowsToInsert.length; i += INTERNAL_BATCH_SIZE) {
    const batch = rowsToInsert.slice(i, i + INTERNAL_BATCH_SIZE);
    try {
      const resp = await sheets.spreadsheets.values.append({
        spreadsheetId,
        range: `${resolvedSheetTitle}!A1`,
        valueInputOption: 'RAW',
        insertDataOption: 'INSERT_ROWS',
        requestBody: { values: batch as unknown[][] },
      });
      const app2 = resp.data as sheets_v4.Schema$AppendValuesResponse;
      const updatedRows = Number(app2?.updates?.updatedRows || batch.length);
      totalUpdated += updatedRows;
    } catch (error) {
      // If a batch fails, try with smaller batches to handle partial success
      if (batch.length > 1) {
        for (const singleRow of batch) {
          try {
            const singleResp = await sheets.spreadsheets.values.append({
              spreadsheetId,
              range: `${resolvedSheetTitle}!A1`,
              valueInputOption: 'RAW',
              insertDataOption: 'INSERT_ROWS',
              requestBody: { values: [singleRow as unknown[]] },
            });
            const singleApp = singleResp.data as sheets_v4.Schema$AppendValuesResponse;
            const singleUpdatedRows = Number(singleApp?.updates?.updatedRows || 1);
            totalUpdated += singleUpdatedRows;
          } catch (singleError) {
            // Skip problematic individual rows but continue processing
            logger.warn?.(`Failed to insert single row: ${singleError instanceof Error ? singleError.message : String(singleError)}`);
          }
        }
      } else {
        // Single row batch failed, skip it
        logger.warn?.(`Failed to insert batch: ${error instanceof Error ? error.message : String(error)}`);
      }
    }
  }
  const rowsSkipped = rows.length - rowsToInsert.length;
  return { updatedRows: totalUpdated, inserted: insertedKeys, rowsSkipped };
}

export function mapRowsToHeader({ rows = [], header = [], canonical = [] as string[] }: { rows?: Row[]; header?: string[]; canonical?: string[] }): Row[] {
  if (!Array.isArray(rows) || rows.length === 0) return [];
  if (!Array.isArray(header) || header.length === 0) return rows;
  if (!Array.isArray(canonical) || canonical.length === 0) return rows;

  const validation = validateAndMapHeaders(header, canonical);
  const mappingLookup = new Map(validation.mappings.map((m) => [m.canonical, m.index]));

  return rows.map((row: Row) => {
    // Fill with null to skip unmapped columns (preserve existing values)
    // Use '' (empty string) only when explicitly wanting to clear a cell
    const mappedRow: Row = new Array(header.length).fill(null);

    canonical.forEach((canonicalCol, canonicalIndex) => {
      const sheetIndex = mappingLookup.get(canonicalCol);
      if (sheetIndex !== undefined && canonicalIndex < row.length) {
        // Preserve null values - they signal "skip this cell"
        mappedRow[sheetIndex] = row[canonicalIndex];
      }
    });

    return mappedRow;
  });
}

export async function snapshotHeaderAndKeys(sheets: sheets_v4.Sheets, spreadsheetId: string, sheetTitle: string, keyColumns: string[] = ['id'], keyStrategy?: Partial<KeyGenerationStrategy>): Promise<{ header: string[]; keySet: Set<string>; keyIndices: number[] }> {
  const header = await discoverHeader(sheets, spreadsheetId, sheetTitle);
  const keySet = new Set<string>();

  if (header.length === 0) {
    return { header, keySet, keyIndices: [] };
  }

  const strategy: KeyGenerationStrategy = {
    keyColumns,
    useProviderIdLogic: keyStrategy?.useProviderIdLogic ?? true,
    separator: keyStrategy?.separator ?? '\\',
  };

  // Find key column indices
  const keyIndices = keyColumns
    .map((name) => {
      const normalizedName = String(name ?? '').toLowerCase();
      const lowerHeader = header.map((h) => String(h ?? '').toLowerCase());
      let index = lowerHeader.indexOf(normalizedName);
      if (index === -1 && normalizedName === 'id') {
        index = lowerHeader.indexOf('messageid');
      }
      return index;
    })
    .filter((index) => index >= 0);

  if (keyIndices.length === 0) {
    return { header, keySet, keyIndices };
  }

  // Read key columns data in chunks
  const chunkSize = 1000;
  let startRow = 2;
  const minIndex = Math.min(...keyIndices);
  const maxIndex = Math.max(...keyIndices);
  const startCol = a1Col(minIndex + 1);
  const endCol = a1Col(maxIndex + 1);

  while (true) {
    const endRow = startRow + chunkSize - 1;
    const range = `${sheetTitle}!${startCol}${startRow}:${endCol}${endRow}`;

    try {
      const response = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range,
        majorDimension: 'ROWS',
      });

      const rows = (response.data?.values || []) as string[][];
      if (rows.length === 0) break;

      for (const rawRow of rows) {
        // Reconstruct full row for key generation
        const fullRow: Row = new Array(header.length).fill('');
        keyIndices.forEach((globalIndex, _localIndex) => {
          const localRowIndex = globalIndex - minIndex;
          if (localRowIndex >= 0 && localRowIndex < rawRow.length) {
            fullRow[globalIndex] = rawRow[localRowIndex];
          }
        });

        const key = generateRowKey(fullRow, header, strategy);
        if (key.replace(/\\+/g, '') !== '') {
          keySet.add(key);
        }
      }

      startRow += rows.length;
      if (rows.length < chunkSize) break; // Last chunk
    } catch {
      break; // End of data or error
    }
  }

  return { header, keySet, keyIndices };
}

/** Tracks row positions for updates in addition to keys */
export async function snapshotHeaderKeysAndPositions(sheets: sheets_v4.Sheets, spreadsheetId: string, sheetTitle: string, keyColumns: string[] = ['id'], keyStrategy?: Partial<KeyGenerationStrategy>): Promise<{ header: string[]; keySet: Set<string>; keyToRowMap: Map<string, number>; keyIndices: number[] }> {
  const header = await discoverHeader(sheets, spreadsheetId, sheetTitle);
  const keySet = new Set<string>();
  const keyToRowMap = new Map<string, number>();

  if (header.length === 0) {
    return { header, keySet, keyToRowMap, keyIndices: [] };
  }

  const strategy: KeyGenerationStrategy = {
    keyColumns,
    useProviderIdLogic: keyStrategy?.useProviderIdLogic ?? true,
    separator: keyStrategy?.separator ?? '\\',
  };

  // Find key column indices
  const keyIndices = keyColumns
    .map((name) => {
      const normalizedName = String(name ?? '').toLowerCase();
      const lowerHeader = header.map((h) => String(h ?? '').toLowerCase());
      let index = lowerHeader.indexOf(normalizedName);
      if (index === -1 && normalizedName === 'id') {
        index = lowerHeader.indexOf('messageid');
      }
      return index;
    })
    .filter((index) => index >= 0);

  if (keyIndices.length === 0) {
    return { header, keySet, keyToRowMap, keyIndices };
  }

  // Read key columns data in chunks and track row positions
  const chunkSize = 1000;
  let startRow = 2; // Start from row 2 (after header)
  const minIndex = Math.min(...keyIndices);
  const maxIndex = Math.max(...keyIndices);
  const startCol = a1Col(minIndex + 1);
  const endCol = a1Col(maxIndex + 1);

  while (true) {
    const endRow = startRow + chunkSize - 1;
    const range = `${sheetTitle}!${startCol}${startRow}:${endCol}${endRow}`;

    try {
      const response = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range,
        majorDimension: 'ROWS',
      });

      const rows = (response.data?.values || []) as string[][];
      if (rows.length === 0) break;

      for (let localRowIndex = 0; localRowIndex < rows.length; localRowIndex++) {
        const rawRow = rows[localRowIndex];
        const globalRowNumber = startRow + localRowIndex;

        // Reconstruct full row for key generation
        const fullRow: Row = new Array(header.length).fill('');
        if (rawRow) {
          keyIndices.forEach((globalIndex, _localIndex) => {
            const localRowArrayIndex = globalIndex - minIndex;
            if (localRowArrayIndex >= 0 && localRowArrayIndex < rawRow.length) {
              fullRow[globalIndex] = rawRow[localRowArrayIndex];
            }
          });
        }

        const key = generateRowKey(fullRow, header, strategy);
        if (key.replace(/\\+/g, '') !== '') {
          keySet.add(key);
          keyToRowMap.set(key, globalRowNumber); // Store 1-indexed row number
        }
      }

      startRow += rows.length;
      if (rows.length < chunkSize) break; // Last chunk
    } catch {
      break; // End of data or error
    }
  }

  return { header, keySet, keyToRowMap, keyIndices };
}

/**
 * Partitions data into updates vs appends based on existing keys with row position tracking
 */
export function partitionDataForUpsert(rows: Row[], header: string[], keyStrategy: KeyGenerationStrategy, existingKeys: Set<string>, allowUpdates = false, keyToRowMap?: Map<string, number>): DataPartition {
  const toAppend: Row[] = [];
  const toUpdate: Array<{ row: Row; existingRowIndex: number }> = [];
  const skippedKeys: string[] = [];

  rows.forEach((row) => {
    const key = generateRowKey(row, header, keyStrategy);
    if (key.replace(/\\+/g, '') === '') {
      // Skip rows with empty keys
      return;
    }

    if (existingKeys.has(key)) {
      if (allowUpdates && keyToRowMap) {
        const existingRowIndex = keyToRowMap.get(key);
        if (existingRowIndex !== undefined) {
          toUpdate.push({ row, existingRowIndex });
        } else {
          skippedKeys.push(key);
        }
      } else {
        skippedKeys.push(key);
      }
    } else {
      toAppend.push(row);
    }
  });

  return {
    toAppend,
    toUpdate,
    skippedKeys,
  };
}

/**
 * Performs batch updates on existing rows using batchUpdate API
 */
export async function performBatchUpdates(sheets: sheets_v4.Sheets, spreadsheetId: string, sheetTitle: string, updates: Array<{ row: Row; existingRowIndex: number }>, header: string[], batchSize: number, valueInputOption: 'RAW' | 'USER_ENTERED'): Promise<{ updatedRows: number; errors?: string[] }> {
  const errors: string[] = [];
  let totalUpdated = 0;

  // Process updates in batches
  for (let i = 0; i < updates.length; i += batchSize) {
    const batch = updates.slice(i, i + batchSize);

    try {
      const requests: sheets_v4.Schema$Request[] = [];

      for (const { row, existingRowIndex } of batch) {
        requests.push({
          updateCells: {
            range: {
              sheetId: 0, // Will be resolved by sheet title in the range
              startRowIndex: existingRowIndex - 1, // Convert to 0-indexed
              endRowIndex: existingRowIndex, // Exclusive
              startColumnIndex: 0,
              endColumnIndex: header.length,
            },
            rows: [
              {
                values: row.map((cellValue) => ({
                  userEnteredValue: {
                    stringValue: String(cellValue ?? ''),
                  },
                })),
              },
            ],
            fields: 'userEnteredValue',
          },
        });
      }

      if (requests.length > 0) {
        await sheets.spreadsheets.batchUpdate({
          spreadsheetId,
          requestBody: {
            requests,
          },
        });

        totalUpdated += batch.length;
      }
    } catch (error) {
      const errorMsg = `Batch update ${Math.floor(i / batchSize) + 1} failed: ${error instanceof Error ? error.message : String(error)}`;
      errors.push(errorMsg);

      // Try individual updates for failed batches
      for (const { row, existingRowIndex } of batch) {
        try {
          const range = `${sheetTitle}!A${existingRowIndex}:${a1Col(header.length)}${existingRowIndex}`;
          await sheets.spreadsheets.values.update({
            spreadsheetId,
            range,
            valueInputOption,
            requestBody: {
              values: [row as (string | number | boolean | null | undefined)[]],
            },
          });
          totalUpdated += 1;
        } catch (singleError) {
          errors.push(`Failed to update row ${existingRowIndex}: ${singleError instanceof Error ? singleError.message : String(singleError)}`);
        }
      }
    }
  }

  const result: { updatedRows: number; errors?: string[] } = {
    updatedRows: totalUpdated,
  };

  if (errors.length > 0) {
    result.errors = errors;
  }

  return result;
}

export async function upsertByKey(
  sheets: sheets_v4.Sheets,
  {
    spreadsheetId,
    sheetTitle,
    sheetRef,
    rows,
    canonicalHeaders,
    options,
    logger,
  }: {
    spreadsheetId: string;
    sheetTitle?: string;
    sheetRef?: string;
    rows: Row[];
    canonicalHeaders?: string[];
    options: UpsertOptions;
    logger: Logger;
  }
): Promise<BatchOperationResult> {
  // Validate inputs
  if (!sheets) throw new Error('upsertByKey: sheets client is required');
  if (!spreadsheetId) throw new Error('upsertByKey: spreadsheetId is required');
  if (!sheetTitle && !sheetRef) throw new Error('upsertByKey: either sheetTitle or sheetRef is required');
  if (!Array.isArray(rows) || rows.length === 0) {
    return { updatedRows: 0, inserted: [], rowsSkipped: 0 };
  }

  // Step 1: Input validation (duplicate key checking)
  let resolvedSheetTitle = sheetTitle;
  if (sheetRef) {
    const sheet = await findSheetByRef(sheets, spreadsheetId, sheetRef, logger);
    resolvedSheetTitle = sheet?.properties?.title || sheetRef;
  }

  if (!resolvedSheetTitle) {
    throw new Error('upsertByKey: could not resolve sheet title');
  }

  // Validate for duplicate keys in input data
  const inputValidation = validateRowKeys(rows, canonicalHeaders || [], options.keyStrategy);
  if (!inputValidation.valid) {
    throw new Error(`upsertByKey: duplicate keys found in input data: ${inputValidation.duplicateKeys.join(', ')}`);
  }

  // Step 2: Header discovery and column mapping
  const currentHeader = await discoverHeader(sheets, spreadsheetId, resolvedSheetTitle);
  let effectiveHeader = currentHeader;

  // If sheet is empty and we have canonical headers, use them
  if (currentHeader.length === 0 && canonicalHeaders && canonicalHeaders.length > 0) {
    effectiveHeader = canonicalHeaders;
    // Write headers to sheet
    await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: `${resolvedSheetTitle}!A1`,
      valueInputOption: options.valueInputOption || 'USER_ENTERED',
      insertDataOption: 'INSERT_ROWS',
      requestBody: { values: [effectiveHeader] },
    });
  }

  // Validate headers if we have requirements
  if (canonicalHeaders && canonicalHeaders.length > 0) {
    const headerValidation = validateAndMapHeaders(effectiveHeader, canonicalHeaders);
    if (!headerValidation.valid) {
      throw new Error(`upsertByKey: missing required columns: ${headerValidation.missingColumns.join(', ')}`);
    }
  }

  // Step 3: Key column reading for existing row lookup with position tracking
  const { keySet: existingKeys, keyToRowMap } = await snapshotHeaderKeysAndPositions(sheets, spreadsheetId, resolvedSheetTitle, options.keyStrategy.keyColumns, options.keyStrategy);

  // Step 4: Partition changes (updates vs appends)
  let processedRows = rows;

  // Map canonical data to sheet structure if needed
  if (canonicalHeaders && canonicalHeaders.length > 0 && effectiveHeader.length > 0) {
    processedRows = mapRowsToHeader({
      rows,
      header: effectiveHeader,
      canonical: canonicalHeaders,
    });
  }

  const partition = partitionDataForUpsert(processedRows, effectiveHeader, options.keyStrategy, existingKeys, options.allowUpdates, keyToRowMap);

  // Step 5: Efficient batch writing
  const batchSize = options.batchSize || 50;
  const errors: string[] = [];
  let totalUpdated = 0;
  const insertedKeys: string[] = [];

  // Handle appends with proper error handling and partial success tracking
  if (partition.toAppend.length > 0) {
    for (let i = 0; i < partition.toAppend.length; i += batchSize) {
      const batch = partition.toAppend.slice(i, i + batchSize);
      try {
        const response = await sheets.spreadsheets.values.append({
          spreadsheetId,
          range: `${resolvedSheetTitle}!A1`,
          valueInputOption: options.valueInputOption || 'USER_ENTERED',
          insertDataOption: 'INSERT_ROWS',
          requestBody: { values: batch },
        });

        const updatedRows = Number(response.data?.updates?.updatedRows || batch.length);
        totalUpdated += updatedRows;

        // Track inserted keys for successful batch
        batch.forEach((row) => {
          const key = generateRowKey(row, effectiveHeader, options.keyStrategy);
          if (key.replace(/\\+/g, '') !== '') {
            insertedKeys.push(key);
          }
        });
      } catch (error) {
        const errorMsg = `Batch ${Math.floor(i / batchSize) + 1} append failed: ${error instanceof Error ? error.message : String(error)}`;
        errors.push(errorMsg);
      }
    }
  }

  // Handle updates with proper row position tracking
  if (partition.toUpdate.length > 0 && options.allowUpdates) {
    try {
      const updateResult = await performBatchUpdates(sheets, spreadsheetId, resolvedSheetTitle, partition.toUpdate, effectiveHeader, batchSize, options.valueInputOption || 'USER_ENTERED');
      totalUpdated += updateResult.updatedRows;
      if (updateResult.errors) {
        errors.push(...updateResult.errors);
      }
    } catch (updateError) {
      const errorMsg = `Update operations failed: ${updateError instanceof Error ? updateError.message : String(updateError)}`;
      errors.push(errorMsg);
    }
  }

  const result: BatchOperationResult = {
    updatedRows: totalUpdated,
    inserted: insertedKeys,
    rowsSkipped: partition.skippedKeys.length,
  };

  if (errors.length > 0) {
    result.errors = errors;
  }

  return result;
}
