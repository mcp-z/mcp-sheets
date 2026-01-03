import type { sheets_v4 } from 'googleapis';
import type { Logger } from '../types.ts';
import { a1Col } from './column-utilities.ts';

type ValueRange = sheets_v4.Schema$ValueRange;

export async function ensureSheetIfNeeded(sheets: sheets_v4.Sheets, id: string, sheetTitle: string, ensureSheet: boolean, headers: string[] | undefined, logger: Logger) {
  if (!sheetTitle) return { sheetCreated: false };
  const ssResp = await sheets.spreadsheets.get({ spreadsheetId: id });
  const ss = ssResp.data;
  const has = Array.isArray(ss?.sheets) && ss.sheets.some((s) => s?.properties?.title === sheetTitle);
  if (has) {
    const sid = ss.sheets?.find((s) => s?.properties?.title === sheetTitle)?.properties?.sheetId;
    const result: { sheetCreated: boolean; sheetGUID?: string; headersEnsured?: boolean } = { sheetCreated: false };
    if (sid != null) result.sheetGUID = String(sid);
    if (headers) {
      try {
        const { header: merged } = await ensureTabAndHeaders(sheets, { spreadsheetId: id, sheetTitle, requiredHeader: headers, logger });
        if (merged) result.headersEnsured = true;
      } catch (e) {
        logger.warn?.('ensureTabAndHeaders after sheet creation failed', e as Record<string, unknown>);
      }
    }
    return result;
  }
  if (!ensureSheet) return { sheetCreated: false };

  const res = await sheets.spreadsheets.batchUpdate({ spreadsheetId: id, requestBody: { requests: [{ addSheet: { properties: { title: sheetTitle } } }] } });
  const createdId = res.data?.replies?.[0]?.addSheet?.properties?.sheetId;
  const result: { sheetCreated: boolean; sheetGUID?: string; headersEnsured?: boolean } = { sheetCreated: !!createdId };
  if (createdId != null) result.sheetGUID = String(createdId);
  if (createdId && headers) {
    try {
      const { header: merged } = await ensureTabAndHeaders(sheets, { spreadsheetId: id, sheetTitle, requiredHeader: headers, logger });
      if (merged) result.headersEnsured = true;
    } catch (e) {
      logger.warn?.('ensureTabAndHeaders after sheet creation failed', e as Record<string, unknown>);
    }
  }
  return result;
}

export async function findSheetByRef(client: sheets_v4.Sheets, spreadsheetId: string, sheetRef: string, logger: Logger) {
  const ssResp = await client.spreadsheets.get({ spreadsheetId });
  const sheets = ssResp.data?.sheets || [];
  const trimmedRef = sheetRef?.trim();

  logger.debug?.('findSheetByRef called', {
    spreadsheetId,
    sheetRef,
    trimmedRef,
    availableSheets: sheets.map((s) => ({
      title: s?.properties?.title,
      id: s?.properties?.sheetId,
    })),
  });

  if (!trimmedRef) {
    logger.warn?.('findSheetByRef empty sheetRef');
    return null;
  }

  // Strategy 1: Try exact title match first (most common case)
  const exactTitleMatch = sheets.find((s) => s?.properties?.title === trimmedRef);
  if (exactTitleMatch) {
    logger.debug?.('findSheetByRef exact title match found');
    return exactTitleMatch;
  }

  // Strategy 2: Try exact GUID match (for when user passes actual sheet ID)
  const exactGuidMatch = sheets.find((s) => String(s?.properties?.sheetId) === trimmedRef);
  if (exactGuidMatch) {
    logger.debug?.('findSheetByRef exact GUID match found');
    return exactGuidMatch;
  }

  // Strategy 3: Case-insensitive title match
  const caseInsensitiveMatch = sheets.find((s) => s?.properties?.title?.toLowerCase() === trimmedRef.toLowerCase());
  if (caseInsensitiveMatch) {
    logger.debug?.('findSheetByRef case-insensitive title match found');
    return caseInsensitiveMatch;
  }

  // Strategy 4: Trimmed title match (handles extra whitespace in sheet names)
  const trimmedTitleMatch = sheets.find((s) => s?.properties?.title?.trim() === trimmedRef);
  if (trimmedTitleMatch) {
    logger.debug?.('findSheetByRef trimmed title match found');
    return trimmedTitleMatch;
  }

  // Strategy 5: Case-insensitive + trimmed title match
  const caseInsensitiveTrimmedMatch = sheets.find((s) => s?.properties?.title?.trim().toLowerCase() === trimmedRef.toLowerCase());
  if (caseInsensitiveTrimmedMatch) {
    logger.debug?.('findSheetByRef case-insensitive trimmed title match found');
    return caseInsensitiveTrimmedMatch;
  }

  // Strategy 6: Partial title match (last resort)
  const partialMatch = sheets.find((s) => s?.properties?.title?.toLowerCase().includes(trimmedRef.toLowerCase()));
  if (partialMatch) {
    logger.debug?.('findSheetByRef partial title match found');
    return partialMatch;
  }

  logger.warn?.('findSheetByRef no match found', {
    searchRef: trimmedRef,
    availableTitles: sheets.map((s) => s?.properties?.title),
    availableIds: sheets.map((s) => s?.properties?.sheetId),
  });
  return null;
}

// Overloaded function signatures
export async function ensureTabAndHeaders(sheets: sheets_v4.Sheets, params: { spreadsheetId: string; sheetTitle: string; requiredHeader?: string[] | null; keyColumns?: string[]; logger: Logger }): Promise<{ header: string[]; keySet: Set<string>; keyColumnsIdx?: number[] }>;

export async function ensureTabAndHeaders(sheets: sheets_v4.Sheets, params: { spreadsheetId: string; sheetRef: string; requiredHeader?: string[] | null; keyColumns?: string[]; logger: Logger }): Promise<{ header: string[]; keySet: Set<string>; keyColumnsIdx?: number[] }>;

export async function ensureTabAndHeaders(
  sheets: sheets_v4.Sheets,
  {
    spreadsheetId,
    sheetTitle,
    sheetRef,
    requiredHeader = null as string[] | null,
    keyColumns = ['id'] as string[],
    logger,
  }: {
    spreadsheetId: string;
    sheetTitle?: string;
    sheetRef?: string;
    requiredHeader?: string[] | null;
    keyColumns?: string[];
    logger: Logger;
  }
) {
  if (!sheets) throw new Error('ensureTabAndHeaders: sheets is required');
  if (!spreadsheetId) throw new Error('ensureTabAndHeaders: spreadsheetId is required');
  if (!sheetTitle && !sheetRef) throw new Error('ensureTabAndHeaders: either sheetTitle or sheetRef is required');

  // Resolve the actual sheet title from sheetRef if provided
  let resolvedSheetTitle = sheetTitle;
  if (sheetRef) {
    const sheet = await findSheetByRef(sheets, spreadsheetId, sheetRef, logger);
    if (sheet) {
      resolvedSheetTitle = sheet.properties?.title || sheetRef;
    } else {
      // Sheet doesn't exist, use the sheetRef as the title for creation
      resolvedSheetTitle = sheetRef;
    }
  }

  if (!resolvedSheetTitle) throw new Error('ensureTabAndHeaders: could not resolve sheet title');

  const defaultHeader = ['id', 'provider', 'threadId', 'to', 'from', 'cc', 'bcc', 'date', 'subject', 'labels', 'snippet', 'body'];
  const header = Array.isArray(requiredHeader) && requiredHeader.length ? requiredHeader : defaultHeader;

  let existingHeader: string[] | null = null;
  try {
    const resp = await sheets.spreadsheets.values.get({ spreadsheetId, range: `${resolvedSheetTitle}!1:1`, majorDimension: 'ROWS' });
    const vr = resp.data as ValueRange;
    const values = ((vr?.values || [])[0] || []) as string[];
    existingHeader = values.length ? values : null;
  } catch (_err: unknown) {
    try {
      await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [{ addSheet: { properties: { title: resolvedSheetTitle } } }] } });
    } catch {}
    await sheets.spreadsheets.values.append({ spreadsheetId, range: `${resolvedSheetTitle}!A1`, valueInputOption: 'RAW', insertDataOption: 'INSERT_ROWS', requestBody: { values: [header] } });
    return { header, keySet: new Set<string>() };
  }

  if (!existingHeader) {
    await sheets.spreadsheets.values.append({ spreadsheetId, range: `${resolvedSheetTitle}!A1`, valueInputOption: 'RAW', insertDataOption: 'INSERT_ROWS', requestBody: { values: [header] } });
    return { header, keySet: new Set<string>() };
  }

  const mergedHeader = existingHeader.slice();
  for (const col of header) if (!mergedHeader.includes(col)) mergedHeader.push(col);

  if (Array.isArray(keyColumns)) {
    for (const name of keyColumns) {
      if (typeof name !== 'string' || name.length === 0) throw new Error('ensureTabAndHeaders: keyColumns must be an array of non-empty strings');
      if (!mergedHeader.includes(name)) mergedHeader.push(name);
    }
  }

  if (mergedHeader.length !== existingHeader.length) {
    await sheets.spreadsheets.values.update({ spreadsheetId, range: `${resolvedSheetTitle}!1:1`, valueInputOption: 'RAW', requestBody: { values: [mergedHeader] } });
  }

  const keySet = new Set<string>();
  const chunkSize = 1000;
  let startRow = 2;

  const normalizeKeyColumnsByName = (kc: string[], hdr: string[]) => {
    if (!Array.isArray(kc) || kc.length === 0) return [] as number[];
    return kc.map((name) => {
      if (typeof name !== 'string' || name.length === 0) return -1;
      const idx = hdr.indexOf(name);
      return Number.isInteger(idx) && idx >= 0 ? idx : -1;
    });
  };

  const keyColsIdx = normalizeKeyColumnsByName(keyColumns, mergedHeader);
  const validKeyIdxs = keyColsIdx.filter((i) => i >= 0);
  if (validKeyIdxs.length === 0) return { header: mergedHeader, keySet: new Set<string>(), keyColumnsIdx: keyColsIdx };

  const minIdx = Math.min(...validKeyIdxs);
  const maxIdx = Math.max(...validKeyIdxs);
  const startCol = a1Col(minIdx + 1);
  const endCol = a1Col(maxIdx + 1);

  while (true) {
    const endRow = startRow + chunkSize - 1;
    const range = `${resolvedSheetTitle}!${startCol}${startRow}:${endCol}${endRow}`;
    const respChunk = await sheets.spreadsheets.values.get({ spreadsheetId, range, majorDimension: 'ROWS' });
    const vrChunk = respChunk.data as ValueRange;
    const chunkRows = (vrChunk?.values || []) as string[][];
    if (!chunkRows || chunkRows.length === 0) break;
    for (const r of chunkRows) {
      const row = Array.isArray(r) ? r : [];

      // Use the same key generation logic as appendRows for consistency
      const lower = mergedHeader.map((h) => String(h ?? '').toLowerCase());
      const provHeaderIdx = lower.indexOf('provider');
      let idHeaderIdx = lower.indexOf('messageid');
      if (idHeaderIdx === -1) idHeaderIdx = lower.indexOf('id');

      let compositeKey: string;

      if (provHeaderIdx >= 0 && idHeaderIdx >= 0) {
        // Special case: if both provider and id columns exist, use them
        const providerOffset = provHeaderIdx - minIdx;
        const idOffset = idHeaderIdx - minIdx;
        const providerVal = providerOffset >= 0 ? String(row[providerOffset] ?? '').trim() : '';
        const messageVal = idOffset >= 0 ? String(row[idOffset] ?? '').trim() : '';
        compositeKey = [providerVal, messageVal].join('\\');
      } else {
        // General case: use the specified key columns
        const comps = keyColsIdx.map((i) => (i < 0 ? '' : String(row[i - minIdx] ?? '').trim()));
        compositeKey = comps.join('\\');
      }

      if (compositeKey.replace(/\\+/g, '') !== '') keySet.add(compositeKey);
    }
    startRow += chunkRows.length;
  }
  return { header: mergedHeader, keySet, keyColumnsIdx: keyColsIdx };
}
