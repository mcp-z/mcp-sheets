import type { drive_v3, sheets_v4 } from 'googleapis';
import type { GoogleApiError } from '../types.ts';

export const SPREADSHEET_URL_RE = /https:\/\/docs.google.com\/spreadsheets\/d\/([a-zA-Z0-9-_]{10,})/;
export const SPREADSHEET_ID_RE = /^[a-zA-Z0-9-_]{30,}$/;

export async function findById(sheets: sheets_v4.Sheets, drive: drive_v3.Drive, id: string) {
  const ssResp = await sheets.spreadsheets.get({ spreadsheetId: id }).catch((e: unknown) => {
    const error = e as GoogleApiError;
    const status = error?.response?.status || error?.status || error?.code;
    if (Number(status) === 404) return null;
    throw e;
  });
  const dmResp = await drive.files.get({ fileId: id, fields: 'modifiedTime,webViewLink,name' }).catch((e: unknown) => {
    const error = e as GoogleApiError;
    const status = error?.response?.status || error?.status || error?.code;
    if (Number(status) === 404) return null;
    throw e;
  });
  const dm = dmResp?.data;
  if (!ssResp) return null;
  const ss = ssResp.data;
  return {
    id: id,
    spreadsheetTitle: ss?.properties?.title ?? dm?.name,
    url: ss?.spreadsheetUrl ?? dm?.webViewLink,
    modifiedTime: dm?.modifiedTime,
  };
}

export async function findByName(drive: drive_v3.Drive, name: string) {
  const escaped = String(name).replace(/['"\\]/g, (m) => `\\${m}`);
  const q = `mimeType='application/vnd.google-apps.spreadsheet' and name contains '${escaped}' and trashed = false`;
  const resp = await drive.files.list({ q, pageSize: 50, fields: 'files(id,name,webViewLink,modifiedTime)' });
  const files = resp.data?.files || [];
  return files.map((f) => ({ id: String(f.id), spreadsheetTitle: String(f.name), url: String(f.webViewLink), modifiedTime: f.modifiedTime }));
}

export async function findSpreadsheetsByRef(sheets: sheets_v4.Sheets, drive: drive_v3.Drive, ref: string) {
  const trimmedRef = ref?.trim();
  if (!trimmedRef) return [];

  // Strategy 1: Try URL extraction (most specific)
  const urlMatch = trimmedRef.match(SPREADSHEET_URL_RE);
  if (urlMatch && typeof urlMatch[1] === 'string' && urlMatch[1].length > 0) {
    const id = urlMatch[1];
    const r = await findById(sheets, drive, id);
    return r ? [r] : [];
  }

  // Strategy 2: Try direct ID match (if it looks like a spreadsheet ID)
  if (SPREADSHEET_ID_RE.test(trimmedRef)) {
    const r = await findById(sheets, drive, trimmedRef);
    return r ? [r] : [];
  }

  // Strategy 3: Search by name
  return await findByName(drive, trimmedRef);
}
