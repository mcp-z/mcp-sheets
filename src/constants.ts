/**
 * Sheets MCP Server Constants
 *
 * These scopes are required for Google Sheets and Drive functionality and are hardcoded
 * rather than externally configured since this server knows its own requirements.
 */

// Google OAuth scopes required for Sheets and Drive operations
export const GOOGLE_SCOPE = 'openid https://www.googleapis.com/auth/userinfo.profile https://www.googleapis.com/auth/userinfo.email https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive';

// Native paste-special modes of the Sheets API CopyPasteRequest (PasteType / PasteOrientation)
export const PASTE_TYPES = ['PASTE_NORMAL', 'PASTE_VALUES', 'PASTE_FORMAT', 'PASTE_NO_BORDERS', 'PASTE_FORMULA', 'PASTE_DATA_VALIDATION', 'PASTE_CONDITIONAL_FORMATTING'] as const;
export const PASTE_ORIENTATIONS = ['NORMAL', 'TRANSPOSE'] as const;

// Native-date guidance shared by the write, format and read tool descriptions; the example arguments are checked against the tool schemas in test/unit/constants.test.ts.
// Docs: https://developers.google.com/workspace/sheets/api/reference/rest/v4/DateTimeRenderOption and .../ValueInputOption
export const NATIVE_DATE_POINTER = 'Dates: a date-looking string is not necessarily a native date. See values-batch-update for the native-date write, format and read-back steps.';

export const RENDER_DESCRIPTION =
  'How to render cell values. FORMATTED_VALUE (default): calculated with formatting. UNFORMATTED_VALUE: calculated without formatting; native dates come back as numeric serials (days since 1899-12-30), text dates as strings, so use it to confirm a date is numeric. FORMULA: show formula text instead of result.';

// Example call arguments (id/gid/includeData omitted), step 1 to 3 of the native-date write, format and read-back.
export const NATIVE_DATE_EXAMPLE_WRITE = { requests: [{ range: 'B2', majorDimension: 'ROWS', values: [[46334.5]] }], valueInputOption: 'RAW' } as const;
export const NATIVE_DATE_EXAMPLE_FORMAT = { requests: [{ range: 'B2', numberFormat: { type: 'DATE', pattern: 'yyyy-mm-dd hh:mm' } }] } as const;
export const NATIVE_DATE_EXAMPLE_READ = { range: 'B2', render: 'UNFORMATTED_VALUE' } as const;

export const NATIVE_DATE_GUIDANCE = [
  "NATIVE DATES: a native Sheets date/time is a serial number plus a date number format. The integer counts days from 1899-12-30; the fraction is the time of day (46334.5 = 2026-11-08 12:00, with a matching format). A serial carries no timezone of its own and is read in the spreadsheet's timezone.",
  'RAW stores values as supplied, so a date-looking string stays text; a numeric serial plus a date format is an explicit native-date write. USER_ENTERED may parse date strings the way the Sheets UI does, using the workbook locale.',
  'Keep the intended wall time/timezone and recorded precision: a date-only record is not an exact midnight event time, and existing timestamps are not reinterpreted.',
  `Example: (1) values-batch-update ${JSON.stringify(NATIVE_DATE_EXAMPLE_WRITE)}; (2) cells-format ${JSON.stringify(NATIVE_DATE_EXAMPLE_FORMAT)}; (3) rows-get (or columns-get / values-search) ${JSON.stringify(NATIVE_DATE_EXAMPLE_READ)} returns 46334.5 (numeric, so it is not text), while render FORMATTED_VALUE returns "2026-11-08 12:00"; then check dependent chronological formulas.`,
  'Docs: https://developers.google.com/workspace/sheets/api/reference/rest/v4/DateTimeRenderOption and https://developers.google.com/workspace/sheets/api/reference/rest/v4/ValueInputOption',
].join(' ');
