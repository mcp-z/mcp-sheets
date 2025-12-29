import type { EnrichedExtra } from '@mcp-z/oauth-google';
import type { ResourceConfig, ResourceModule } from '@mcp-z/server';
import { ResourceTemplate } from '@modelcontextprotocol/sdk/server/mcp.js';
import type { RequestHandlerExtra } from '@modelcontextprotocol/sdk/shared/protocol.js';
import type { ReadResourceResult, ServerNotification, ServerRequest } from '@modelcontextprotocol/sdk/types.js';
import { google, type sheets_v4 } from 'googleapis';
import type { GoogleApiError } from '../../types.js';

export default function createResource(): ResourceModule {
  const template = new ResourceTemplate('sheets://spreadsheets/{spreadsheetId}', { list: undefined });
  const config: ResourceConfig = {
    description: 'Sheets spreadsheet resource',
    mimeType: 'application/json',
  };

  const handler = async (uri: URL, variables: Record<string, string | string[]>, extra: RequestHandlerExtra<ServerRequest, ServerNotification>): Promise<ReadResourceResult> => {
    const spreadsheetId = Array.isArray(variables.spreadsheetId) ? variables.spreadsheetId[0] : variables.spreadsheetId;

    if (!spreadsheetId) {
      return {
        contents: [
          {
            uri: uri.href,
            mimeType: 'application/json',
            text: JSON.stringify({ error: 'spreadsheetId is required' }),
          },
        ],
      };
    }

    try {
      // Safe type guard to access middleware-enriched extra
      const { logger, authContext } = extra as unknown as EnrichedExtra;
      const sheets = google.sheets({ version: 'v4', auth: authContext.auth });
      const resp = await sheets.spreadsheets.get({
        spreadsheetId,
        fields: 'spreadsheetId,properties.title,sheets.properties',
      });
      const data = resp.data as sheets_v4.Schema$Spreadsheet;
      logger.info('sheets-spreadsheet resource fetch success', {
        spreadsheetId: data.spreadsheetId,
        title: data?.properties?.title,
        sheetCount: (data?.sheets || []).length,
      });
      return {
        contents: [
          {
            uri: uri.href,
            mimeType: 'application/json',
            text: JSON.stringify({
              id: data?.spreadsheetId,
              title: data?.properties?.title,
              sheets: (data?.sheets || []).map((s) => ({
                id: s?.properties?.sheetId,
                title: s?.properties?.title,
                rowCount: s?.properties?.gridProperties?.rowCount,
                columnCount: s?.properties?.gridProperties?.columnCount,
              })),
            }),
          },
        ],
      };
    } catch (error) {
      const { logger } = extra as unknown as EnrichedExtra;
      logger.info('sheets-spreadsheet resource fetch failed', error as Record<string, unknown>);
      return {
        contents: [
          {
            uri: uri.href,
            mimeType: 'application/json',
            text: JSON.stringify({
              error: String((error as GoogleApiError)?.message ?? error),
            }),
          },
        ],
      };
    }
  };

  return {
    name: 'spreadsheet',
    template,
    config,
    handler,
  };
}
