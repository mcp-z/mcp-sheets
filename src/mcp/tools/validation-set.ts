import type { EnrichedExtra } from '@mcp-z/oauth-google';
import { schemas } from '@mcp-z/oauth-google';

const { AuthRequiredBranchSchema } = schemas;

import type { ToolModule } from '@mcp-z/server';
import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js';
import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { google, type sheets_v4 } from 'googleapis';
import { z } from 'zod';
import { SheetGidOutput, SheetGidSchema, SpreadsheetIdOutput, SpreadsheetIdSchema } from '../../schemas/index.ts';
import { parseA1Notation, rangeReferenceToGridRange } from '../../spreadsheet/range-operations.ts';

// Discriminated union for validation rule types
const ValidationRuleSchema = z.discriminatedUnion('conditionType', [
  // Dropdown with hardcoded list
  z.object({
    conditionType: z.literal('ONE_OF_LIST'),
    values: z.array(z.string()).min(1).describe('Hardcoded dropdown options'),
    showDropdown: z.boolean().default(true).describe('Show dropdown UI in cells'),
    strict: z.boolean().default(true).describe('Reject invalid entries'),
  }),
  // Dropdown from range
  z.object({
    conditionType: z.literal('ONE_OF_RANGE'),
    sourceRange: z.string().min(1).describe('A1 notation range containing dropdown options (e.g., "Options!A1:A10")'),
    showDropdown: z.boolean().default(true).describe('Show dropdown UI in cells'),
    strict: z.boolean().default(true).describe('Reject invalid entries'),
  }),
  // Numeric constraints
  z.object({
    conditionType: z.enum(['NUMBER_GREATER', 'NUMBER_LESS', 'NUMBER_BETWEEN']),
    values: z.array(z.number()).min(1).max(2).describe('1 value for GREATER/LESS, 2 values for BETWEEN'),
    strict: z.boolean().default(true).describe('Reject invalid entries'),
  }),
  // Text constraints
  z.object({
    conditionType: z.enum(['TEXT_CONTAINS', 'TEXT_IS_EMAIL', 'TEXT_IS_URL']),
    values: z.array(z.string()).optional().describe('Text to match (for TEXT_CONTAINS only)'),
    strict: z.boolean().default(true).describe('Reject invalid entries'),
  }),
  // Date constraints
  z.object({
    conditionType: z.enum(['DATE_AFTER', 'DATE_BEFORE', 'DATE_BETWEEN']),
    values: z.array(z.string()).min(1).max(2).describe('ISO date strings (YYYY-MM-DD)'),
    strict: z.boolean().default(true).describe('Reject invalid entries'),
  }),
  // Custom formula
  z.object({
    conditionType: z.literal('CUSTOM_FORMULA'),
    formula: z.string().min(1).describe('Custom validation formula (e.g., "=A1>0")'),
    strict: z.boolean().default(true).describe('Reject invalid entries'),
  }),
]);

// Input schema for validation requests
const ValidationRequestSchema = z.object({
  range: z.string().min(1).describe('A1 notation range for validation (e.g., "B2:B100")'),
  rule: ValidationRuleSchema,
  inputMessage: z.string().optional().describe('Help text shown when cell is selected'),
});

const inputSchema = z.object({
  id: SpreadsheetIdSchema,
  gid: SheetGidSchema,
  requests: z.array(ValidationRequestSchema).min(1).max(50).describe('Array of validation rules. Batch multiple ranges for efficiency.'),
});

const successBranchSchema = z.object({
  type: z.literal('success'),
  id: SpreadsheetIdOutput,
  gid: SheetGidOutput,
  sheetTitle: z.string().describe('Title of the sheet with validation'),
  sheetUrl: z.string().describe('URL of the sheet with validation'),
  successCount: z.number().int().nonnegative().describe('Number of validation rules successfully applied'),
  failedRanges: z
    .array(
      z.object({
        range: z.string().describe('A1 notation of range that failed validation setup'),
        error: z.string().describe('Why validation failed for this range'),
      })
    )
    .optional()
    .describe('Only populated if some validation rules failed'),
});

const outputSchema = z.discriminatedUnion('type', [successBranchSchema, AuthRequiredBranchSchema]);

const config = {
  description: 'Add data validation rules including dropdowns, numeric constraints, text patterns, date ranges, custom formulas. Supports batch operations for efficiency. Use discriminated conditionType to specify rule type. Best for enforcing data integrity and providing user-friendly input controls.',
  inputSchema,
  outputSchema: z.object({
    result: outputSchema,
  }),
} as const;

export type Input = z.infer<typeof inputSchema>;
export type Output = z.infer<typeof outputSchema>;

// Map condition types to Google Sheets API format
function buildCondition(rule: z.infer<typeof ValidationRuleSchema>) {
  switch (rule.conditionType) {
    case 'ONE_OF_LIST':
      return {
        type: 'ONE_OF_LIST',
        values: rule.values.map((value) => ({ userEnteredValue: value })),
      };
    case 'ONE_OF_RANGE':
      return {
        type: 'ONE_OF_RANGE',
        values: [{ userEnteredValue: `=${rule.sourceRange}` }],
      };
    case 'NUMBER_GREATER':
      return {
        type: 'NUMBER_GREATER',
        values: [{ userEnteredValue: String(rule.values[0]) }],
      };
    case 'NUMBER_LESS':
      return {
        type: 'NUMBER_LESS',
        values: [{ userEnteredValue: String(rule.values[0]) }],
      };
    case 'NUMBER_BETWEEN':
      return {
        type: 'NUMBER_BETWEEN',
        values: [{ userEnteredValue: String(rule.values[0]) }, { userEnteredValue: String(rule.values[1]) }],
      };
    case 'TEXT_CONTAINS':
      return {
        type: 'TEXT_CONTAINS',
        values: rule.values ? [{ userEnteredValue: rule.values[0] }] : [],
      };
    case 'TEXT_IS_EMAIL':
      return {
        type: 'TEXT_IS_EMAIL',
        values: [],
      };
    case 'TEXT_IS_URL':
      return {
        type: 'TEXT_IS_URL',
        values: [],
      };
    case 'DATE_AFTER':
      return {
        type: 'DATE_AFTER',
        values: [{ userEnteredValue: rule.values[0] }],
      };
    case 'DATE_BEFORE':
      return {
        type: 'DATE_BEFORE',
        values: [{ userEnteredValue: rule.values[0] }],
      };
    case 'DATE_BETWEEN':
      return {
        type: 'DATE_BETWEEN',
        values: [{ userEnteredValue: rule.values[0] }, { userEnteredValue: rule.values[1] }],
      };
    case 'CUSTOM_FORMULA':
      return {
        type: 'CUSTOM_FORMULA',
        values: [{ userEnteredValue: rule.formula.startsWith('=') ? rule.formula : `=${rule.formula}` }],
      };
    default:
      // Exhaustive check - all valid conditionType values should be handled above
      throw new Error(`Unknown condition type: ${(rule as { conditionType: string }).conditionType}`);
  }
}

async function handler({ id, gid, requests }: Input, extra: EnrichedExtra): Promise<CallToolResult> {
  const logger = extra.logger;
  logger.info('sheets.validation.set called', {
    id,
    gid,
    requestCount: requests.length,
  });

  try {
    const sheets = google.sheets({ version: 'v4', auth: extra.authContext.auth });

    // Get spreadsheet and sheet info in single API call
    const spreadsheetResponse = await sheets.spreadsheets.get({
      spreadsheetId: id,
      fields: 'sheets.properties.sheetId,sheets.properties.title',
    });

    // Find sheet by gid
    const sheet = spreadsheetResponse.data.sheets?.find((s) => String(s.properties?.sheetId) === gid);
    if (!sheet?.properties) {
      logger.info('Sheet not found for validation set', { id, gid, requestCount: requests.length });
      throw new McpError(ErrorCode.InvalidParams, `Sheet not found: ${gid}`);
    }

    const sheetTitle = sheet.properties.title ?? gid;
    const sheetId = sheet.properties.sheetId;
    const sheetUrl = `https://docs.google.com/spreadsheets/d/${id}/edit#gid=${sheetId}`;

    // Validate sheetId exists
    if (sheetId === undefined || sheetId === null) {
      logger.error('Sheet ID not available for validation set', { id, gid, sheetTitle });
      throw new McpError(ErrorCode.InternalError, `Sheet ID not available for ${gid}. Cannot perform validation operations without valid sheet ID.`);
    }

    // Build batch update requests
    const batchRequests: sheets_v4.Schema$Request[] = [];
    const failedRanges: Array<{ range: string; error: string }> = [];

    for (const request of requests) {
      try {
        // Parse A1 notation to range reference
        const rangeRef = parseA1Notation(request.range);

        // Build grid range from range reference using helper function
        const gridRange = rangeReferenceToGridRange(rangeRef, sheetId);

        // Build validation rule
        const condition = buildCondition(request.rule);

        // Build data validation rule - structure matches Schema$DataValidationRule
        const dataValidationRule = {
          condition,
          strict: request.rule.strict,
          showCustomUi: 'showDropdown' in request.rule ? request.rule.showDropdown : undefined,
          inputMessage: request.inputMessage,
        } as sheets_v4.Schema$DataValidationRule;

        // Add setDataValidation request
        batchRequests.push({
          setDataValidation: {
            range: gridRange,
            rule: dataValidationRule,
          },
        });
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        logger.info('Failed to build validation rule', {
          range: request.range,
          error: message,
        });
        failedRanges.push({
          range: request.range,
          error: `Failed to build validation rule: ${message}`,
        });
      }
    }

    // Early return if all ranges failed
    if (batchRequests.length === 0) {
      const result: Output = {
        type: 'success' as const,
        id,
        gid: String(sheetId),
        sheetTitle,
        sheetUrl,
        successCount: 0,
        failedRanges: failedRanges.length > 0 ? failedRanges : undefined,
      };

      return {
        content: [{ type: 'text' as const, text: JSON.stringify(result) }],
        structuredContent: { result },
      };
    }

    logger.info('sheets.validation.set executing batch request', {
      spreadsheetId: id,
      sheetTitle,
      batchRequestsCount: batchRequests.length,
    });

    // Execute the batch update
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: id,
      requestBody: {
        requests: batchRequests,
      },
    });

    logger.info('sheets.validation.set completed successfully', {
      successCount: requests.length - failedRanges.length,
      failedCount: failedRanges.length,
    });

    const result: Output = {
      type: 'success' as const,
      id,
      gid: String(sheetId),
      sheetTitle,
      sheetUrl,
      successCount: requests.length - failedRanges.length,
      failedRanges: failedRanges.length > 0 ? failedRanges : undefined,
    };

    return {
      content: [{ type: 'text' as const, text: JSON.stringify(result) }],
      structuredContent: { result },
    };
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    logger.error('Validation set operation failed', {
      id,
      gid,
      requestCount: requests.length,
      error: message,
    });

    throw new McpError(ErrorCode.InternalError, `Error setting validation: ${message}`, {
      stack: error instanceof Error ? error.stack : undefined,
    });
  }
}

export default function createTool() {
  return {
    name: 'validation-set',
    config,
    handler,
  } satisfies ToolModule;
}
