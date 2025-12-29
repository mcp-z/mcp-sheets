/**
 * Sheets MCP Server Constants
 *
 * These scopes are required for Google Sheets and Drive functionality and are hardcoded
 * rather than externally configured since this server knows its own requirements.
 */

// Google OAuth scopes required for Sheets and Drive operations
export const GOOGLE_SCOPE = 'openid https://www.googleapis.com/auth/userinfo.profile https://www.googleapis.com/auth/userinfo.email https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive';
