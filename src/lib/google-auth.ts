import { attachTokenProvider, type GoogleAuthProvider } from '@mcp-z/oauth-google';
import { google } from 'googleapis';

/**
 * A Google API client bound to this request's token provider.
 *
 * `@mcp-z/oauth-google` mints tokens and depends on no Google SDK, so the client
 * is built here, from this package's own `googleapis`. That keeps one copy of
 * `google-auth-library` in the tree by construction rather than by matching pins.
 */
export function googleAuth(auth: GoogleAuthProvider) {
  return attachTokenProvider(new google.auth.OAuth2(), auth);
}
