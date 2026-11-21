import type { GoogleApiError } from '../types.ts';

/** True when Google refused the request itself (HTTP 400: invalid argument, failed precondition), not when the provider failed. */
export function isInvalidRequest(error: unknown): boolean {
  if (typeof error !== 'object' || error === null) return false;
  const { response, status, code } = error as GoogleApiError;
  return Number(response?.status ?? status ?? code) === 400;
}
