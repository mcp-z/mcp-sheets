/** Streaming CSV parsing utilities for memory-efficient large file processing */

import { createReadStream } from 'fs';
import { resolve } from 'path';
import { Readable } from 'stream';
import type { ReadableStream as NodeReadableStream } from 'stream/web';

/**
 * Get readable stream from CSV URI
 *
 * Memory efficiency:
 * - file:// URIs stream directly from disk
 * - http:// URIs stream directly from response (no temp files!)
 *
 * @example
 * ```ts
 * const readStream = await getCsvReadStream(csvUri);
 * const parser = readStream.pipe(parse({ columns: true }));
 * for await (const record of parser) {
 *   // Process record
 * }
 * ```
 */
export async function getCsvReadStream(csvUri: string): Promise<Readable> {
  if (csvUri.startsWith('file://')) {
    // Local file - stream directly from disk
    const filePath = csvUri.replace('file://', '');
    const resolvedPath = resolve(filePath);
    return createReadStream(resolvedPath, { encoding: 'utf-8' });
  }

  if (csvUri.startsWith('http://') || csvUri.startsWith('https://')) {
    // Remote file - stream directly from fetch response
    const response = await fetch(csvUri);
    if (!response.ok) {
      throw new Error(`Failed to fetch CSV from ${csvUri}: ${response.statusText}`);
    }

    if (!response.body) {
      throw new Error(`No response body from ${csvUri}`);
    }

    // Convert web stream to Node.js stream
    // response.body is ReadableStream<Uint8Array> from fetch API
    // Cast to Node.js ReadableStream type for compatibility with Readable.fromWeb
    return Readable.fromWeb(response.body as unknown as NodeReadableStream<Uint8Array>);
  }

  throw new Error(`Invalid CSV URI: ${csvUri}. Must start with file://, http://, or https://`);
}
