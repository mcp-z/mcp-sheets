// Network: first-row CSV cancellation uses a real loopback HTTP server.
import '../lib/env-loader.ts';
import { mcp } from '@mcp-z/mcp-sheets';
import assert from 'assert';
import { once } from 'events';
import { createServer, type ServerResponse } from 'http';
import type { AddressInfo } from 'net';
import type { Output } from '../../src/mcp/tools/csv-get-columns.ts';
import { createExtra } from '../lib/create-extra.ts';

const handler = mcp.toolFactories.csvGetColumns().handler;

// About 40 MiB is the maximum body; backpressure and cancellation limit actual writes.
it('sheets-csv-get-columns cancels a real HTTP body after reading its first row', async () => {
  const chunk = 'value1,value2,value3\n'.repeat(2048);
  const totalChunks = 1024;
  let sentChunks = 0;
  let response: ServerResponse | undefined;
  let resolveClosed!: (finished: boolean) => void;
  const responseClosed = new Promise<boolean>((resolve) => {
    resolveClosed = resolve;
  });
  const server = createServer((_request, res) => {
    response = res;
    res.once('close', () => resolveClosed(res.writableFinished));
    res.writeHead(200, { 'Content-Type': 'text/csv' });
    res.write('Header1,Header2,Header3\n');
    function writeRows() {
      while (!res.destroyed && sentChunks < totalChunks) {
        sentChunks++;
        if (!res.write(chunk)) {
          res.once('drain', writeRows);
          return;
        }
      }
      if (!res.destroyed) res.end();
    }
    writeRows();
  });
  const failures: unknown[] = [];
  try {
    const listening = once(server, 'listening');
    server.listen(0, '127.0.0.1');
    await listening;
    const { port } = server.address() as AddressInfo;
    const result = await handler({ sourceUri: `http://127.0.0.1:${port}/first-row.csv` }, createExtra());
    const structured = (result.structuredContent as { result: Output }).result;
    assert.equal(structured.type, 'success');
    if (structured.type === 'success') {
      assert.deepEqual(structured.columns, ['Header1', 'Header2', 'Header3']);
      assert.equal(structured.isEmpty, false);
    }
    assert.equal(await responseClosed, false, 'The client must cancel before the response finishes');
    assert.ok(sentChunks < totalChunks, 'First-row reading must not consume the complete body');
  } catch (error) {
    failures.push(error);
  }
  try {
    response?.destroy();
    if (server.listening) {
      const closed = new Promise<void>((resolve, reject) => server.close((error) => (error ? reject(error) : resolve())));
      server.closeAllConnections();
      await closed;
    }
  } catch (error) {
    failures.push(error);
  }
  if (failures.length) {
    const [primary] = failures;
    if (failures.length > 1 && primary instanceof Error) Object.assign(primary, { cleanupErrors: failures.slice(1) });
    throw primary;
  }
});
