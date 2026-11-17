import assert from 'assert';
import { copyFileSync, mkdirSync, rmSync } from 'fs';
import { KeyvFile } from 'keyv-file';
import * as path from 'path';
import { pathToFileURL } from 'url';
import createStore from '../../../src/lib/create-store.ts';
import { createTokenStore } from '../../../src/setup/runtime.ts';

// test/data/keyv5-tokens.json was written by keyv 5 with dummy values.
const dir = path.resolve('.tmp', `create-store-${process.pid}`);

describe('create-store keyv 5 migration', () => {
  after(() => rmSync(dir, { recursive: true, force: true }));

  it('finds a token saved by keyv 5 through the token store', async () => {
    mkdirSync(dir, { recursive: true });
    copyFileSync(path.resolve('test/data/keyv5-tokens.json'), path.join(dir, 'tokens.json'));
    const messages: string[] = [];
    const logger = { info: (msg: string) => messages.push(msg) } as never;
    const saved = process.env.TOKEN_STORE_URI;
    delete process.env.TOKEN_STORE_URI;
    try {
      const store = await createTokenStore(dir, logger);
      const token = (await store.get('user@example.com:sheets:token')) as { accessToken?: string } | undefined;
      assert.strictEqual(token?.accessToken, 'dummy-access');
      assert.deepStrictEqual(messages, ['Moved 4 saved sign-in entries to the keyv 6 format']);
      await store.disconnect();
    } finally {
      if (saved !== undefined) process.env.TOKEN_STORE_URI = saved;
    }
  });

  it('logs nothing and does not log the URI when there is nothing to move', async () => {
    mkdirSync(dir, { recursive: true });
    const messages: string[] = [];
    const store = await createStore<unknown>(pathToFileURL(path.join(dir, 'empty.json')).href, { info: (msg: string) => messages.push(msg) } as never);
    assert.deepStrictEqual(messages, []);
    await store.disconnect();
  });

  it('rejects, and the server does not start, when a store read fails during the migration', async () => {
    mkdirSync(dir, { recursive: true });
    const file = path.join(dir, 'failing.json');
    copyFileSync(path.resolve('test/data/keyv5-tokens.json'), file);
    const original = KeyvFile.prototype.get;
    let failed = false;
    KeyvFile.prototype.get = async function (this: KeyvFile, key: string) {
      if (!failed && key.endsWith(':active')) {
        failed = true;
        throw new Error('injected read failure');
      }
      return original.call(this, key);
    } as typeof original;
    try {
      await assert.rejects(createStore<unknown>(pathToFileURL(file).href), /active/);
      assert.ok(failed, 'the failure was injected');
    } finally {
      KeyvFile.prototype.get = original;
    }
  });

  it('names only the protocol when the store cannot be created, never the URI', async () => {
    await assert.rejects(createStore<unknown>('nosuchproto://user:secretpw@host/0'), (err: Error) => err.message.startsWith('Failed to create store for nosuchproto:') && !err.message.includes('secretpw'));
  });
});
