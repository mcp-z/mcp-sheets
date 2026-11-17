import type { Logger } from '@mcp-z/server';
import type Keyv from 'keyv';
import keyvRegistry, { migrateKeyv5Keys } from 'keyv-registry';

// keyv-file is a direct dependency so the default file:// URIs resolve; keyv-registry
// loads adapters with require() rather than depending on any of them itself.

export default async function createStore<T>(uri: string, logger?: Logger): Promise<Keyv<T>> {
  // Errors name the protocol only: a redis URI can hold a password.
  const protocol = `${uri.split(':')[0]}:`;
  let store: Keyv<T> | undefined;
  try {
    store = await keyvRegistry<T>(uri);
  } catch (cause) {
    const error = new Error(`Failed to create store for ${protocol}`);
    (error as Error & { cause?: unknown }).cause = cause;
    throw error;
  }
  if (!store) throw new Error(`Failed to create store for ${protocol}`);
  try {
    // One-time move of sign-ins saved by keyv 5 versions of this server. Removed in the next major.
    const moved = await migrateKeyv5Keys(store as Keyv);
    if (moved > 0) logger?.info(`Moved ${moved} saved sign-in entries to the keyv 6 format`);
  } catch (error) {
    // Keep the migration error as the one reported, even if closing the store fails too.
    await store.disconnect().catch(() => {});
    throw error;
  }
  return store;
}
