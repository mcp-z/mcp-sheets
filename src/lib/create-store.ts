import type Keyv from 'keyv';
import keyvRegistry from 'keyv-registry';

// keyv-file is a direct dependency so the default file:// URIs resolve; keyv-registry
// loads adapters with require() rather than depending on any of them itself.

export default async function createStore<T>(uri: string): Promise<Keyv<T>> {
  const store = await keyvRegistry<T>(uri);
  if (!store) throw new Error(`Failed to create store for URI: ${uri}`);
  return store;
}
