import type Keyv from 'keyv';
import keyvRegistry from 'keyv-registry';

export default async function createStore<T>(uri: string): Promise<Keyv<T>> {
  const store = await keyvRegistry<T>(uri);
  if (!store) throw new Error(`Failed to create store for URI: ${uri}`);
  return store;
}
