import type Keyv from 'keyv';
import keyvRegistry from 'keyv-registry';

function parseDefaultTtl(uri: string): { uri: string; ttl?: number } {
  try {
    const url = new URL(uri);
    const ttlParam = url.searchParams.get('ttl');
    const ttlSecondsParam = url.searchParams.get('ttlSeconds');
    const ttlMs = ttlSecondsParam ? Number(ttlSecondsParam) * 1000 : ttlParam ? Number(ttlParam) : undefined;
    url.searchParams.delete('ttl');
    url.searchParams.delete('ttlSeconds');
    return Number.isFinite(ttlMs) && (ttlMs as number) > 0 ? { uri: url.toString(), ttl: ttlMs } : { uri: url.toString() };
  } catch {
    return { uri };
  }
}

export default async function createStore<T>(uri: string): Promise<Keyv<T>> {
  const { uri: parsedUri, ttl: defaultTtl } = parseDefaultTtl(uri);
  const store = await keyvRegistry<T>(parsedUri);
  if (!store) throw new Error(`Failed to create store for URI: ${uri}`);
  if (defaultTtl !== undefined) {
    const originalSet = store.set.bind(store);
    store.set = ((key, value, ttl) => originalSet(key, value, ttl ?? defaultTtl)) as typeof store.set;
  }
  return store;
}
