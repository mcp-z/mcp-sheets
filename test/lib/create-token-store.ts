import type { Keyv } from 'keyv';
import keyvRegistry, { migrateKeyv5Keys } from 'keyv-registry';

// The local .tokens/ store. Sign-ins that keyv 5 saved are moved to the keyv 6 keys once; later runs find nothing to move.
export default async function createTokenStore(): Promise<Keyv<unknown>> {
  const store = await keyvRegistry<unknown>('file://.//.tokens/store.json');
  if (!store) throw new Error('Failed to create the .tokens store');
  await migrateKeyv5Keys(store as Keyv);
  return store;
}
