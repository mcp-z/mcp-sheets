import { spawnSync } from 'node:child_process';
import { lstat, readFile } from 'node:fs/promises';
import path from 'node:path';

export type AccountStorage = {
  readonly kind: 'account';
  readonly path: string;
  readonly service: string;
  readonly accountId: { readonly env: string } | { readonly value: string };
};

export type DcrStorage = {
  readonly kind: 'dcr';
  readonly path: string;
  readonly key: string;
  readonly registration: { readonly clientIdSecret: string; readonly clientSecretSecret: string };
};

export type Flow = {
  readonly id: string;
  readonly refreshTokenSecret: string;
  readonly scopeSecret: string;
  readonly clientIdSecret: string;
  readonly clientSecretSecret: string;
  readonly clientSecretRequired: boolean;
  readonly tenantSecret: string | null;
  readonly storage: AccountStorage | DcrStorage;
};

export type DcrFlow = Flow & { readonly storage: DcrStorage };

function isDcrFlow(flow: Flow): flow is DcrFlow {
  return flow.storage.kind === 'dcr';
}

export type CiConfig = {
  readonly provider: 'google' | 'microsoft';
  readonly flows: readonly Flow[];
  readonly requiredTestSecrets: readonly string[];
  readonly serviceAccount: { readonly output: string; readonly secrets: Readonly<Record<string, string>> } | null;
};

export type Env = Readonly<Record<string, string | undefined>>;

export type SecretSource =
  | { readonly kind: 'account-id'; readonly env: string }
  | { readonly kind: 'scope' }
  | { readonly kind: 'env' }
  | { readonly kind: 'registration'; readonly flow: DcrFlow; readonly field: 'clientId' | 'clientSecret' }
  | { readonly kind: 'service-account'; readonly field: string }
  | { readonly kind: 'refresh-token'; readonly flow: Flow };

export type PlannedSecret = { readonly name: string; readonly required: boolean; readonly source: SecretSource };

export type SetSecret = (name: string, value: string, repository: string) => void;

export async function exists(filePath: string): Promise<boolean> {
  try {
    await lstat(filePath);
    return true;
  } catch (error) {
    if (error && typeof error === 'object' && 'code' in error && error.code === 'ENOENT') return false;
    throw error;
  }
}

export async function openStore(filePath: string) {
  const [{ default: Keyv }, { KeyvFile }] = await Promise.all([import('keyv'), import('keyv-file')]);
  return new Keyv({ store: new KeyvFile({ filename: filePath }) });
}

export function flowAccountId(storage: Flow['storage'], humanAccountId: string): string | undefined {
  if (storage.kind !== 'account') return undefined;
  return 'env' in storage.accountId ? humanAccountId : storage.accountId.value;
}

export async function readFlowRefreshToken(root: string, flow: Flow, accountId: string): Promise<string | undefined> {
  const filePath = path.join(root, flow.storage.path);
  if (!(await exists(filePath))) return undefined;
  const store = await openStore(filePath);
  try {
    if (flow.storage.kind === 'account') {
      const effectiveAccountId = flowAccountId(flow.storage, accountId);
      if (!effectiveAccountId) return undefined;
      const token = await (await import('@mcp-z/oauth')).getToken<{ refreshToken?: string }>(store, { accountId: effectiveAccountId, service: flow.storage.service });
      return token?.refreshToken;
    }
    const token = (await store.get(flow.storage.key)) as { providerRefreshToken?: string } | undefined;
    return token?.providerRefreshToken;
  } finally {
    await store.disconnect();
  }
}

export async function readRepository(root: string): Promise<string> {
  const manifest = JSON.parse(await readFile(path.join(root, 'package.json'), 'utf8')) as { repository?: { url?: string } };
  const repositoryUrl = manifest.repository?.url?.replace(/^git\+/, '');
  if (!repositoryUrl) throw new Error('Package does not declare its GitHub repository');
  const url = new URL(repositoryUrl);
  if (url.hostname !== 'github.com') throw new Error('Package repository must use GitHub');
  return url.pathname.replace(/^\//, '').replace(/\.git$/, '');
}

/** Every secret `ci-tokens.ts seed` reads, in the order they are set: refresh tokens last, because the seed writes newer ones back. */
export function planSecrets(config: CiConfig): PlannedSecret[] {
  const planned = new Map<string, PlannedSecret>();
  const add = (name: string, required: boolean, source: SecretSource) => {
    const previous = planned.get(name);
    planned.set(name, { name, required: required || Boolean(previous?.required), source: previous?.source ?? source });
  };

  for (const flow of config.flows) {
    if (flow.storage.kind === 'account' && 'env' in flow.storage.accountId) add(flow.storage.accountId.env, true, { kind: 'account-id', env: flow.storage.accountId.env });
  }
  for (const flow of config.flows) add(flow.scopeSecret, true, { kind: 'scope' });
  for (const flow of config.flows) {
    add(flow.clientIdSecret, true, { kind: 'env' });
    add(flow.clientSecretSecret, flow.clientSecretRequired, { kind: 'env' });
    if (flow.tenantSecret) add(flow.tenantSecret, true, { kind: 'env' });
  }
  for (const name of config.requiredTestSecrets) add(name, true, { kind: 'env' });
  for (const flow of config.flows) {
    if (!isDcrFlow(flow)) continue;
    add(flow.storage.registration.clientIdSecret, true, { kind: 'registration', flow, field: 'clientId' });
    add(flow.storage.registration.clientSecretSecret, true, { kind: 'registration', flow, field: 'clientSecret' });
  }
  if (config.serviceAccount) {
    for (const [field, name] of Object.entries(config.serviceAccount.secrets)) add(name, true, { kind: 'service-account', field });
  }
  for (const flow of config.flows) add(flow.refreshTokenSecret, true, { kind: 'refresh-token', flow });
  return [...planned.values()];
}

export function missingRequired(plan: readonly PlannedSecret[], values: ReadonlyMap<string, string>): string[] {
  return plan.filter((secret) => secret.required && !values.has(secret.name)).map((secret) => secret.name);
}

async function readAccountId(config: CiConfig, root: string): Promise<string | undefined> {
  const accounts = new Set<string>();
  const checked = new Set<string>();
  const stores: string[] = [];
  for (const flow of config.flows) {
    const { storage } = flow;
    if (storage.kind !== 'account' || !('env' in storage.accountId)) continue;
    const filePath = path.join(root, storage.path);
    const checkKey = `${filePath}:${storage.service}`;
    if (checked.has(checkKey) || !(await exists(filePath))) continue;
    checked.add(checkKey);
    stores.push(storage.path);

    // Flows with a fixed account id (device code) are not the human test account.
    const fixed = new Set(config.flows.flatMap((other) => (other.storage.kind === 'account' && other.storage.service === storage.service && 'value' in other.storage.accountId ? [other.storage.accountId.value] : [])));
    const store = await openStore(filePath);
    try {
      for (const accountId of await (await import('@mcp-z/oauth')).listAccountIds(store, storage.service)) {
        if (!fixed.has(accountId)) accounts.add(accountId);
      }
    } finally {
      await store.disconnect();
    }
  }
  if (accounts.size > 1) throw new Error(`${accounts.size} accounts hold a token for this package (${stores.join(', ')}). Keep one test account and run the sign-in again.`);
  return [...accounts][0];
}

async function readRegistration(root: string, storage: DcrStorage): Promise<{ clientId?: string; clientSecret?: string } | undefined> {
  const filePath = path.join(root, storage.path);
  if (!(await exists(filePath))) return undefined;
  const store = await openStore(filePath);
  try {
    return (await store.get(storage.key)) as { clientId?: string; clientSecret?: string } | undefined;
  } finally {
    await store.disconnect();
  }
}

async function readServiceAccount(root: string, output: string): Promise<Record<string, unknown> | undefined> {
  const filePath = path.join(root, output);
  if (!(await exists(filePath))) return undefined;
  return JSON.parse(await readFile(filePath, 'utf8')) as Record<string, unknown>;
}

/** Reads the value of every planned secret that exists locally. A missing store, file or key leaves the secret out; an unreadable one throws. */
export async function readLocalValues(config: CiConfig, root: string, env: Env, scope: string): Promise<Map<string, string>> {
  const plan = planSecrets(config);
  const values = new Map<string, string>();
  const set = (name: string, value: unknown) => {
    if (typeof value === 'string' && value) values.set(name, value);
  };

  const accountId = await readAccountId(config, root);
  const registrations = new Map<string, { clientId?: string; clientSecret?: string } | undefined>();
  const serviceAccount = config.serviceAccount ? await readServiceAccount(root, config.serviceAccount.output) : undefined;

  for (const { name, source } of plan) {
    if (source.kind === 'account-id') set(name, accountId);
    else if (source.kind === 'scope') set(name, scope);
    else if (source.kind === 'env') set(name, env[name]);
    else if (source.kind === 'service-account') set(name, serviceAccount?.[source.field]);
    else if (source.kind === 'registration') {
      if (!registrations.has(source.flow.id)) registrations.set(source.flow.id, await readRegistration(root, source.flow.storage));
      set(name, registrations.get(source.flow.id)?.[source.field]);
    } else set(name, await readFlowRefreshToken(root, source.flow, accountId ?? ''));
  }
  return values;
}

export type UploadOptions = {
  readonly config: CiConfig;
  readonly root: string;
  readonly env: Env;
  readonly scope: string;
  readonly repository: string;
  readonly dry: boolean;
  readonly setSecret: SetSecret;
};

/** Sets every planned secret that exists locally, or reports them without setting when dry or when a required one is missing. Never returns a value. */
export async function upload(options: UploadOptions): Promise<{ lines: string[]; ok: boolean }> {
  const plan = planSecrets(options.config);
  const values = await readLocalValues(options.config, options.root, options.env, options.scope);
  const missing = missingRequired(plan, values);
  const lines: string[] = [];

  if (options.dry || missing.length > 0) {
    for (const { name, required } of plan) {
      if (values.has(name)) lines.push(`${name}: present${options.dry ? ' (dry)' : ' (not set: required secrets are missing)'}`);
      else lines.push(`${name}: ${required ? 'missing' : 'absent (optional)'}`);
    }
    return { lines, ok: missing.length === 0 };
  }

  let failed = false;
  for (const { name } of plan) {
    const value = values.get(name);
    if (value === undefined) {
      lines.push(`${name}: absent (optional)`);
      continue;
    }
    if (failed) {
      lines.push(`${name}: not set (an earlier secret failed)`);
      continue;
    }
    try {
      options.setSecret(name, value, options.repository);
      lines.push(`${name}: set`);
    } catch (error) {
      failed = true;
      lines.push(`${name}: FAILED ${redact(error instanceof Error ? error.message : 'unknown error', [...values.values()])}`);
    }
  }
  return { lines, ok: !failed };
}

export function redact(message: string, secretValues: readonly string[]): string {
  let sanitized = message;
  for (const secret of [...new Set(secretValues)].sort((left, right) => right.length - left.length)) sanitized = sanitized.split(secret).join('[REDACTED]');
  return sanitized.split('\n')[0].slice(0, 300);
}

/** Sets one `live-test` environment secret with `gh`, passing the value on stdin. */
export function ghSetSecret(name: string, value: string, repository: string): void {
  const result = spawnSync('gh', ['secret', 'set', name, '--env', 'live-test', '--repo', repository], {
    input: value,
    encoding: 'utf8',
    timeout: 30000,
    stdio: ['pipe', 'ignore', 'pipe'],
  });
  if (result.error || result.status !== 0) throw new Error(result.error?.message || result.stderr?.trim() || `gh exited with status ${result.status}`);
}

/** Reports whether CI_SECRETS_TOKEN exists in the `live-test` environment. It is never uploaded. */
export function secretsTokenStatus(repository: string): string {
  const result = spawnSync('gh', ['secret', 'list', '--env', 'live-test', '--repo', repository, '--json', 'name'], { encoding: 'utf8', timeout: 30000 });
  if (result.error || result.status !== 0) return `CI_SECRETS_TOKEN: unknown (${redact(result.error?.message || result.stderr?.trim() || `gh exited with status ${result.status}`, [])})`;
  const names = (JSON.parse(result.stdout) as { name: string }[]).map((entry) => entry.name);
  return names.includes('CI_SECRETS_TOKEN') ? 'CI_SECRETS_TOKEN: exists (not uploaded)' : 'CI_SECRETS_TOKEN: missing (create the fine-grained PAT described in CONTRIBUTING; it is not uploaded)';
}
