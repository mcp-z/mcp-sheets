import { spawnSync } from 'node:child_process';
import { chmod, lstat, mkdir, readFile, rm, writeFile } from 'node:fs/promises';
import path from 'node:path';
import type { CachedToken } from '@mcp-z/oauth';

const config = {
  provider: 'google',
  flows: [
    {
      id: 'mcp-sheets',
      refreshTokenSecret: 'TEST_REFRESH_TOKEN',
      scopeSecret: 'TEST_SCOPE',
      clientIdSecret: 'GOOGLE_CLIENT_ID',
      clientSecretSecret: 'GOOGLE_CLIENT_SECRET',
      clientSecretRequired: false,
      tenantSecret: null,
      storage: {
        kind: 'account',
        path: '.tokens/store.json',
        service: 'mcp-sheets',
        accountId: {
          env: 'TEST_ACCOUNT_ID',
        },
      },
    },
  ],
  requiredTestSecrets: [],
  serviceAccount: null,
} as const;

const root = path.resolve(import.meta.dirname, '..');
const ownerFile = path.join(root, '.ci-live-token-owner');
const envFile = path.join(root, '.env.test');
const tokenRoot = path.join(root, '.tokens');

type TokenResponse = {
  access_token?: string;
  refresh_token?: string;
  expires_in?: number;
  scope?: string;
  error?: string;
};

function required(name: string): string {
  const value = process.env[name];
  if (!value) throw new Error(`Missing ${name} in the live-test environment`);
  return value;
}

function mask(value: string): void {
  const escaped = value.replaceAll('%', '%25').replaceAll('\r', '%0D').replaceAll('\n', '%0A');
  process.stdout.write(`::add-mask::${escaped}\n`);
}

function optional(name: string | null): string | undefined {
  const value = name ? process.env[name] : undefined;
  return value || undefined;
}

function redactDiagnostic(message: string, extra: string[]): string {
  const secretNames = config.flows
    .flatMap((flow) => [flow.refreshTokenSecret, flow.scopeSecret, flow.clientIdSecret, flow.clientSecretSecret, flow.tenantSecret, ...(flow.storage.kind === 'dcr' ? [flow.storage.registration.clientIdSecret, flow.storage.registration.clientSecretSecret] : [])])
    .filter((name): name is string => Boolean(name));
  if (config.serviceAccount) secretNames.push(...Object.values(config.serviceAccount.secrets));
  secretNames.push(...config.requiredTestSecrets);
  const secretValues = [...new Set([...secretNames.map((name) => process.env[name]), process.env.CI_SECRETS_TOKEN, ...extra].filter((value): value is string => Boolean(value)))].sort((left, right) => right.length - left.length);

  let sanitized = message;
  for (const secret of secretValues) sanitized = sanitized.replaceAll(secret, '[REDACTED]');
  return sanitized.slice(0, 600);
}

async function exists(filePath: string): Promise<boolean> {
  try {
    await lstat(filePath);
    return true;
  } catch (error) {
    if (error && typeof error === 'object' && 'code' in error && error.code === 'ENOENT') return false;
    throw error;
  }
}

async function assertFreshCheckout(): Promise<void> {
  const paths = [envFile, tokenRoot, ownerFile];
  if (config.serviceAccount) paths.push(path.join(root, config.serviceAccount.output));

  for (const filePath of paths) {
    if (await exists(filePath)) throw new Error(`Refusing to overwrite existing ${path.basename(filePath)}`);
  }
}

async function assertActionsContext(): Promise<{ repository: string }> {
  if (process.env.GITHUB_ACTIONS !== 'true') throw new Error('This command is only for an ephemeral GitHub Actions checkout');
  if (process.env.GITHUB_REF !== 'refs/heads/master') throw new Error('Live tests may run only from master');

  const manifest = JSON.parse(await readFile(path.join(root, 'package.json'), 'utf8')) as { repository?: { url?: string } };
  const repository = required('GITHUB_REPOSITORY');
  const repositoryUrl = manifest.repository?.url?.replace(/^git\+/, '');
  if (!repositoryUrl) throw new Error('Package does not declare its GitHub repository');
  const expectedUrl = new URL(repositoryUrl);
  if (expectedUrl.hostname !== 'github.com') throw new Error('Package repository must use GitHub');
  const expectedRepository = expectedUrl.pathname.replace(/^\//, '').replace(/\.git$/, '');
  if (repository !== expectedRepository) throw new Error('Repository does not match this package');

  return { repository };
}

function flowAccountId(storage: (typeof config.flows)[number]['storage'], humanAccountId: string): string | undefined {
  if (storage.kind !== 'account') return undefined;
  return 'env' in storage.accountId ? humanAccountId : storage.accountId.value;
}

function validatedFlows(humanAccountId: string) {
  return config.flows.map((flow) => {
    const refreshToken = required(flow.refreshTokenSecret);
    const scope = required(flow.scopeSecret);
    const clientId = required(flow.clientIdSecret);
    const clientSecret = optional(flow.clientSecretSecret);
    if (flow.clientSecretRequired && !clientSecret) throw new Error(`Missing ${flow.clientSecretSecret} in the live-test environment`);

    const tenantId = flow.tenantSecret ? required(flow.tenantSecret) : undefined;
    if (tenantId && !/^[a-zA-Z0-9.-]+$/.test(tenantId)) throw new Error(`Invalid ${flow.tenantSecret}`);
    const registration =
      flow.storage.kind === 'dcr'
        ? {
            clientId: required(flow.storage.registration.clientIdSecret),
            clientSecret: required(flow.storage.registration.clientSecretSecret),
          }
        : undefined;

    mask(refreshToken);
    if (clientSecret) mask(clientSecret);
    if (registration) {
      mask(registration.clientId);
      if (registration.clientSecret) mask(registration.clientSecret);
    }

    return { flow, refreshToken, scope, clientId, clientSecret, tenantId, registration, accountId: flowAccountId(flow.storage, humanAccountId) };
  });
}

function writeSecret(name: string, value: string, context: { repository: string; secretsToken: string }): void {
  mask(value);
  const ghEnv = { ...process.env, GH_TOKEN: context.secretsToken, GITHUB_TOKEN: '', GH_ENTERPRISE_TOKEN: '', GITHUB_ENTERPRISE_TOKEN: '' };
  delete ghEnv.CI_SECRETS_TOKEN;

  const result = spawnSync('gh', ['secret', 'set', name, '--env', 'live-test', '--repo', context.repository], {
    input: value,
    encoding: 'utf8',
    timeout: 30000,
    env: ghEnv,
    stdio: ['pipe', 'ignore', 'pipe'],
  });
  if (result.error || result.status !== 0) {
    const diagnostic = result.error?.message || result.stderr?.trim() || `gh exited with status ${result.status}`;
    const safeDiagnostic = redactDiagnostic(diagnostic, [value, context.secretsToken]);
    throw new Error(`Could not persist ${name}: ${safeDiagnostic}. Check CI_SECRETS_TOKEN has Environments:write access to this repository.`);
  }
}

async function openStore(filePath: string) {
  const [{ default: Keyv }, { KeyvFile }] = await Promise.all([import('keyv'), import('keyv-file')]);
  return new Keyv({ store: new KeyvFile({ filename: filePath }) });
}

async function saveFlowToken(input: ReturnType<typeof validatedFlows>[number], token: CachedToken): Promise<void> {
  const { flow } = input;
  const filePath = path.join(root, flow.storage.path);
  const store = await openStore(filePath);
  try {
    if (flow.storage.kind === 'account') {
      const accountId = input.accountId;
      if (!accountId) throw new Error(`Missing account identity for ${flow.id}`);
      const params = { accountId, service: flow.storage.service };
      const api = await import('@mcp-z/oauth');
      await api.setToken(store, params, token);
      await api.addAccount(store, params);
      await api.setAccountInfo(store, params, { email: accountId, addedAt: new Date().toISOString() });
    } else {
      if (!input.registration) throw new Error(`Missing saved DCR client registration for ${flow.id}`);
      await store.set(flow.storage.key, {
        providerAccessToken: token.accessToken,
        providerRefreshToken: token.refreshToken,
        providerExpiresAt: token.expiresAt,
        clientId: input.registration.clientId,
        clientSecret: input.registration.clientSecret || '',
      });
    }
  } finally {
    await store.disconnect();
  }
  await chmod(filePath, 0o600);
}

async function refresh(input: ReturnType<typeof validatedFlows>[number]): Promise<CachedToken> {
  const body = new URLSearchParams({
    grant_type: 'refresh_token',
    refresh_token: input.refreshToken,
    client_id: input.clientId,
  });
  if (config.provider === 'microsoft') body.set('scope', input.scope);
  if (input.clientSecret) body.set('client_secret', input.clientSecret);

  const endpoint = config.provider === 'microsoft' ? `https://login.microsoftonline.com/${input.tenantId}/oauth2/v2.0/token` : 'https://oauth2.googleapis.com/token';
  const response = await fetch(endpoint, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body,
    signal: AbortSignal.timeout(30000),
  });
  let payload: TokenResponse;
  try {
    payload = (await response.json()) as TokenResponse;
  } catch {
    throw new Error(`${input.flow.id} token refresh returned invalid JSON (HTTP ${response.status})`);
  }
  if (!response.ok) {
    const code = typeof payload.error === 'string' && /^[a-zA-Z0-9_.-]+$/.test(payload.error) ? payload.error : 'unknown_error';
    throw new Error(`${input.flow.id} token refresh failed: HTTP ${response.status}, ${code}. Reauthorization may be required.`);
  }
  if (!payload.access_token || typeof payload.expires_in !== 'number' || !Number.isFinite(payload.expires_in) || payload.expires_in <= 0) {
    throw new Error(`${input.flow.id} token refresh returned an incomplete response`);
  }

  const refreshToken = payload.refresh_token || input.refreshToken;
  mask(payload.access_token);
  mask(refreshToken);
  return {
    accessToken: payload.access_token,
    refreshToken,
    expiresAt: Date.now() + payload.expires_in * 1000,
    scope: payload.scope || input.scope,
  };
}

async function preparePrivateFiles(serviceAccount: Record<string, string> | undefined): Promise<void> {
  await writeFile(ownerFile, `${JSON.stringify({ seededFlows: [] })}\n`, { flag: 'wx', mode: 0o600 });
  await writeFile(envFile, 'NODE_ENV=test\nHEADLESS=true\nTEST_INCLUDE_MANUAL=false\n', { flag: 'wx', mode: 0o600 });
  await mkdir(tokenRoot, { mode: 0o700 });
  await chmod(tokenRoot, 0o700);
  for (const flow of config.flows) {
    const directory = path.dirname(path.join(root, flow.storage.path));
    await mkdir(directory, { recursive: true, mode: 0o700 });
    await chmod(directory, 0o700);
  }
  if (config.serviceAccount && serviceAccount) {
    const filePath = path.join(root, config.serviceAccount.output);
    await writeFile(filePath, `${JSON.stringify({ type: 'service_account', ...serviceAccount }, null, 2)}\n`, { flag: 'wx', mode: 0o600 });
    await chmod(filePath, 0o600);
  }
}

async function markSeeded(flowId: string): Promise<void> {
  const manifest = JSON.parse(await readFile(ownerFile, 'utf8')) as { seededFlows?: unknown };
  if (!Array.isArray(manifest.seededFlows) || !manifest.seededFlows.every((value) => typeof value === 'string')) {
    throw new Error('CI token ownership manifest is invalid');
  }
  if (!manifest.seededFlows.includes(flowId)) manifest.seededFlows.push(flowId);
  await writeFile(ownerFile, `${JSON.stringify(manifest)}\n`, { mode: 0o600 });
}

async function validateInputs() {
  const context = await assertActionsContext();
  const secretsToken = required('CI_SECRETS_TOKEN');
  mask(secretsToken);
  const needsHumanId = config.flows.some((flow) => flow.storage.kind === 'account' && 'env' in flow.storage.accountId);
  const accountId = needsHumanId ? required('TEST_ACCOUNT_ID') : '';
  if (accountId) mask(accountId);
  const flows = validatedFlows(accountId);
  for (const name of config.requiredTestSecrets) required(name);

  let serviceAccount: Record<string, string> | undefined;
  if (config.serviceAccount) {
    serviceAccount = Object.fromEntries(Object.entries(config.serviceAccount.secrets).map(([field, name]) => [field, required(name)]));
    mask(serviceAccount.private_key);
    for (const field of ['auth_uri', 'token_uri']) {
      const address = new URL(serviceAccount[field]);
      if (address.protocol !== 'https:') throw new Error(`Invalid ${config.serviceAccount.secrets[field as keyof typeof config.serviceAccount.secrets]}`);
    }
  }

  return { context: { ...context, secretsToken }, accountId, flows, serviceAccount };
}

async function seed(): Promise<void> {
  await assertFreshCheckout();
  const inputs = await validateInputs();
  const firstFlow = inputs.flows[0];
  if (!firstFlow) throw new Error('No token flows configured');
  writeSecret(firstFlow.flow.refreshTokenSecret, firstFlow.refreshToken, inputs.context);
  await preparePrivateFiles(inputs.serviceAccount);

  for (const input of inputs.flows) {
    const token = await refresh(input);
    await saveFlowToken(input, token);
    await markSeeded(input.flow.id);
    writeSecret(input.flow.refreshTokenSecret, token.refreshToken as string, inputs.context);
  }
}

async function readFlowRefreshToken(flow: (typeof config.flows)[number], accountId: string): Promise<string | undefined> {
  const filePath = path.join(root, flow.storage.path);
  if (!(await exists(filePath))) return undefined;
  const store = await openStore(filePath);
  try {
    if (flow.storage.kind === 'account') {
      const effectiveAccountId = flowAccountId(flow.storage, accountId);
      if (!effectiveAccountId) return undefined;
      const token = await (await import('@mcp-z/oauth')).getToken<CachedToken>(store, { accountId: effectiveAccountId, service: flow.storage.service });
      return token?.refreshToken;
    }
    const token = (await store.get(flow.storage.key)) as { providerRefreshToken?: string } | undefined;
    return token?.providerRefreshToken;
  } finally {
    await store.disconnect();
  }
}

async function readOwnerManifest(): Promise<string[]> {
  const manifest = JSON.parse(await readFile(ownerFile, 'utf8')) as { seededFlows?: unknown };
  if (!Array.isArray(manifest.seededFlows) || !manifest.seededFlows.every((value) => typeof value === 'string')) {
    throw new Error('CI token ownership manifest is invalid');
  }
  const knownIds = new Set(config.flows.map((flow) => flow.id));
  if (manifest.seededFlows.some((flowId) => !knownIds.has(flowId))) throw new Error('CI token ownership manifest contains an unknown flow');
  return manifest.seededFlows;
}

async function persist(): Promise<void> {
  const context = await assertActionsContext();
  const seedOutcome = required('CI_SEED_OUTCOME');
  if (!['success', 'failure', 'cancelled'].includes(seedOutcome)) throw new Error('Invalid CI_SEED_OUTCOME');
  if (!(await exists(ownerFile))) {
    if (seedOutcome === 'success') throw new Error('Seed succeeded without creating its ownership manifest');
    console.log('Seed created no owned token files; its failed result remains authoritative.');
    return;
  }

  const accountId = config.flows.some((flow) => flow.storage.kind === 'account' && 'env' in flow.storage.accountId) ? required('TEST_ACCOUNT_ID') : '';
  const secretsToken = required('CI_SECRETS_TOKEN');
  mask(secretsToken);
  const secretContext = { ...context, secretsToken };
  const seededFlowIds = new Set(await readOwnerManifest());
  const errors: string[] = [];
  if (seedOutcome === 'success') {
    for (const flow of config.flows) {
      if (!seededFlowIds.has(flow.id)) errors.push(`${flow.refreshTokenSecret} was not marked as seeded`);
    }
  }

  let available = 0;
  for (const flow of config.flows) {
    let refreshToken: string | undefined;
    try {
      refreshToken = await readFlowRefreshToken(flow, accountId);
    } catch {
      errors.push(`${flow.refreshTokenSecret} could not be read`);
      continue;
    }
    if (!refreshToken) {
      if (seededFlowIds.has(flow.id) || seedOutcome === 'success') errors.push(`${flow.refreshTokenSecret} is missing from its token store`);
      continue;
    }

    available++;
    try {
      writeSecret(flow.refreshTokenSecret, refreshToken, secretContext);
    } catch (error) {
      errors.push(error instanceof Error ? error.message : `Could not persist ${flow.refreshTokenSecret}`);
    }
  }

  if (available === 0 && seedOutcome !== 'success' && errors.length === 0) {
    console.log('No completed token flows were available; no provider request was retried.');
  } else if (available === 0 && errors.length === 0) {
    errors.push('No token stores were available after a successful seed');
  }
  if (errors.length) throw new Error(`Token finalization failed: ${[...new Set(errors)].join('; ')}. Check the seed result and CI_SECRETS_TOKEN access.`);
}

async function cleanup(): Promise<void> {
  await assertActionsContext();
  if (!(await exists(ownerFile))) {
    console.log('No CI token files are owned by this run.');
    return;
  }

  const paths = [envFile, tokenRoot];
  if (config.serviceAccount) paths.push(path.join(root, config.serviceAccount.output));
  const failures: string[] = [];
  for (const filePath of paths) {
    try {
      await rm(filePath, { recursive: true, force: true, maxRetries: 3, retryDelay: 200 });
    } catch {
      failures.push(path.basename(filePath));
    }
  }
  if (failures.length) throw new Error(`Could not remove private runner files: ${failures.join(', ')}`);
  await rm(ownerFile, { force: true });
}

async function main(): Promise<void> {
  const command = process.argv[2];
  if (command === 'seed') await seed();
  else if (command === 'persist') await persist();
  else if (command === 'cleanup') await cleanup();
  else throw new Error('Usage: node scripts/ci-tokens.ts seed|persist|cleanup');
}

main().catch((error: unknown) => {
  console.error(error instanceof Error ? error.message : 'CI token operation failed');
  process.exitCode = 1;
});
