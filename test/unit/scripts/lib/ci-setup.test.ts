import { setToken } from '@mcp-z/oauth';
import assert from 'assert';
import * as fs from 'fs';
import { safeRmSync } from 'fs-remove-compat';
import * as path from 'path';
import { config as repoConfig } from '../../../../scripts/ci-config.ts';
import { type CiConfig, missingRequired, openStore, planSecrets, readLocalValues, upload } from '../../../../scripts/lib/ci-setup.ts';

const config: CiConfig = repoConfig;
const REPOSITORY = 'example-org/example-repo';
const HUMAN_ACCOUNT = 'ci-human@example.com';
const SCOPE = 'dummy-scope https://example.com/dummy';

// The refresh token each secret must carry, written out here rather than derived from the config under test.
const REFRESH_TOKEN_OF = {
  TEST_REFRESH_TOKEN: 'dummy-refresh-human-account',
  TEST_DEVICE_REFRESH_TOKEN: 'dummy-refresh-device-code-account',
  TEST_DCR_REFRESH_TOKEN: 'dummy-refresh-dcr-registration',
} as const;

const plan = planSecrets(config);
const dummy = (name: string) => `dummy-value-for-${name}`;

function neverSet(): never {
  throw new Error('a dry or blocked run must not set secrets');
}

/** Writes dummy local state for every planned secret under `root` and returns the env that supplies the rest. */
async function writeFixture(root: string, options: { accounts?: string[] } = {}): Promise<{ env: Record<string, string>; expected: Map<string, string> }> {
  const expected = new Map<string, string>();
  const env: Record<string, string> = {};
  const accounts = options.accounts ?? [HUMAN_ACCOUNT];

  for (const flow of config.flows) {
    fs.mkdirSync(path.join(root, path.dirname(flow.storage.path)), { recursive: true });
    const store = await openStore(path.join(root, flow.storage.path));
    try {
      if (flow.storage.kind === 'account') {
        const ids = 'env' in flow.storage.accountId ? accounts : [flow.storage.accountId.value];
        for (const accountId of ids) {
          const human = 'env' in flow.storage.accountId;
          const refreshToken = !human ? REFRESH_TOKEN_OF.TEST_DEVICE_REFRESH_TOKEN : accountId === accounts[0] ? REFRESH_TOKEN_OF.TEST_REFRESH_TOKEN : 'dummy-refresh-other-account';
          await setToken(store, { accountId, service: flow.storage.service }, { accessToken: 'dummy-access', refreshToken, expiresAt: 1 });
          if (accountId === accounts[0] || !('env' in flow.storage.accountId)) expected.set(flow.refreshTokenSecret, refreshToken);
        }
      } else {
        const refreshToken = REFRESH_TOKEN_OF.TEST_DCR_REFRESH_TOKEN;
        await store.set(flow.storage.key, { providerRefreshToken: refreshToken, clientId: dummy(flow.storage.registration.clientIdSecret), clientSecret: dummy(flow.storage.registration.clientSecretSecret) });
        expected.set(flow.refreshTokenSecret, refreshToken);
        expected.set(flow.storage.registration.clientIdSecret, dummy(flow.storage.registration.clientIdSecret));
        expected.set(flow.storage.registration.clientSecretSecret, dummy(flow.storage.registration.clientSecretSecret));
      }
    } finally {
      await store.disconnect();
    }
  }

  if (config.serviceAccount) {
    const key: Record<string, string> = { type: 'service_account' };
    for (const [field, name] of Object.entries(config.serviceAccount.secrets)) {
      key[field] = dummy(name);
      expected.set(name, dummy(name));
    }
    fs.writeFileSync(path.join(root, config.serviceAccount.output), JSON.stringify(key));
  }

  for (const { name, source } of plan) {
    if (source.kind === 'env') {
      env[name] = dummy(name);
      expected.set(name, dummy(name));
    } else if (source.kind === 'scope') expected.set(name, SCOPE);
    else if (source.kind === 'account-id') expected.set(name, accounts[0]);
  }
  return { env, expected };
}

describe('ci-setup', () => {
  const scratchRoot = path.resolve('.tmp', `ci-setup-${process.pid}-${Date.now()}`);
  let caseIndex = 0;
  const freshRoot = () => path.join(scratchRoot, `case-${caseIndex++}`);

  after(() => safeRmSync(scratchRoot, { recursive: true, force: true }));

  describe('planSecrets', () => {
    it('names the secrets CONTRIBUTING documents, plus the one it does not upload', () => {
      const contributing = fs.readFileSync(path.resolve('CONTRIBUTING.md'), 'utf8');
      const documented = [...contributing.matchAll(/^- `([A-Z][A-Z0-9_]+)`/gm)].map((match) => match[1]);
      assert.deepStrictEqual([...planSecrets(config).map((secret) => secret.name), 'CI_SECRETS_TOKEN'].sort(), [...new Set(documented)].sort());
    });

    it('lists each secret once and every refresh token after all other secrets', () => {
      const names = plan.map((secret) => secret.name);
      assert.strictEqual(new Set(names).size, names.length);
      const refreshNames = config.flows.map((flow) => flow.refreshTokenSecret);
      assert.deepStrictEqual(names.slice(-refreshNames.length), refreshNames);
    });
  });

  describe('readLocalValues', () => {
    it('reads every secret from the local stores, env and key file, with the scope from the config', async () => {
      const root = freshRoot();
      const { env, expected } = await writeFixture(root);
      const values = await readLocalValues(config, root, env, SCOPE);
      assert.deepStrictEqual([...values.entries()].sort(), [...expected.entries()].sort());
    });

    it("puts each flow's own stored refresh token under its secret name", async () => {
      const root = freshRoot();
      const { env } = await writeFixture(root);
      const values = await readLocalValues(config, root, env, SCOPE);
      for (const flow of config.flows) {
        const expected = REFRESH_TOKEN_OF[flow.refreshTokenSecret as keyof typeof REFRESH_TOKEN_OF];
        assert.ok(expected, `no expected token for ${flow.refreshTokenSecret}`);
        assert.strictEqual(values.get(flow.refreshTokenSecret), expected);
      }
    });

    it('leaves out a secret whose store is missing', async () => {
      const root = freshRoot();
      const { env } = await writeFixture(root);
      const firstFlow = config.flows[0];
      safeRmSync(path.join(root, firstFlow.storage.path), { force: true });
      const values = await readLocalValues(config, root, env, SCOPE);
      assert.strictEqual(values.has(firstFlow.refreshTokenSecret), false);
      assert.ok(missingRequired(plan, values).includes(firstFlow.refreshTokenSecret));
    });

    it('stops with a clear message when the store holds more than one account', async () => {
      const root = freshRoot();
      const { env } = await writeFixture(root, { accounts: [HUMAN_ACCOUNT, 'ci-other@example.com'] });
      await assert.rejects(readLocalValues(config, root, env, SCOPE), (error: Error) => {
        assert.ok(/2 accounts hold a token/.test(error.message));
        assert.ok(!error.message.includes('example.com'));
        return true;
      });
    });
  });

  describe('upload', () => {
    it('--dry reports every secret as present and sets nothing', async () => {
      const root = freshRoot();
      const { env, expected } = await writeFixture(root);
      const result = await upload({ config, root, env, scope: SCOPE, repository: REPOSITORY, dry: true, setSecret: neverSet });
      assert.strictEqual(result.ok, true);
      assert.deepStrictEqual(
        result.lines,
        plan.map((secret) => `${secret.name}: present (dry)`)
      );
      for (const value of expected.values()) assert.ok(!result.lines.some((line) => line.includes(value)));
    });

    it('--dry reports a secret with no local value as missing', async () => {
      const root = freshRoot();
      const { env } = await writeFixture(root);
      const requiredEnv = plan.find((secret) => secret.source.kind === 'env' && secret.required);
      assert.ok(requiredEnv);
      delete env[requiredEnv.name];
      const result = await upload({ config, root, env, scope: SCOPE, repository: REPOSITORY, dry: true, setSecret: neverSet });
      assert.strictEqual(result.ok, false);
      assert.ok(result.lines.includes(`${requiredEnv.name}: missing`));
    });

    it('sets nothing when a required secret is missing', async () => {
      const root = freshRoot();
      const { env } = await writeFixture(root);
      const requiredEnv = plan.find((secret) => secret.source.kind === 'env' && secret.required);
      assert.ok(requiredEnv);
      delete env[requiredEnv.name];
      const result = await upload({ config, root, env, scope: SCOPE, repository: REPOSITORY, dry: false, setSecret: neverSet });
      assert.strictEqual(result.ok, false);
      assert.ok(result.lines.includes(`${requiredEnv.name}: missing`));
    });

    it('sets each secret with its value, refresh tokens last', async () => {
      const root = freshRoot();
      const { env, expected } = await writeFixture(root);
      const calls: Array<[string, string, string]> = [];
      const result = await upload({ config, root, env, scope: SCOPE, repository: REPOSITORY, dry: false, setSecret: (name, value, repository) => calls.push([name, value, repository]) });
      assert.strictEqual(result.ok, true);
      assert.deepStrictEqual(
        calls.map(([name]) => name),
        plan.map((secret) => secret.name)
      );
      for (const [name, value, repository] of calls) {
        assert.strictEqual(value, expected.get(name));
        assert.strictEqual(repository, REPOSITORY);
      }
      assert.deepStrictEqual(
        result.lines,
        plan.map((secret) => `${secret.name}: set`)
      );
      for (const value of expected.values()) assert.ok(!result.lines.some((line) => line.includes(value)));
    });

    it('stops at the first failure without printing its value, so later refresh tokens are not set', async () => {
      const root = freshRoot();
      const { env, expected } = await writeFixture(root);
      const failing = plan[0].name;
      const calls: string[] = [];
      const result = await upload({
        config,
        root,
        env,
        scope: SCOPE,
        repository: REPOSITORY,
        dry: false,
        setSecret: (name, value) => {
          calls.push(name);
          if (name === failing) throw new Error(`gh rejected ${value}`);
        },
      });
      assert.strictEqual(result.ok, false);
      assert.deepStrictEqual(calls, [failing]);
      assert.ok(result.lines[0].startsWith(`${failing}: FAILED`));
      for (const value of expected.values()) assert.ok(!result.lines.some((line) => line.includes(value)));
    });
  });
});
