import type { CiConfig } from './lib/ci-setup.ts';

export const config = {
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
} as const satisfies CiConfig;
