import { serverPrivateEnvPath } from '@mcp-z/server';
import { loadEnv } from 'portable-env/node';

if (loadEnv('.env.test').error) loadEnv(serverPrivateEnvPath({ name: 'mcp-sheets', fileName: '.env.test' }));
