import { config } from 'dotenv';
import { resolve } from 'path';

// Load .env.test if it exists
config({ path: resolve(process.cwd(), '.env.test') });
