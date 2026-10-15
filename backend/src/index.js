import { loadConfig, ConfigError } from './config.js';
import { createGoogleAuth } from './google-auth.js';
import { createSheetsClient } from './sheets-client.js';
import { createXlsxStore } from './xlsx-store.js';
import { createJobService } from './jobs.js';
import { createApp } from './app.js';
import { createLogger } from './logger.js';

const log = createLogger();
let cfg;
try {
  cfg = loadConfig();
} catch (e) {
  if (e instanceof ConfigError) { console.error(e.message); process.exit(1); }
  throw e;
}

const sheet = cfg.source.type === 'xlsx'
  ? createXlsxStore({ cfg })
  : createSheetsClient({ cfg, auth: createGoogleAuth(cfg.google) });
log('info', 'data_source', { type: cfg.source.type, file: cfg.source.path ? cfg.source.path.split(/[\\/]/).pop() : undefined });
const service = createJobService({ sheet, cfg, log });
const app = createApp({ service, cfg, log });

const server = app.listen(cfg.port, cfg.host, () => log('info', 'listening', { host: cfg.host, port: cfg.port }));
for (const sig of ['SIGINT', 'SIGTERM']) process.on(sig, () => server.close(() => process.exit(0)));
