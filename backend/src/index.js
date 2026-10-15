import { loadConfig, ConfigError } from './config.js';
import { createGoogleAuth } from './google-auth.js';
import { createSheetsClient } from './sheets-client.js';
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

const auth = createGoogleAuth(cfg.google);
const sheet = createSheetsClient({ cfg, auth });
const service = createJobService({ sheet, cfg, log });
const app = createApp({ service, cfg, log });

const server = app.listen(cfg.port, cfg.host, () => log('info', 'listening', { host: cfg.host, port: cfg.port }));
for (const sig of ['SIGINT', 'SIGTERM']) process.on(sig, () => server.close(() => process.exit(0)));
