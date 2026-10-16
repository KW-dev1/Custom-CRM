import { loadConfig, ConfigError } from './config.js';
import { createGoogleAuth, createServiceAccountAuth, loadServiceAccountKey, GoogleAuthError } from './google-auth.js';
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

let auth = null;
if (cfg.source.type === 'google') {
  try {
    auth = cfg.google.serviceAccountFile
      ? createServiceAccountAuth(await loadServiceAccountKey(cfg.google.serviceAccountFile))
      : createGoogleAuth(cfg.google);
  } catch (e) {
    if (e instanceof GoogleAuthError) { console.error(e.message); process.exit(1); }
    throw e;
  }
}
const sheet = cfg.source.type === 'xlsx' ? createXlsxStore({ cfg }) : createSheetsClient({ cfg, auth });
log('info', 'data_source', {
  type: cfg.source.type,
  file: cfg.source.path ? cfg.source.path.split(/[\\/]/).pop() : undefined,
  login: auth ? (cfg.google.serviceAccountFile ? 'service_account' : 'oauth_user') : undefined,
  shareSheetWith: auth && cfg.google.serviceAccountFile ? auth.identity : undefined,
});
const service = createJobService({ sheet, cfg, log });
const app = createApp({ service, cfg, log });

const server = app.listen(cfg.port, cfg.host, () => log('info', 'listening', { host: cfg.host, port: cfg.port }));
for (const sig of ['SIGINT', 'SIGTERM']) process.on(sig, () => server.close(() => process.exit(0)));
