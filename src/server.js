// Inicialização do servidor web do CLEAN.
import { loadConfig } from './config.js';
import { Store } from './store.js';
import { ScanManager } from './scan/manager.js';
import { createApp } from './app.js';

const config = loadConfig();
const store = await new Store(config.dataDir).init();
const manager = new ScanManager(store, { maxConcurrent: config.maxConcurrentScans });
const app = createApp({ store, manager, config });

const server = app.listen(config.port, config.host, () => {
  const host = config.host === '0.0.0.0' || config.host === '::' ? 'localhost' : config.host;
  console.log(`[CLEAN] Servidor iniciado em http://${host}:${config.port}`);
  console.log(`[CLEAN] Dados em ${config.dataDir}`);
  const local = ['127.0.0.1', 'localhost', '::1'].includes(config.host);
  if (!local && !(config.authUser && config.authPassword)) {
    console.warn('[CLEAN] ATENÇÃO: o servidor aceita conexões da rede sem senha. Defina AUTH_USER e AUTH_PASSWORD.');
  }
});

server.on('error', (err) => {
  console.error(`[CLEAN] Não foi possível iniciar o servidor: ${err.message}`);
  process.exit(1);
});

let closing = false;
async function shutdown(signal) {
  if (closing) return;
  closing = true;
  console.log(`[CLEAN] Encerrando (${signal})...`);
  server.close();
  await manager.shutdown().catch(() => {});
  process.exit(0);
}
process.on('SIGINT', () => shutdown('SIGINT'));
process.on('SIGTERM', () => shutdown('SIGTERM'));
