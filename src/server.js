// Inicialização do servidor web do CLEAN.
import { loadConfig } from './config.js';
import { Store } from './store.js';
import { ScanManager } from './scan/manager.js';
import { createApp } from './app.js';
import { Scheduler } from './schedule/scheduler.js';
import { cleanupTempFiles } from './scan/powershell.js';

const config = loadConfig();
const store = await new Store(config.dataDir).init();
await cleanupTempFiles();
const manager = new ScanManager(store, { maxConcurrent: config.maxConcurrentScans });
const scheduler = new Scheduler({ store, manager });
const app = createApp({ store, manager, config, scheduler });

const server = app.listen(config.port, config.host, () => {
  const host = config.host === '0.0.0.0' || config.host === '::' ? 'localhost' : config.host;
  console.log(`[CLEAN] Servidor iniciado em http://${host}:${config.port}`);
  console.log(`[CLEAN] Dados em ${config.dataDir}`);
  const local = ['127.0.0.1', 'localhost', '::1'].includes(config.host);
  if (!local && !(config.authUser && config.authPassword)) {
    console.warn('[CLEAN] ATENÇÃO: o servidor aceita conexões da rede sem senha. Defina AUTH_USER e AUTH_PASSWORD.');
  }
  // Agendamentos: a primeira verificação executa (ou registra como perdidos) os horários que
  // passaram enquanto o CLEAN estava parado.
  scheduler.start();
});

server.on('error', (err) => {
  console.error(`[CLEAN] Não foi possível iniciar o servidor: ${err.message}`);
  process.exit(1);
});

// Um erro solto não deve derrubar o servidor (as análises rodam em threads separadas).
process.on('unhandledRejection', (reason) => {
  console.error('[CLEAN] Erro não tratado:', reason);
});

let closing = false;
async function shutdown(signal) {
  if (closing) return;
  closing = true;
  console.log(`[CLEAN] Encerrando (${signal})...`);
  server.close();
  await scheduler.stop().catch(() => {});
  await manager.shutdown().catch(() => {});
  process.exit(0);
}
process.on('SIGINT', () => shutdown('SIGINT'));
process.on('SIGTERM', () => shutdown('SIGTERM'));
