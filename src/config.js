// Configuração por variáveis de ambiente (ou arquivo .env na pasta do projeto).
import fs from 'node:fs';
import path from 'node:path';

export const PROJECT_ROOT = path.resolve(import.meta.dirname, '..');

export function loadConfig(env = process.env) {
  const envFile = path.join(PROJECT_ROOT, '.env');
  if (env === process.env && fs.existsSync(envFile) && typeof process.loadEnvFile === 'function') {
    process.loadEnvFile(envFile);
  }
  const int = (value, fallback) => {
    const n = Number.parseInt(value, 10);
    return Number.isFinite(n) && n > 0 ? n : fallback;
  };
  return {
    port: int(env.PORT, 3000),
    host: env.HOST || '127.0.0.1',
    dataDir: path.resolve(PROJECT_ROOT, env.DATA_DIR || 'data'),
    authUser: env.AUTH_USER || '',
    authPassword: env.AUTH_PASSWORD || '',
    maxConcurrentScans: int(env.MAX_CONCURRENT_SCANS, 1),
    // Nomes pelos quais o CLEAN pode ser acessado, além de localhost e do nome/IPs da máquina
    // (ex.: um apelido DNS ou o endereço publicado por um proxy). "*" desativa a verificação.
    allowedHosts: String(env.ALLOWED_HOSTS || '')
      .split(',')
      .map((h) => h.trim().toLowerCase())
      .filter(Boolean),
  };
}
