// Execução de um reinício: lista os aparelhos do escopo e envia a ação de reinício a cada um,
// registrando o resultado (sucesso/falha) por aparelho. Usado pelo agendador e pela ação "executar
// agora".
import { pool } from '../mail/http.js';
import { friendlyError } from '../scan/errors.js';
import { IntuneConnector, APPLE_OS } from './connector.js';

// Teto de aparelhos por execução (proteção) e concorrência padrão das ações de reinício.
export const MAX_DEVICES = 2000;
const DEFAULT_CONCURRENCY = 4;

// O reinício do Intune (rebootNow) é enfileirado e só chega ao aparelho na próxima comunicação dele
// com o Intune. Não há status "online" direto: um aparelho que não sincroniza há mais de algumas
// horas provavelmente está offline — o reinício fica pendente até ele reconectar. Usamos a última
// sincronização como indício de que o aparelho está (ou não) acessível.
export const STALE_SYNC_HOURS = 24;
export const STALE_SYNC_MS = STALE_SYNC_HOURS * 60 * 60 * 1000;

/** Indício de aparelho offline: sem sincronização registrada ou sincronização antiga demais. */
export function deviceStale(lastSync, now = Date.now(), threshold = STALE_SYNC_MS) {
  if (!lastSync) return true;
  const t = Date.parse(lastSync);
  return !Number.isFinite(t) || now - t > threshold;
}

const norm = (v) => String(v || '').trim().toLowerCase();

/** Normaliza o escopo do reinício: todos os iPhone/iPad ('all') ou uma lista ('list'). */
export function sanitizeScope(input) {
  const i = input && typeof input === 'object' && !Array.isArray(input) ? input : {};
  const mode = i.mode === 'list' ? 'list' : 'all';
  const names = (Array.isArray(i.names) ? i.names : String(i.names || '').split(/[\n,;]+/))
    .map((n) => String(n).trim())
    .filter(Boolean)
    .slice(0, 1000);
  if (mode === 'list' && names.length === 0) {
    const err = new Error('Na lista específica, informe ao menos um aparelho (nome do dispositivo, e-mail do usuário ou número de série).');
    err.status = 400;
    throw err;
  }
  return mode === 'list' ? { mode, names } : { mode: 'all', names: [] };
}

/** Função que decide se um aparelho entra no escopo. */
function deviceMatcher(scope) {
  if (scope.mode !== 'list') return () => true;
  const set = new Set(scope.names.map(norm));
  return (d) => set.has(norm(d.name)) || set.has(norm(d.user)) || set.has(norm(d.serialNumber));
}

/** Os aparelhos (iPhone/iPad) do escopo, já filtrados. Lança em caso de falha de acesso/listagem. */
export async function listDevices(connector, scope) {
  const match = deviceMatcher(scope);
  const out = [];
  for await (const device of connector.devices({ os: APPLE_OS })) {
    if (match(device)) out.push(device);
    if (out.length >= MAX_DEVICES) break;
  }
  return out;
}

const summary = (device, status, { error = null, stale = false } = {}) => ({
  id: device.id,
  name: device.name,
  user: device.user,
  userName: device.userName,
  model: device.model,
  os: device.os,
  osVersion: device.osVersion,
  serialNumber: device.serialNumber,
  lastSync: device.lastSync,
  // Provável offline: o reinício foi aceito pelo Intune, mas só chega ao aparelho quando ele voltar
  // a sincronizar.
  stale,
  status, // 'ok' (reinício enviado) | 'failed' | 'skipped'
  error,
});

/**
 * Reinicia os aparelhos do escopo. Entrega { startedAt, finishedAt, total, ok, failed, skipped,
 * stale, devices: [...] }. "stale" conta os aparelhos provavelmente offline (sem sincronizar com o
 * Intune há mais de STALE_SYNC_HOURS): o reinício fica pendente até eles reconectarem. simulate: só
 * lista os aparelhos (não envia o reinício). signal: cancela.
 */
export async function runReboot(connector, scope, { concurrency = DEFAULT_CONCURRENCY, log = () => {}, signal = null, simulate = false, staleMs = STALE_SYNC_MS } = {}) {
  const startedAt = new Date().toISOString();
  const nowMs = Date.parse(startedAt);
  const devices = await listDevices(connector, scope);
  log('info', `${devices.length} aparelho(s) no escopo${simulate ? ' (simulação: nada é reiniciado)' : ''}.`);
  const results = [];
  let ok = 0;
  let failed = 0;
  let skipped = 0;
  let stale = 0;
  const limit = Math.min(Math.max(1, Number(concurrency) || DEFAULT_CONCURRENCY), 8);
  const worker = async (device) => {
    const offline = deviceStale(device.lastSync, nowMs, staleMs);
    if (offline) stale++;
    const mk = (status, error = null) => summary(device, status, { error, stale: offline });
    if (signal?.aborted) {
      skipped++;
      results.push(mk('skipped', 'Execução cancelada.'));
      return undefined;
    }
    if (simulate) {
      skipped++;
      results.push(mk('skipped'));
      return undefined;
    }
    try {
      await connector.reboot(device.id, { signal });
      ok++;
      results.push(mk('ok'));
      if (offline) log('warn', `"${device.name}": reinício enviado, mas o aparelho não sincroniza com o Intune desde ${device.lastSync || 'nunca'} — provavelmente offline; o reinício só ocorre quando ele reconectar.`);
    } catch (err) {
      failed++;
      const message = friendlyError(err);
      results.push(mk('failed', message));
      log('warn', `Falha ao reiniciar "${device.name}": ${message}`);
    }
    return undefined;
  };
  for await (const _ of pool(devices, limit, worker)); // eslint-disable-line no-unused-vars
  results.sort((a, b) => a.name.localeCompare(b.name, 'pt-BR'));
  return { startedAt, finishedAt: new Date().toISOString(), total: devices.length, ok, failed, skipped, stale, devices: results };
}

/** Monta o conector do Intune a partir de uma conexão Microsoft 365 (Graph) com os segredos decifrados. */
export function intuneConnectorFor(source, options = {}) {
  return new IntuneConnector(source, options);
}
