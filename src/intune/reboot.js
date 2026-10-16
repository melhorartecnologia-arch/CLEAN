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

const summary = (device, status, { error = null, stale = false, lastSync = device.lastSync, checked = false, online = null } = {}) => ({
  id: device.id,
  name: device.name,
  user: device.user,
  userName: device.userName,
  model: device.model,
  os: device.os,
  osVersion: device.osVersion,
  // Último horário online conhecido (o mais recente obtido na verificação em tempo real).
  lastSync,
  // Verificação em tempo real antes de reiniciar: checked (foi possível consultar o Intune agora),
  // online (true/false; null quando não verificado).
  checked,
  online,
  serialNumber: device.serialNumber,
  // Provável offline: o reinício foi aceito pelo Intune, mas só chega ao aparelho quando ele voltar
  // a sincronizar.
  stale,
  status, // 'ok' (reinício enviado) | 'failed' | 'skipped'
  error,
});

/** Data/hora legível de uma sincronização (ou "nunca"). */
const syncText = (iso) => {
  if (!iso) return 'nunca';
  const d = new Date(iso);
  return Number.isNaN(+d) ? String(iso) : d.toLocaleString('pt-BR');
};

/**
 * Verificação em tempo real do estado do aparelho, logo antes de reiniciar: relê o dispositivo no
 * Intune (a última sincronização atual) e decide se ele está provavelmente online. Se não for
 * possível consultar agora, o reinício é enviado assim mesmo, guardando o último horário online
 * conhecido. O Intune não expõe um status "online" instantâneo para iOS: a sincronização mais
 * recente é o melhor indício disponível.
 */
async function verifyOnline(connector, device, nowMs, staleMs, signal) {
  if (typeof connector.deviceState !== 'function') {
    return { verified: false, online: null, lastSync: device.lastSync, error: 'releitura não suportada por esta conexão' };
  }
  try {
    const fresh = await connector.deviceState(device.id, { signal });
    const lastSync = fresh.lastSync || device.lastSync;
    return { verified: true, online: !deviceStale(lastSync, nowMs, staleMs), lastSync };
  } catch (err) {
    return { verified: false, online: null, lastSync: device.lastSync, error: friendlyError(err) };
  }
}

/**
 * Reinicia os aparelhos do escopo. Entrega { startedAt, finishedAt, total, ok, failed, skipped,
 * stale, devices: [...] }. Antes de cada reinício, faz uma verificação em tempo real (verify) do
 * estado do aparelho no Intune e registra o resultado no log; se o aparelho estiver (ou parecer)
 * offline, o reinício é enviado assim mesmo (fica pendente até ele reconectar). "stale" conta os
 * aparelhos provavelmente offline. simulate: só verifica e lista (não reinicia). signal: cancela.
 */
export async function runReboot(connector, scope, { concurrency = DEFAULT_CONCURRENCY, log = () => {}, signal = null, simulate = false, staleMs = STALE_SYNC_MS, verify = true } = {}) {
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

  /** Faz (e registra) a verificação em tempo real e devolve { lastSync, checked, online, offline }. */
  const check = async (device) => {
    if (!verify) {
      const offline = deviceStale(device.lastSync, nowMs, staleMs);
      return { lastSync: device.lastSync, checked: false, online: null, offline };
    }
    const r = await verifyOnline(connector, device, nowMs, staleMs, signal);
    const offline = r.online === null ? deviceStale(r.lastSync, nowMs, staleMs) : !r.online;
    const tail = simulate ? '' : ' Enviando o reinício.';
    if (!r.verified) {
      log('warn', `"${device.name}": não foi possível verificar o estado em tempo real (${r.error}). Último horário online conhecido: ${syncText(r.lastSync)}.${simulate ? '' : ' Enviando o reinício mesmo assim.'}`);
    } else if (r.online) {
      log('info', `"${device.name}": verificação em tempo real — última sincronização com o Intune em ${syncText(r.lastSync)} (provavelmente online).${tail}`);
    } else {
      log('warn', `"${device.name}": verificação em tempo real — sem sincronizar desde ${syncText(r.lastSync)} (provavelmente offline). O reinício será entregue quando o aparelho reconectar.${simulate ? '' : ' Enviando mesmo assim.'}`);
    }
    return { lastSync: r.lastSync, checked: r.verified, online: r.online, offline };
  };

  const worker = async (device) => {
    if (signal?.aborted) {
      skipped++;
      const offline = deviceStale(device.lastSync, nowMs, staleMs);
      if (offline) stale++;
      results.push(summary(device, 'skipped', { error: 'Execução cancelada.', stale: offline }));
      return undefined;
    }
    const c = await check(device);
    if (c.offline) stale++;
    const fields = { stale: c.offline, lastSync: c.lastSync, checked: c.checked, online: c.online };
    if (simulate) {
      skipped++;
      results.push(summary(device, 'skipped', fields));
      return undefined;
    }
    try {
      await connector.reboot(device.id, { signal });
      ok++;
      results.push(summary(device, 'ok', fields));
    } catch (err) {
      failed++;
      const message = friendlyError(err);
      results.push(summary(device, 'failed', { ...fields, error: message }));
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
