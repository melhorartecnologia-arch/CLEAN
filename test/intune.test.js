// Reinício agendado de dispositivos no Intune (Microsoft Graph): conector (listar iPhone/iPad e
// rebootNow), motor do reinício (status por aparelho), agendador (recorrência, horários perdidos) e
// API /api/intune, com o Graph simulado.
import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { Store } from '../src/store.js';
import { ScanManager } from '../src/scan/manager.js';
import { createApp } from '../src/app.js';
import { IntuneScheduler, jobProblems } from '../src/intune/scheduler.js';
import { IntuneConnector } from '../src/intune/connector.js';
import { runReboot, sanitizeScope, listDevices, deviceStale } from '../src/intune/reboot.js';
import { validateRule } from '../src/schedule/recurrence.js';
import { startMockApis } from './helpers/mock-apis.js';

const GRAPH_TENANT = 'contoso.onmicrosoft.com';
const GRAPH_CLIENT = '11111111-2222-3333-4444-555555555555';
const GRAPH_SECRET = 'segredo-super-secreto';

const hoursAgo = (h) => new Date(Date.now() - h * 60 * 60 * 1000).toISOString();

function graphData(extra = {}) {
  return {
    tenant: GRAPH_TENANT,
    clientId: GRAPH_CLIENT,
    secret: GRAPH_SECRET,
    users: [],
    // d-iphone2 não sincroniza há 48h (provável offline); os demais sincronizaram há pouco.
    intuneDevices: [
      { id: 'd-iphone1', deviceName: 'iPhone de João', operatingSystem: 'iOS', osVersion: '17.5', model: 'iPhone 14', serialNumber: 'SER-IPHONE-1', userPrincipalName: 'joao@contoso.com', userDisplayName: 'João', lastSyncDateTime: hoursAgo(1) },
      { id: 'd-ipad1', deviceName: 'iPad RH', operatingSystem: 'iPadOS', osVersion: '17.4', model: 'iPad Air', serialNumber: 'SER-IPAD-1', userPrincipalName: 'rh@contoso.com', userDisplayName: 'RH', lastSyncDateTime: hoursAgo(3) },
      { id: 'd-iphone2', deviceName: 'iPhone de Maria', operatingSystem: 'iOS', osVersion: '16.0', model: 'iPhone 12', serialNumber: 'SER-IPHONE-2', userPrincipalName: 'maria@contoso.com', userDisplayName: 'Maria', lastSyncDateTime: hoursAgo(48) },
      { id: 'd-android', deviceName: 'Galaxy', operatingSystem: 'Android', osVersion: '14', model: 'S22', serialNumber: 'SER-AND', userPrincipalName: 'x@contoso.com', lastSyncDateTime: hoursAgo(1) },
      { id: 'd-win', deviceName: 'Notebook', operatingSystem: 'Windows', osVersion: '11', serialNumber: 'SER-WIN', lastSyncDateTime: hoursAgo(1) },
    ],
    ...extra,
  };
}

const RULE = validateRule({ frequency: 'daily', interval: 1, startDate: '2026-10-16', time: '02:00', end: 'never' });

let root;
before(() => {
  root = fs.mkdtempSync(path.join(os.tmpdir(), 'clean-intune-'));
});
after(() => {
  fs.rmSync(root, { recursive: true, force: true });
});

/** Store com uma conexão Microsoft 365 (Graph) cujos segredos ficam cifrados. */
async function freshStore() {
  const store = await new Store(path.join(root, `data-${Math.random().toString(36).slice(2)}`)).init();
  const source = store.createMailSource({
    name: 'Microsoft 365',
    type: 'graph',
    graph: { tenantId: GRAPH_TENANT, clientId: GRAPH_CLIENT },
    secrets: { clientSecret: store.secrets.seal(GRAPH_SECRET) },
  });
  return { store, source };
}

const connectorFor = (store, source, endpoints) => new IntuneConnector({ ...source, secrets: store.openMailSecrets(source) }, { endpoints });

// ---------------------------------------------------------------------------------------------
// Conector e motor do reinício

test('escopo: lista vazia é recusada; lista normaliza nomes/e-mails/série', () => {
  assert.deepEqual(sanitizeScope({ mode: 'all' }), { mode: 'all', names: [] });
  assert.deepEqual(sanitizeScope({ mode: 'list', names: ['A', ' b ', ''] }), { mode: 'list', names: ['A', 'b'] });
  assert.throws(() => sanitizeScope({ mode: 'list', names: [] }), /ao menos um aparelho/);
});

test('offline: sem sincronização ou sincronização antiga demais conta como provável offline', () => {
  const now = Date.parse('2026-10-16T12:00:00Z');
  assert.equal(deviceStale(null, now), true, 'nunca sincronizou');
  assert.equal(deviceStale('2026-10-16T11:00:00Z', now), false, '1h atrás: acessível');
  assert.equal(deviceStale('2026-10-14T00:00:00Z', now), true, '60h atrás: provável offline');
});

test('conector: lista só iPhone/iPad (ignora Android e Windows), paginando', async () => {
  const mock = await startMockApis({ graph: graphData({ intunePageSize: 2 }) });
  const { store, source } = await freshStore();
  try {
    const connector = connectorFor(store, source, mock.endpoints);
    const devices = await listDevices(connector, { mode: 'all', names: [] });
    assert.deepEqual(devices.map((d) => d.id).sort(), ['d-ipad1', 'd-iphone1', 'd-iphone2']);
    const joao = devices.find((d) => d.id === 'd-iphone1');
    assert.equal(joao.name, 'iPhone de João');
    assert.equal(joao.user, 'joao@contoso.com');
    assert.equal(joao.os, 'iOS');
    assert.equal(joao.serialNumber, 'SER-IPHONE-1');
    await connector.close();
  } finally {
    await mock.close();
  }
});

test('motor: reinicia todos os aparelhos e apura sucesso por aparelho', async () => {
  const data = graphData();
  const mock = await startMockApis({ graph: data });
  const { store, source } = await freshStore();
  try {
    const result = await runReboot(connectorFor(store, source, mock.endpoints), { mode: 'all', names: [] });
    assert.equal(result.total, 3);
    assert.equal(result.ok, 3);
    assert.equal(result.failed, 0);
    assert.ok(result.devices.every((d) => d.status === 'ok'));
    assert.deepEqual((data.rebooted || []).sort(), ['d-ipad1', 'd-iphone1', 'd-iphone2']);
    // Offline: o iPhone de Maria não sincroniza há 48h — o reinício foi enviado, mas fica pendente.
    assert.equal(result.stale, 1);
    assert.equal(result.devices.find((d) => d.id === 'd-iphone2').stale, true);
    assert.ok(result.devices.filter((d) => d.id !== 'd-iphone2').every((d) => d.stale === false));
  } finally {
    await mock.close();
  }
});

test('motor: escopo por lista e falha de um aparelho vira "partial"', async () => {
  const data = graphData({ rebootErrors: new Set(['d-iphone2']) });
  const mock = await startMockApis({ graph: data });
  const { store, source } = await freshStore();
  try {
    const result = await runReboot(connectorFor(store, source, mock.endpoints), { mode: 'list', names: ['iPhone de João', 'SER-IPHONE-2'] });
    assert.equal(result.total, 2, 'só os dois da lista (por nome e por número de série)');
    assert.equal(result.ok, 1);
    assert.equal(result.failed, 1);
    const maria = result.devices.find((d) => d.id === 'd-iphone2');
    assert.equal(maria.status, 'failed');
    assert.ok(maria.error, 'a falha fica registrada no aparelho');
    assert.ok(!(data.rebooted || []).includes('d-iphone2'));
  } finally {
    await mock.close();
  }
});

test('motor: simulação só lista (nada é reiniciado)', async () => {
  const data = graphData();
  const mock = await startMockApis({ graph: data });
  const { store, source } = await freshStore();
  try {
    const result = await runReboot(connectorFor(store, source, mock.endpoints), { mode: 'all', names: [] }, { simulate: true });
    assert.equal(result.total, 3);
    assert.equal(result.skipped, 3);
    assert.equal(result.ok, 0);
    assert.deepEqual(data.rebooted, undefined, 'nenhum rebootNow enviado');
  } finally {
    await mock.close();
  }
});

test('conector: permissão insuficiente (403) vira mensagem com a providência', async () => {
  const mock = await startMockApis({ graph: graphData({ intuneForbidden: true }) });
  const { store, source } = await freshStore();
  try {
    await assert.rejects(listDevices(connectorFor(store, source, mock.endpoints), { mode: 'all', names: [] }), /DeviceManagementManagedDevices/);
  } finally {
    await mock.close();
  }
});

// ---------------------------------------------------------------------------------------------
// Agendador (recorrência)

function makeScheduler(store, endpoints, now) {
  return new IntuneScheduler({ store, endpoints, now: () => now, tickMs: 10 ** 9, graceMs: 5 * 60000 });
}

test('agendador: horário chegado dispara o reinício e registra no histórico', async () => {
  const data = graphData();
  const mock = await startMockApis({ graph: data });
  const { store, source } = await freshStore();
  try {
    const now = new Date('2026-10-16T02:00:00');
    const scheduler = makeScheduler(store, mock.endpoints, now);
    const job = store.createIntuneJob({ name: 'Madrugada', sourceId: source.id, sourceName: source.name, scope: { mode: 'all', names: [] }, rule: RULE, enabled: true, catchUp: true, runCount: 0, countDone: 0, history: [], nextRunAt: now.toISOString() });
    await scheduler.tick();
    const saved = store.getIntuneJob(job.id);
    assert.equal(saved.history.length, 1);
    assert.equal(saved.history[0].status, 'ok');
    assert.equal(saved.history[0].ok, 3);
    assert.equal(saved.runCount, 1);
    assert.equal((data.rebooted || []).length, 3);
    // A próxima execução foi reprogramada para o dia seguinte.
    assert.ok(new Date(saved.nextRunAt) > now);
  } finally {
    await mock.close();
  }
});

test('agendador: horário perdido sem "executar depois" fica só registrado', async () => {
  const data = graphData();
  const mock = await startMockApis({ graph: data });
  const { store, source } = await freshStore();
  try {
    const due = new Date('2026-10-16T02:00:00');
    const now = new Date('2026-10-16T10:00:00'); // muito depois da tolerância
    const scheduler = makeScheduler(store, mock.endpoints, now);
    const job = store.createIntuneJob({ name: 'Madrugada', sourceId: source.id, sourceName: source.name, scope: { mode: 'all', names: [] }, rule: RULE, enabled: true, catchUp: false, runCount: 0, countDone: 0, history: [], nextRunAt: due.toISOString() });
    await scheduler.tick();
    const saved = store.getIntuneJob(job.id);
    assert.equal(saved.history[0].status, 'missed');
    assert.equal(saved.runCount, 0, 'um horário perdido não conta como execução');
    assert.equal(data.rebooted, undefined);
  } finally {
    await mock.close();
  }
});

test('agendador: horário perdido com "executar depois" dispara uma vez (catch-up)', async () => {
  const data = graphData();
  const mock = await startMockApis({ graph: data });
  const { store, source } = await freshStore();
  try {
    const due = new Date('2026-10-16T02:00:00');
    const now = new Date('2026-10-16T10:00:00');
    const scheduler = makeScheduler(store, mock.endpoints, now);
    const job = store.createIntuneJob({ name: 'Madrugada', sourceId: source.id, sourceName: source.name, scope: { mode: 'all', names: [] }, rule: RULE, enabled: true, catchUp: true, runCount: 0, countDone: 0, history: [], nextRunAt: due.toISOString() });
    await scheduler.tick();
    const saved = store.getIntuneJob(job.id);
    assert.equal(saved.history[0].status, 'ok');
    assert.equal(saved.history[0].trigger, 'catch-up');
    assert.equal((data.rebooted || []).length, 3);
  } finally {
    await mock.close();
  }
});

test('agendador: conexão excluída é um problema e a execução no horário falha sem reiniciar', async () => {
  const data = graphData();
  const mock = await startMockApis({ graph: data });
  const { store, source } = await freshStore();
  try {
    const now = new Date('2026-10-16T02:00:00');
    const scheduler = makeScheduler(store, mock.endpoints, now);
    const job = store.createIntuneJob({ name: 'Órfão', sourceId: source.id, sourceName: source.name, scope: { mode: 'all', names: [] }, rule: RULE, enabled: true, catchUp: true, runCount: 0, countDone: 0, history: [], nextRunAt: now.toISOString() });
    store.deleteMailSource(source.id);
    assert.ok(jobProblems(store, store.getIntuneJob(job.id)).some((p) => /excluída/.test(p)));
    await scheduler.tick();
    const saved = store.getIntuneJob(job.id);
    assert.equal(saved.history[0].status, 'failed');
    assert.equal(data.rebooted, undefined);
  } finally {
    await mock.close();
  }
});

// ---------------------------------------------------------------------------------------------
// API /api/intune

async function startApp(data) {
  const mock = await startMockApis({ graph: data });
  const store = await new Store(path.join(root, `app-${Math.random().toString(36).slice(2)}`)).init();
  const manager = new ScanManager(store, { mailEndpoints: mock.endpoints });
  const app = createApp({ store, manager, config: { authUser: '', authPassword: '', mailEndpoints: mock.endpoints } });
  const srv = await new Promise((resolve) => {
    const x = app.listen(0, '127.0.0.1', () => resolve(x));
  });
  const base = `http://127.0.0.1:${srv.address().port}`;
  const api = async (method, url, body) => {
    const res = await fetch(base + url, { method, headers: { 'Content-Type': 'application/json', 'X-CLEAN': '1' }, body: body === undefined ? undefined : JSON.stringify(body) });
    const type = res.headers.get('content-type') || '';
    return { status: res.status, data: type.includes('json') ? await res.json() : Buffer.from(await res.arrayBuffer()) };
  };
  return { store, mock, api, close: async () => (srv.close(), await manager.shutdown(), await mock.close()) };
}

async function graphSource(api) {
  const res = await api('POST', '/api/mail-sources', { name: 'Microsoft 365', type: 'graph', scope: 'all', graph: { tenantId: GRAPH_TENANT, clientId: GRAPH_CLIENT, clientSecret: GRAPH_SECRET } });
  assert.equal(res.status, 201, JSON.stringify(res.data));
  return res.data;
}

const dailyRule = { frequency: 'daily', interval: 1, startDate: new Date().toISOString().slice(0, 10), time: '02:00', end: 'never' };

test('API: conexões disponíveis, teste de acesso e listagem de aparelhos', async () => {
  const app = await startApp(graphData());
  try {
    const source = await graphSource(app.api);
    const sources = await app.api('GET', '/api/intune/sources');
    assert.deepEqual(sources.data.map((s) => s.id), [source.id]);

    const test = await app.api('POST', '/api/intune/test', { sourceId: source.id });
    assert.equal(test.data.ok, true);
    assert.match(test.data.message, /iPhone\/iPad/);

    const devices = await app.api('GET', `/api/intune/devices?sourceId=${source.id}`);
    assert.equal(devices.data.total, 3);
    assert.deepEqual(devices.data.devices.map((d) => d.id).sort(), ['d-ipad1', 'd-iphone1', 'd-iphone2']);
    // A listagem já marca o provável offline (última sincronização antiga).
    assert.equal(devices.data.devices.find((d) => d.id === 'd-iphone2').stale, true);
    assert.equal(devices.data.devices.find((d) => d.id === 'd-iphone1').stale, false);
  } finally {
    await app.close();
  }
});

test('API: criar, simular, reiniciar agora e ver o histórico', async () => {
  const app = await startApp(graphData());
  try {
    const source = await graphSource(app.api);
    const created = await app.api('POST', '/api/intune', { name: 'Madrugada', sourceId: source.id, scope: { mode: 'all', names: [] }, rule: dailyRule });
    assert.equal(created.status, 201, JSON.stringify(created.data));
    assert.equal(created.data.state, 'active');
    assert.ok(created.data.nextRunAt);
    const id = created.data.id;

    // Simular: lista os aparelhos, nada é reiniciado.
    const sim = await app.api('POST', `/api/intune/${id}/run`, { simulate: true });
    assert.equal(sim.status, 201);
    assert.equal(sim.data.run.status, 'ok');
    assert.equal(sim.data.run.skipped, 3);
    assert.equal(app.mock.calls.filter((c) => c.includes('rebootNow')).length, 0);

    // Reinício real exige confirmação.
    const semConfirm = await app.api('POST', `/api/intune/${id}/run`, {});
    assert.equal(semConfirm.status, 400);

    const run = await app.api('POST', `/api/intune/${id}/run`, { confirm: true });
    assert.equal(run.status, 201);
    assert.equal(run.data.run.status, 'ok');
    assert.equal(run.data.run.ok, 3);
    assert.equal(app.mock.calls.filter((c) => c.includes('rebootNow')).length, 3);

    const full = await app.api('GET', `/api/intune/${id}`);
    assert.equal(full.data.history.length, 2, 'simulação + reinício');
    assert.equal(full.data.history[0].status, 'ok');
  } finally {
    await app.close();
  }
});

test('API: pausar e retomar', async () => {
  const app = await startApp(graphData());
  try {
    const source = await graphSource(app.api);
    const created = await app.api('POST', '/api/intune', { name: 'Madrugada', sourceId: source.id, scope: { mode: 'all', names: [] }, rule: dailyRule });
    const id = created.data.id;
    const paused = await app.api('POST', `/api/intune/${id}/pause`);
    assert.equal(paused.data.state, 'paused');
    assert.equal(paused.data.nextRunAt, null);
    const resumed = await app.api('POST', `/api/intune/${id}/resume`);
    assert.equal(resumed.data.state, 'active');
    assert.ok(resumed.data.nextRunAt);
  } finally {
    await app.close();
  }
});

test('API: validação (lista vazia, conexão inexistente) e exclusão', async () => {
  const app = await startApp(graphData());
  try {
    const source = await graphSource(app.api);
    const empty = await app.api('POST', '/api/intune', { name: 'X', sourceId: source.id, scope: { mode: 'list', names: [] }, rule: dailyRule });
    assert.equal(empty.status, 400);
    const noSource = await app.api('POST', '/api/intune', { name: 'X', sourceId: 'não-existe', scope: { mode: 'all', names: [] }, rule: dailyRule });
    assert.equal(noSource.status, 400);

    const created = await app.api('POST', '/api/intune', { name: 'Para excluir', sourceId: source.id, scope: { mode: 'all', names: [] }, rule: dailyRule });
    const del = await app.api('DELETE', `/api/intune/${created.data.id}`);
    assert.equal(del.status, 204);
    const gone = await app.api('GET', `/api/intune/${created.data.id}`);
    assert.equal(gone.status, 404);
  } finally {
    await app.close();
  }
});
