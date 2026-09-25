import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { Store } from '../src/store.js';
import { ScanManager } from '../src/scan/manager.js';
import { createApp } from '../src/app.js';
import { Scheduler, INCREMENTAL_MARGIN_MS } from '../src/schedule/scheduler.js';
import { validateRule } from '../src/schedule/recurrence.js';
import { deletionPins } from '../src/schedule/scheduler.js';

let root;
const stores = [];
const at = (y, m, d, h = 0, min = 0, s = 0) => new Date(y, m - 1, d, h, min, s);

before(() => {
  root = fs.mkdtempSync(path.join(os.tmpdir(), 'clean-sched-'));
});

after(async () => {
  await Promise.all(stores.map((s) => s.close()));
  fs.rmSync(root, { recursive: true, force: true });
});

async function newStore() {
  const store = await new Store(path.join(root, `data-${Math.random().toString(36).slice(2)}`)).init();
  stores.push(store);
  return store;
}

/** Gerenciador simulado: registra os pedidos e deixa o teste decidir quando cada análise termina. */
function stubManager(store) {
  const active = new Set();
  const calls = [];
  return {
    calls,
    active,
    itemDeletions: new Set(),
    hasItemDeletion: () => false,
    isActive: (id) => active.has(id),
    async start(body, { by, schedule }) {
      const scan = store.createScan({ kind: body.kind, name: body.name, status: 'queued', options: body.options, startedBy: by, scheduleId: schedule.id, scheduleName: schedule.name, stats: {} });
      calls.push({ body, by, schedule, scanId: scan.id });
      active.add(scan.id);
      return scan;
    },
    finish(id, { status = 'completed', startedAt, errors = 0 } = {}) {
      active.delete(id);
      store.updateScan(id, { status, startedAt: startedAt.toISOString(), finishedAt: startedAt.toISOString(), stats: { filesMatched: 2, errors } });
    },
  };
}

async function world({ period = { type: 'all' }, rule, catchUp = true, keepLast = 0, action = 'analyze' } = {}) {
  const store = await newStore();
  const dir = path.join(root, `repo-${Math.random().toString(36).slice(2)}`);
  fs.mkdirSync(dir);
  const repo = store.createRepository({ type: 'local', name: 'Arquivos', path: dir, exclude: [], allowDelete: action === 'delete' });
  const list = store.createList({ name: 'Sensíveis', terms: [{ id: 't1', type: 'text', value: 'salário' }] });
  const manager = stubManager(store);
  const clock = { now: at(2026, 9, 25, 10) };
  const scheduler = new Scheduler({ store, manager, now: () => new Date(clock.now) });
  const schedule = store.createSchedule({
    name: 'Varredura noturna',
    kind: 'files',
    enabled: true,
    targetIds: [repo.id],
    listIds: [list.id],
    options: { checkName: true, checkContent: true, nameTarget: 'file', deleteMatches: false },
    action,
    rule: validateRule(rule || { frequency: 'daily', startDate: '2026-09-25', time: '02:00' }),
    period,
    catchUp,
    keepLast,
    names: { [repo.id]: repo.name, [list.id]: list.name },
    deleteConfirmation: action === 'delete' ? { by: 'acesso local', at: clock.now.toISOString(), targets: deletionPins('files', [repo]) } : null,
    history: [],
    runCount: 0,
  });
  store.updateScheduleState(schedule.id, { nextRunAt: scheduler.plan(schedule) });
  const run = async (date) => {
    clock.now = date;
    await scheduler.tick();
    return store.getSchedule(schedule.id);
  };
  return { store, repo, list, manager, clock, scheduler, schedule, run };
}

test('execução no horário, próxima execução e registro no histórico', async () => {
  const { store, manager, schedule, run } = await world();
  assert.equal(schedule.nextRunAt, at(2026, 9, 26, 2).toISOString());
  let s = await run(at(2026, 9, 26, 1, 59, 50));
  assert.equal(manager.calls.length, 0, 'antes do horário nada acontece');
  s = await run(at(2026, 9, 26, 2, 0, 10));
  assert.equal(manager.calls.length, 1);
  const call = manager.calls[0];
  assert.equal(call.body.repositoryIds[0], schedule.targetIds[0]);
  assert.equal(call.body.options.deleteMatches, false);
  assert.equal(call.body.options.modifiedAfter, null);
  assert.equal(call.body.confirmDelete, '');
  assert.match(call.body.name, /^Varredura noturna – 26\/09\/2026/);
  assert.equal(call.by, 'agendamento "Varredura noturna"');
  assert.deepEqual(call.schedule, { id: schedule.id, name: 'Varredura noturna' });
  assert.equal(s.nextRunAt, at(2026, 9, 27, 2).toISOString());
  assert.equal(s.runCount, 1);
  assert.equal(s.history[0].status, 'started');
  assert.equal(s.history[0].trigger, 'schedule');
  assert.equal(s.history[0].plannedFor, at(2026, 9, 26, 2).toISOString());
  const scanId = s.history[0].scanId;
  assert.equal(store.getScan(scanId).scheduleId, schedule.id);

  // Mesmo horário de novo (outra verificação): não repete.
  s = await run(at(2026, 9, 26, 2, 0, 40));
  assert.equal(manager.calls.length, 1);

  // Execução anterior ainda em andamento no próximo horário: esta é pulada.
  s = await run(at(2026, 9, 27, 2, 0, 5));
  assert.equal(manager.calls.length, 1);
  assert.equal(s.history[0].status, 'skipped');
  assert.match(s.history[0].message, /ainda está em andamento/);
  assert.equal(s.nextRunAt, at(2026, 9, 28, 2).toISOString());

  // A análise terminou: o resultado fica anotado no histórico (mesmo que o relatório seja excluído).
  manager.finish(scanId, { startedAt: at(2026, 9, 26, 2, 0, 11) });
  s = await run(at(2026, 9, 27, 12));
  const entry = s.history.find((h) => h.scanId === scanId);
  assert.equal(entry.outcome.status, 'completed');
  assert.equal(entry.outcome.matched, 2);
});

test('horários perdidos: executar uma vez na volta ou apenas registrar', async () => {
  const late = await world({ rule: { frequency: 'hourly', startDate: '2026-09-25', time: '00:00', untilTime: '23:59', weekdays: [0, 1, 2, 3, 4, 5, 6] } });
  assert.equal(late.schedule.nextRunAt, at(2026, 9, 25, 11).toISOString());
  // O CLEAN ficou parado das 10h às 15h20: o horário das 11h (e os das 12h às 15h) foram perdidos.
  let s = await late.run(at(2026, 9, 25, 15, 20));
  assert.equal(late.manager.calls.length, 1, 'uma única execução atrasada');
  assert.equal(s.history[0].trigger, 'catch-up');
  assert.match(s.history[0].message, /Execução atrasada\. O horário de 25\/09\/2026 11:00 foi perdido.*Outros 4 horários também foram perdidos/);
  assert.equal(s.nextRunAt, at(2026, 9, 25, 16).toISOString());
  const log = late.store.getScan(s.history[0].scanId).log.map((l) => l.message).join(' ');
  assert.match(log, /Execução atrasada/);

  // Atraso pequeno (servidor ocupado): é a execução normal.
  late.manager.active.clear();
  s = await late.run(at(2026, 9, 25, 16, 3));
  assert.equal(s.history[0].status, 'started');
  assert.equal(s.history[0].trigger, 'schedule');
  assert.equal(s.history[0].message, '');

  const skip = await world({ catchUp: false });
  s = await skip.run(at(2026, 9, 28, 9));
  assert.equal(skip.manager.calls.length, 0);
  assert.equal(s.history[0].status, 'missed');
  assert.match(s.history[0].message, /O horário de 26\/09\/2026 02:00 foi perdido.*Outros 2 horários também foram perdidos/);
  assert.equal(s.nextRunAt, at(2026, 9, 29, 2).toISOString());
});

test('execução única termina o agendamento; pausado não executa', async () => {
  const once = await world({ rule: { frequency: 'once', startDate: '2026-09-25', time: '18:30' } });
  let s = await once.run(at(2026, 9, 25, 18, 30, 5));
  assert.equal(once.manager.calls.length, 1);
  assert.equal(s.nextRunAt, null);
  s = await once.run(at(2026, 9, 30));
  assert.equal(once.manager.calls.length, 1);

  const paused = await world();
  paused.store.updateSchedule(paused.schedule.id, { enabled: false, nextRunAt: null });
  await paused.run(at(2026, 9, 26, 2, 0, 5));
  assert.equal(paused.manager.calls.length, 0);
});

test('período: últimos N dias e incremental com análise completa periódica', async () => {
  const days = await world({ period: { type: 'days', days: 7 } });
  await days.run(at(2026, 9, 26, 2, 0, 5));
  assert.equal(days.manager.calls[0].body.options.modifiedAfter, new Date(+at(2026, 9, 26, 2, 0, 5) - 7 * 86400000).toISOString());

  const inc = await world({ period: { type: 'since-last', fullEvery: 3 } });
  const { manager, run, store, list } = inc;
  const from = (i) => manager.calls[i].body.options.modifiedAfter;
  const finish = (i, date) => manager.finish(manager.calls[i].scanId, { startedAt: date });

  await run(at(2026, 9, 26, 2, 0, 5));
  assert.equal(from(0), null, 'primeira execução: completa');
  finish(0, at(2026, 9, 26, 2, 0, 6));
  let s = await run(at(2026, 9, 27, 2, 0, 5));
  assert.equal(from(1), new Date(+at(2026, 9, 26, 2, 0, 6) - INCREMENTAL_MARGIN_MS).toISOString(), 'desde o início da anterior, com a margem');
  assert.match(s.history[0].message, /análise incremental/);
  finish(1, at(2026, 9, 27, 2, 0, 6));
  await run(at(2026, 9, 28, 2, 0, 5));
  assert.equal(from(2), new Date(+at(2026, 9, 27, 2, 0, 6) - INCREMENTAL_MARGIN_MS).toISOString());
  finish(2, at(2026, 9, 28, 2, 0, 6));
  s = await run(at(2026, 9, 29, 2, 0, 5));
  assert.equal(from(3), null, 'a cada 3 execuções, uma completa');
  assert.match(s.history[0].message, /completa periódica/);

  // Uma execução que falhou não serve de base: a seguinte parte da última concluída.
  manager.finish(s.history[0].scanId, { status: 'failed', startedAt: at(2026, 9, 29, 2, 0, 6) });
  await run(at(2026, 9, 30, 2, 0, 5));
  assert.equal(from(4), null, 'a completa periódica que falhou é repetida');
  manager.finish(store.getSchedule(inc.schedule.id).history[0].scanId, { startedAt: at(2026, 9, 30, 2, 0, 6) });

  // Termos alterados: as execuções anteriores não viram os termos novos, então a próxima é completa.
  await run(at(2026, 10, 1, 2, 0, 5));
  assert.ok(from(5), 'incremental');
  manager.finish(store.getSchedule(inc.schedule.id).history[0].scanId, { startedAt: at(2026, 10, 1, 2, 0, 6) });
  store.updateList(list.id, { terms: [...list.terms, { id: 't2', type: 'text', value: 'CPF' }] });
  s = await run(at(2026, 10, 2, 2, 0, 5));
  assert.equal(from(6), null);
  assert.match(s.history[0].message, /não há execução anterior concluída com os mesmos locais, termos e opções/);
});

test('exclusão automática só com o que foi confirmado ao salvar', async () => {
  const w = await world({ action: 'delete' });
  let s = await w.run(at(2026, 9, 26, 2, 0, 5));
  const call = w.manager.calls[0];
  assert.equal(call.body.options.deleteMatches, true);
  assert.equal(call.body.confirmDelete, 'EXCLUIR');
  assert.match(call.by, /^agendamento "Varredura noturna" \(exclusão automática confirmada por acesso local em 25\/09\/2026 10:00\)$/);
  w.manager.finish(s.history[0].scanId, { startedAt: at(2026, 9, 26, 2, 0, 6) });

  // Exclusão desligada no repositório: a execução falha com a explicação.
  w.store.updateRepository(w.repo.id, { allowDelete: false });
  s = await w.run(at(2026, 9, 27, 2, 0, 5));
  assert.equal(w.manager.calls.length, 1);
  assert.equal(s.history[0].status, 'failed');
  assert.match(s.history[0].message, /A exclusão foi desligada no cadastro do repositório "Arquivos"/);
  assert.equal(s.nextRunAt, at(2026, 9, 28, 2).toISOString(), 'o agendamento continua');

  // Permitida de novo, mas em outra pasta: não vale a confirmação antiga.
  w.store.updateRepository(w.repo.id, { allowDelete: true, path: path.join(root, 'outra-pasta') });
  s = await w.run(at(2026, 9, 28, 2, 0, 5));
  assert.equal(s.history[0].status, 'failed');
  assert.match(s.history[0].message, /mudou \(o caminho, as contas ou os sites\) depois que a exclusão automática foi confirmada/);

  // Repositório na nuvem que passou da lixeira para a exclusão definitiva.
  const cloud = { type: 'sharepoint', name: 'SP', path: 'SharePoint: todos os sites', graph: { tenantId: 'contoso.onmicrosoft.com' }, cloud: { scope: 'all', accounts: [], sites: [], exclude: [] }, deleteMode: 'trash', allowDelete: true };
  w.store.updateRepository(w.repo.id, cloud);
  w.store.updateSchedule(w.schedule.id, { deleteConfirmation: { by: 'x', at: new Date().toISOString(), targets: deletionPins('files', [w.store.getRepository(w.repo.id)]) } });
  w.store.updateRepository(w.repo.id, { deleteMode: 'permanent' });
  s = await w.run(at(2026, 9, 29, 2, 0, 5));
  assert.match(s.history[0].message, /passou a ser definitiva/);
  w.store.updateRepository(w.repo.id, { deleteMode: 'trash' });
  s = await w.run(at(2026, 9, 30, 2, 0, 5));
  assert.equal(s.history[0].status, 'started');

  // Repositório excluído do cadastro: falha com o nome salvo no agendamento.
  w.manager.active.clear();
  w.store.deleteRepository(w.repo.id);
  s = await w.run(at(2026, 10, 1, 2, 0, 5));
  assert.match(s.history[0].message, /O repositório "Arquivos" foi excluído do cadastro/);
});

test('guarda só os últimos relatórios do agendamento', async () => {
  const w = await world({ keepLast: 2 });
  const ids = [];
  for (let day = 26; day <= 29; day++) {
    const s = await w.run(at(2026, 9, day, 2, 0, 5));
    ids.push(s.history[0].scanId);
    w.manager.finish(s.history[0].scanId, { startedAt: at(2026, 9, day, 2, 0, 6) });
  }
  const own = w.store.listScans().filter((x) => x.scheduleId === w.schedule.id).map((x) => x.id);
  assert.deepEqual(own.sort(), ids.slice(-2).sort(), 'os dois mais recentes ficam');
  const s = w.store.getSchedule(w.schedule.id);
  const first = s.history.find((h) => h.scanId === ids[0]);
  assert.equal(first.outcome.status, 'completed', 'o resultado continua no histórico');
  assert.match(w.store.getScan(ids[3]).log.map((l) => l.message).join(' '), /1 relatório\(s\) antigo\(s\) excluído\(s\)/);
});

// ------------------------------------------------------------------------------------------------
// API, com análises de verdade.

async function startApi() {
  const store = await newStore();
  const manager = new ScanManager(store);
  const clock = { now: at(2026, 9, 25, 10) };
  const scheduler = new Scheduler({ store, manager, now: () => new Date(clock.now) });
  const app = createApp({ store, manager, scheduler, config: { authUser: '', authPassword: '' } });
  const server = await new Promise((resolve) => {
    const s = app.listen(0, '127.0.0.1', () => resolve(s));
  });
  const base = `http://127.0.0.1:${server.address().port}`;
  const api = async (method, url, body) => {
    const res = await fetch(base + url, { method, headers: { 'Content-Type': 'application/json', 'X-CLEAN': '1' }, body: body === undefined ? undefined : JSON.stringify(body) });
    return { status: res.status, data: res.status === 204 ? null : await res.json() };
  };
  return { store, manager, scheduler, clock, server, api };
}

async function waitScan(api, id) {
  for (let i = 0; i < 200; i++) {
    const { data } = await api('GET', `/api/scans/${id}`);
    if (!['queued', 'running'].includes(data.status)) return data;
    await new Promise((r) => setTimeout(r, 50));
  }
  throw new Error('a análise não terminou');
}

test('API dos agendamentos', async () => {
  const { api, clock, scheduler, server, manager } = await startApi();
  try {
    const dir = path.join(root, `api-${Math.random().toString(36).slice(2)}`);
    fs.mkdirSync(dir);
    fs.writeFileSync(path.join(dir, 'folha-salario.txt'), 'confidencial');
    const repo = (await api('POST', '/api/repositories', { name: 'Arquivos', path: dir })).data;
    const list = (await api('POST', '/api/lists', { name: 'Sensíveis', terms: [{ type: 'text', value: 'salário' }] })).data;
    const rule = { frequency: 'daily', startDate: '2026-09-25', time: '02:00' };
    const body = { kind: 'files', name: 'Noturna', repositoryIds: [repo.id], listIds: [list.id], options: { checkName: true, checkContent: true }, rule, period: { type: 'since-last', fullEvery: 0 } };

    const info = await api('GET', '/api/info');
    assert.ok(typeof info.data.timeZone === 'string');

    // Validação.
    assert.match((await api('POST', '/api/schedules', { ...body, name: '' })).data.error, /Informe o nome do agendamento/);
    assert.match((await api('POST', '/api/schedules', { ...body, rule: { ...rule, time: '25:00' } })).data.error, /HH:MM/);
    assert.match((await api('POST', '/api/schedules', { ...body, repositoryIds: [] })).data.error, /repositórios válidos/);
    assert.match((await api('POST', '/api/schedules', { ...body, period: { type: 'days', days: 0 } })).data.error, /de 1 a 3650 dias/);
    const past = await api('POST', '/api/schedules', { ...body, rule: { frequency: 'once', startDate: '2026-09-25', time: '09:00' } });
    assert.equal(past.status, 400);
    assert.match(past.data.error, /nunca seria executado/);
    const noDelete = await api('POST', '/api/schedules', { ...body, options: { ...body.options, deleteMatches: true }, confirmDelete: 'EXCLUIR' });
    assert.match(noDelete.data.error, /A exclusão não está permitida em "Arquivos"/);

    // Prévia da regra.
    const preview = await api('POST', '/api/schedules/preview', { rule: { frequency: 'weekly', startDate: '2026-09-25', time: '22:00', weekdays: [1, 5] } });
    assert.equal(preview.data.description, 'Às segundas e sextas, às 22:00');
    assert.equal(preview.data.next.length, 5);
    assert.equal(preview.data.next[0], at(2026, 9, 25, 22).toISOString());
    assert.equal((await api('POST', '/api/schedules/preview', { rule: { frequency: 'weekly' } })).status, 400);

    const created = await api('POST', '/api/schedules', body);
    assert.equal(created.status, 201);
    const id = created.data.id;
    assert.equal(created.data.state, 'active');
    assert.equal(created.data.nextRunAt, at(2026, 9, 26, 2).toISOString());
    assert.equal(created.data.description, 'Todos os dias, às 02:00');
    assert.equal(created.data.createdBy, 'acesso local');
    assert.deepEqual(created.data.targets, [{ id: repo.id, name: 'Arquivos', missing: false }]);
    assert.equal(created.data.names, undefined);

    // O repositório e a lista em uso não podem ser excluídos.
    const inUse = await api('DELETE', `/api/repositories/${repo.id}`);
    assert.equal(inUse.status, 409);
    assert.equal(inUse.data.code, 'in-use');
    assert.match(inUse.data.error, /O repositório está em uso no agendamento "Noturna"\. Retire-o do agendamento/);
    assert.match((await api('DELETE', `/api/lists/${list.id}`)).data.error, /A lista está em uso no agendamento "Noturna"\. Retire-a/);

    // Primeira execução no horário (completa) e a segunda, incremental.
    clock.now = at(2026, 9, 26, 2, 0, 5);
    await scheduler.tick();
    let s = (await api('GET', `/api/schedules/${id}`)).data;
    const first = await waitScan(api, s.history[0].scanId);
    assert.equal(first.status, 'completed');
    assert.equal(first.stats.filesMatched, 1);
    assert.equal(first.scheduleId, id);
    assert.equal(first.options.modifiedAfter, null);
    s = (await api('GET', `/api/schedules/${id}`)).data;
    assert.equal(s.lastRun.outcome.status, 'completed');
    assert.equal(s.lastRun.scanStatus, 'completed');
    clock.now = at(2026, 9, 27, 2, 0, 5);
    await scheduler.tick();
    s = (await api('GET', `/api/schedules/${id}`)).data;
    const second = await waitScan(api, s.history[0].scanId);
    assert.equal(second.options.modifiedAfter, new Date(Date.parse(first.startedAt) - INCREMENTAL_MARGIN_MS).toISOString());
    assert.equal(second.stats.filesMatched, 1, 'o arquivo foi criado depois da margem');

    // Executar agora (não muda a próxima execução programada).
    const now = await api('POST', `/api/schedules/${id}/run`, {});
    assert.equal(now.status, 201);
    assert.match(now.data.scan.startedBy, /^acesso local \(Executar agora, agendamento "Noturna"\)$/);
    assert.equal(now.data.schedule.nextRunAt, at(2026, 9, 28, 2).toISOString());
    await waitScan(api, now.data.scan.id);
    assert.equal(manager.isActive(now.data.scan.id), false);

    // Pausar e retomar: os horários da pausa não são executados.
    const paused = await api('POST', `/api/schedules/${id}/pause`);
    assert.equal(paused.data.state, 'paused');
    assert.equal(paused.data.nextRunAt, null);
    clock.now = at(2026, 9, 29, 12);
    await scheduler.tick();
    assert.equal((await api('GET', `/api/schedules/${id}`)).data.history.length, 3);
    const resumed = await api('POST', `/api/schedules/${id}/resume`);
    assert.equal(resumed.data.state, 'active');
    assert.equal(resumed.data.nextRunAt, at(2026, 9, 30, 2).toISOString());

    // Exclusão automática: permitida no repositório e confirmada ao salvar.
    await api('PUT', `/api/repositories/${repo.id}`, { name: 'Arquivos', path: dir, allowDelete: true });
    const withDelete = { ...body, options: { ...body.options, deleteMatches: true } };
    assert.match((await api('PUT', `/api/schedules/${id}`, withDelete)).data.error, /digite EXCLUIR/);
    const saved = await api('PUT', `/api/schedules/${id}`, { ...withDelete, confirmDelete: 'excluir' });
    assert.equal(saved.status, 200);
    assert.equal(saved.data.action, 'delete');
    assert.equal(saved.data.deleteConfirmation.by, 'acesso local');
    assert.equal(saved.data.deleteConfirmation.targets, undefined, 'o alcance registrado não vai para a interface');
    assert.equal((await api('POST', `/api/schedules/${id}/run`, {})).status, 400, 'executar agora com exclusão pede confirmação');

    // O repositório mudou de pasta: a exclusão confirmada não vale mais.
    const other = path.join(root, `api-outra-${Math.random().toString(36).slice(2)}`);
    fs.mkdirSync(other);
    await api('PUT', `/api/repositories/${repo.id}`, { name: 'Arquivos', path: other, allowDelete: true });
    s = (await api('GET', `/api/schedules/${id}`)).data;
    assert.equal(s.problems.length, 1);
    assert.match(s.problems[0], /mudou/);
    const refused = await api('POST', `/api/schedules/${id}/run`, { confirm: true });
    assert.equal(refused.status, 400);
    assert.match(refused.data.error, /confirme a exclusão de novo/);
    s = (await api('GET', `/api/schedules/${id}`)).data;
    assert.equal(s.history[0].status, 'failed');

    // Tipo do agendamento não muda; exclusão do agendamento libera o repositório.
    assert.match((await api('PUT', `/api/schedules/${id}`, { ...body, kind: 'mail' })).data.error, /trocar o tipo/);
    assert.equal((await api('DELETE', `/api/schedules/${id}`)).status, 204);
    assert.equal((await api('GET', `/api/schedules/${id}`)).status, 404);
    assert.equal((await api('DELETE', `/api/repositories/${repo.id}`)).status, 204);
  } finally {
    server.close();
    await scheduler.stop();
  }
});
