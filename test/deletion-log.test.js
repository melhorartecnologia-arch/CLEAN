// Registro geral de exclusões (data/exclusoes.ndjson): contexto gravado com cada exclusão, leitura
// em fluxo, filtros (período no fuso do servidor, situação, origem, finalidade, tipo de local,
// política, repositório, análise e busca), paginação, resumo (itens únicos e espaço liberado),
// linhas antigas sem os campos novos e as exportações (Excel, CSV e HTML).
import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import crypto from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { unzipSync, strFromU8 } from 'fflate';
import { Store } from '../src/store.js';
import { ScanManager } from '../src/scan/manager.js';
import { createApp } from '../src/app.js';
import { DeletionLog, parseDeletionFilters, NO_POLICY, UNKNOWN_REPOSITORY } from '../src/report/deletion-log.js';

let root;
const cleanups = [];

before(() => {
  root = fs.mkdtempSync(path.join(os.tmpdir(), 'clean-deletion-log-'));
});
after(async () => {
  for (const fn of cleanups.reverse()) await fn();
  fs.rmSync(root, { recursive: true, force: true });
});

async function newStore() {
  const store = await new Store(path.join(root, `data-${crypto.randomUUID()}`)).init();
  cleanups.push(() => store.close());
  return store;
}

async function startApp(store) {
  const manager = new ScanManager(store);
  const app = createApp({ store, manager, config: { authUser: '', authPassword: '' } });
  const srv = await new Promise((resolve) => {
    const s = app.listen(0, '127.0.0.1', () => resolve(s));
  });
  const base = `http://127.0.0.1:${srv.address().port}`;
  const get = async (url) => {
    const res = await fetch(base + url);
    const type = res.headers.get('content-type') || '';
    return { status: res.status, headers: res.headers, data: type.includes('json') ? await res.json() : Buffer.from(await res.arrayBuffer()) };
  };
  cleanups.push(async () => {
    srv.close();
    await manager.shutdown();
  });
  return { get };
}

/** Instante no fuso do servidor (os filtros de período valem nele). */
const local = (y, m, d, h = 12, min = 0, s = 0) => new Date(y, m - 1, d, h, min, s).toISOString();
const appendLines = (store, lines) => fs.appendFileSync(store.deletionLog, lines.map((l) => `${typeof l === 'string' ? l : JSON.stringify(l)}\n`).join(''));
const GONE = '99999999-9999-4999-8999-999999999999'; // análise já excluída
const GONE_TEAMS = '88888888-8888-4888-8888-888888888888';

/**
 * Registro com linhas novas (com o contexto da análise e o local do item) e antigas (sem eles), de
 * análises existentes e de análises já excluídas, uma linha corrompida e a última incompleta.
 */
async function seed() {
  const store = await newStore();
  const policy = store.createSchedule({ name: 'Limpeza anual', purpose: 'retention', kind: 'files' });
  const types = store.createSchedule({ name: 'Vídeos semanais', purpose: 'types', kind: 'files' });
  const retention = { criterion: 'modified', amount: 5, unit: 'years', cutoff: '2021-10-01T03:00:00.000Z', maxDeletions: 100, deleteMode: 'permanent' };
  const scanA = store.createScan({
    kind: 'files',
    name: 'Limpeza anual – 01/10/2026',
    retention,
    scheduleId: policy.id,
    scheduleName: 'Limpeza anual',
    summary: { repositories: [{ id: 'repo-local', name: 'Arquivos RH', path: 'C:\\RH', type: 'local' }] },
  });
  const scanB = store.createScan({
    kind: 'files',
    name: 'Análise financeira',
    summary: {
      repositories: [
        { id: 'repo-fin', name: 'Financeiro', path: 'D:\\Fin', type: 'local' },
        { id: 'repo-sp', name: 'Site Financeiro', path: 'SharePoint: fin', type: 'sharepoint' },
      ],
    },
  });
  const scanC = store.createScan({
    kind: 'files',
    name: 'Vídeos',
    fileTypes: { categories: ['video'] },
    scheduleId: types.id,
    scheduleName: 'Vídeos semanais',
    summary: { repositories: [{ id: 'repo-od', name: 'OneDrive da empresa', path: 'OneDrive', type: 'onedrive' }] },
  });
  const ctxA = { scanKind: 'files', purpose: 'retention', scheduleId: policy.id, scheduleName: 'Limpeza anual', policy: { criterion: 'modified', cutoff: retention.cutoff, amount: 5, unit: 'years' } };
  const fileA = (rel, size) => ({ type: 'local', repositoryId: 'repo-local', repositoryName: 'Arquivos RH', relativePath: rel, size, criterionDate: '2019-05-01T10:00:00.000Z' });
  const by = 'política de retenção "Limpeza anual"';
  const lines = [
    // 1-3: novas (retenção); o item 2 falhou e depois foi excluído (conta uma vez)
    { scanId: scanA.id, scanName: scanA.name, ...ctxA, recordId: 1, status: 'deleted', error: null, note: null, mode: 'retention', method: 'file', by, item: 'C:\\RH\\a\\velho.doc', target: fileA('a\\velho.doc', 2048), at: local(2026, 10, 1, 0, 0, 0) },
    { scanId: scanA.id, scanName: scanA.name, ...ctxA, recordId: 2, status: 'failed', error: 'Acesso negado.', note: null, mode: 'retention', method: 'file', by, item: 'C:\\RH\\b.doc', target: fileA('b.doc', 1000), at: local(2026, 10, 5) },
    { scanId: scanA.id, scanName: scanA.name, ...ctxA, recordId: 2, status: 'deleted', error: null, note: null, mode: 'retention', method: 'file', by, item: 'C:\\RH\\b.doc', target: fileA('b.doc', 1000), at: local(2026, 10, 6) },
    // 4-5: antigas, de uma análise que ainda existe (repositório e tipo pelo resumo da análise)
    { scanId: scanB.id, scanName: scanB.name, recordId: 10, status: 'deleted', error: null, note: null, mode: 'manual', method: 'file', by: 'ana (acesso local)', item: 'D:\\Fin\\2020\\relatorio.xlsx', at: local(2026, 10, 10, 23, 59, 0) },
    { scanId: scanB.id, scanName: scanB.name, recordId: 11, status: 'missing', error: 'Não existe mais.', note: null, mode: 'manual', method: 'trash', by: 'ana (acesso local)', item: 'https://contoso.sharepoint.com/sites/fin/Docs/x.xlsx', at: local(2026, 10, 11, 0, 0, 30) },
    // 6: antiga, de uma análise de e-mail já excluída (a política vem do campo "por")
    { scanId: GONE, scanName: 'Velha', recordId: 5, status: 'deleted', error: null, note: null, mode: 'retention', method: 'permanent', by: 'política de retenção "E-mails antigos"', item: 'ana@contoso.com › Caixa de Entrada › <b>x</b>', at: local(2026, 9, 30, 23, 59, 59) },
    // 7: antiga, do Teams (análise excluída)
    { scanId: GONE_TEAMS, scanName: 'Teams antigo', recordId: 6, status: 'failed', error: 'Proibido', note: null, mode: 'auto', method: 'teams', by: 'acesso local', item: 'Equipe › Geral › Olá', at: local(2026, 10, 15) },
    // 8: nova, busca por tipo no OneDrive
    { scanId: scanC.id, scanName: scanC.name, scanKind: 'files', purpose: 'types', scheduleId: types.id, scheduleName: 'Vídeos semanais', recordId: 1, status: 'deleted', error: null, note: 'movido para a lixeira', mode: 'auto', method: 'trash', by: 'agendamento "Vídeos semanais"', item: 'https://contoso-my.sharepoint.com/personal/ana/Documents/v.mp4', target: { type: 'onedrive', repositoryId: 'repo-od', repositoryName: 'OneDrive da empresa', relativePath: 'Vídeos/v.mp4', size: 500, account: 'Ana', library: 'Documentos' }, at: local(2026, 10, 12) },
    // 9: antiga, de análise excluída, com um nome que o Excel leria como fórmula
    { scanId: GONE, scanName: 'Velha', recordId: 77, status: 'deleted', error: null, note: null, mode: 'manual', method: 'file', by: 'acesso local', item: '=1+1.tmp', at: local(2026, 10, 16) },
    '{"linha corrompida',
  ];
  appendLines(store, lines);
  // Última linha ainda sendo gravada (sem a quebra de linha): ignorada.
  fs.appendFileSync(store.deletionLog, JSON.stringify({ scanId: GONE, recordId: 500, status: 'deleted', mode: 'manual', item: 'parcial', at: local(2026, 10, 17) }));
  return { store, policy, types, scanA, scanB, scanC };
}

// ------------------------------------------------------------------------------------------------

test('appendDeletions grava o contexto da análise só no registro geral', async () => {
  const store = await newStore();
  const schedule = store.createSchedule({ name: 'Retenção RH', purpose: 'retention', kind: 'files' });
  const retention = { criterion: 'modified', amount: 3, unit: 'years', cutoff: '2023-10-21T03:00:00.000Z', maxDeletions: 10, deleteMode: 'permanent', patterns: [] };
  const scan = store.createScan({ kind: 'files', name: 'Retenção RH – 21/10', retention, scheduleId: schedule.id, scheduleName: schedule.name });
  const event = { recordId: 1, status: 'deleted', error: null, note: null, mode: 'retention', method: 'file', by: 'x', item: 'C:\\a.txt', at: new Date().toISOString() };
  await store.appendDeletions(scan.id, [event, { ...event, recordId: 2, scheduleName: 'do evento' }]);
  // Política por data fixa e com pastas; busca por tipo sem agendamento; análise de e-mail.
  const since = store.createScan({ kind: 'files', name: 'Desde', retention: { criterion: 'used', since: '2020-01-01', cutoff: '2020-01-01T03:00:00.000Z', folders: [{ repositoryId: 'r', path: 'Arquivo morto' }] } });
  const typesScan = store.createScan({ kind: 'files', name: 'Tipos', fileTypes: { categories: ['video'] } });
  const mailScan = store.createScan({ kind: 'mail', name: 'E-mail' });
  await store.appendDeletions(since.id, [event]);
  await store.appendDeletions(typesScan.id, [event]);
  await store.appendDeletions(mailScan.id, [event]);
  await store.flushDeletionLog();

  const lines = fs.readFileSync(store.deletionLog, 'utf8').trim().split('\n').map((l) => JSON.parse(l));
  assert.equal(lines.length, 5);
  assert.deepEqual(Object.keys(lines[0]).slice(0, 7), ['scanId', 'scanName', 'scanKind', 'purpose', 'scheduleId', 'scheduleName', 'policy']);
  assert.equal(lines[0].scanId, scan.id);
  assert.equal(lines[0].scanName, 'Retenção RH – 21/10');
  assert.equal(lines[0].scanKind, 'files');
  assert.equal(lines[0].purpose, 'retention');
  assert.equal(lines[0].scheduleId, schedule.id);
  assert.equal(lines[0].scheduleName, 'Retenção RH');
  assert.deepEqual(lines[0].policy, { criterion: 'modified', cutoff: retention.cutoff, amount: 3, unit: 'years' });
  assert.equal(lines[0].item, 'C:\\a.txt');
  assert.equal(lines[1].scheduleName, 'do evento', 'os campos do evento prevalecem');
  assert.deepEqual(lines[2].policy, { criterion: 'used', cutoff: '2020-01-01T03:00:00.000Z', since: '2020-01-01', folders: [{ repositoryId: 'r', path: 'Arquivo morto' }] });
  assert.equal(lines[2].scheduleId, null);
  assert.equal(lines[3].purpose, 'types');
  assert.equal(lines[3].scheduleName, null);
  assert.equal('policy' in lines[3], false);
  assert.equal(lines[4].scanKind, 'mail');
  assert.equal(lines[4].purpose, 'terms');

  // O registro da análise continua só com o evento.
  await store.flushScan(scan.id);
  const own = fs.readFileSync(path.join(store.scanDir(scan.id), 'deletions.ndjson'), 'utf8').trim().split('\n').map((l) => JSON.parse(l));
  assert.deepEqual(own[0], event);
  for (const key of ['scanId', 'scanName', 'scanKind', 'purpose', 'scheduleId', 'scheduleName', 'policy']) assert.equal(key in own[0], false, key);
});

test('readDeletionLog lê em fluxo, com a posição de cada linha, e ignora linhas ruins', async () => {
  const { store } = await seed();
  const entries = [];
  for await (const entry of store.readDeletionLog()) entries.push(entry);
  assert.equal(entries.length, 9, 'sem a linha corrompida e sem a última, incompleta');
  const bytes = fs.readFileSync(store.deletionLog);
  for (const { offset, length, event } of entries) assert.deepEqual(JSON.parse(bytes.subarray(offset, offset + length).toString('utf8')), event);
  // Leitura de linhas pela posição, em qualquer ordem.
  const refs = [entries[7], entries[0], entries[4]].map(({ offset, length }) => ({ offset, length }));
  const events = await store.readDeletionLines(refs);
  assert.deepEqual(
    events.map((e) => e.recordId),
    [1, 1, 11],
  );
  assert.equal(events[0].scanName, 'Vídeos');
  // Sem o arquivo ou com uma pasta no lugar dele: nada.
  const empty = await newStore();
  const none = [];
  for await (const entry of empty.readDeletionLog()) none.push(entry);
  assert.deepEqual(none, []);
  fs.mkdirSync(empty.deletionLog);
  for await (const entry of empty.readDeletionLog()) none.push(entry);
  assert.deepEqual(none, []);
  assert.deepEqual(await empty.readDeletionLines([{ offset: 0, length: 10 }]), [null]);
});

test('filtros: validação e período no fuso do servidor', () => {
  assert.deepEqual(parseDeletionFilters({ from: '2026-10-01', to: '2026-10-21', status: 'failed,deleted', foo: 'x', mode: 'manual' }), {
    from: '2026-10-01',
    to: '2026-10-21',
    status: 'deleted,failed',
    mode: 'manual',
  });
  assert.deepEqual(parseDeletionFilters({ status: ['missing', 'changed'] }), { status: 'missing,changed' });
  for (const query of [
    { from: '2026-02-30' },
    { from: '2026-13-01' },
    { to: '20261001' },
    { from: '01/10/2026' },
    { from: '2026-10-10', to: '2026-10-01' },
    { status: 'deleted,foo' },
    { mode: 'robot' },
    { purpose: 'x' },
    { kind: 'ftp' },
    { q: 'x'.repeat(501) },
  ]) {
    assert.throws(() => parseDeletionFilters(query), (err) => err.status === 400, JSON.stringify(query).slice(0, 80));
  }
  assert.deepEqual(parseDeletionFilters({ from: '2026-10-10', to: '2026-10-10' }), { from: '2026-10-10', to: '2026-10-10' });
});

test('API /api/deletions: filtros, paginação, linhas antigas e resumo', async () => {
  const { store, policy, types, scanA, scanB } = await seed();
  const { get } = await startApp(store);
  const total = async (query) => {
    const res = await get(`/api/deletions?${query}`);
    assert.equal(res.status, 200, `${query}: ${JSON.stringify(res.data)}`);
    return res.data.total;
  };

  const all = await get('/api/deletions');
  assert.equal(all.status, 200);
  assert.equal(all.data.total, 9);
  assert.equal(all.data.page, 1);
  assert.equal(all.data.pages, 1);
  assert.equal(all.data.pageSize, 50);
  assert.equal(all.data.logBytes, fs.statSync(store.deletionLog).size);
  // As mais recentes (no fim do arquivo) primeiro.
  assert.deepEqual(
    all.data.items.map((n) => n.recordId),
    [77, 1, 6, 5, 11, 10, 2, 2, 1],
  );
  assert.equal('time' in all.data.items[0], false);

  // Período: limites pelo dia no fuso do servidor (o dia final inteiro).
  assert.equal(await total('from=2026-10-01&to=2026-10-10'), 4);
  assert.equal(await total('from=2026-10-11'), 4);
  assert.equal(await total('to=2026-09-30'), 1);
  assert.equal(await total('from=2026-10-10&to=2026-10-10'), 1);
  for (const bad of ['from=2026-02-30', 'to=2026-1-5', 'from=2026-10-10&to=2026-10-01', 'status=sumiu', 'q=' + 'a'.repeat(501)]) {
    const res = await get(`/api/deletions?${bad}`);
    assert.equal(res.status, 400, bad);
    assert.ok(res.data.error, bad);
    assert.equal((await get(`/api/deletions/summary?${bad}`)).status, 400);
    assert.equal((await get(`/api/deletions/export.csv?${bad}`)).status, 400);
  }

  assert.equal(await total('status=deleted,missing'), 7);
  assert.equal(await total('status=failed'), 2);
  assert.equal(await total('mode=retention'), 4);
  assert.equal(await total('mode=manual'), 3);
  assert.equal(await total('mode=auto'), 2);
  assert.equal(await total('purpose=retention'), 4, 'inclusive a linha antiga da política (pelo campo "por")');
  assert.equal(await total('purpose=types'), 1);
  assert.equal(await total('purpose=terms'), 2);
  assert.equal(await total('kind=local'), 5);
  assert.equal(await total('kind=cloud'), 2);
  assert.equal(await total('kind=sharepoint'), 1);
  assert.equal(await total('kind=onedrive'), 1);
  assert.equal(await total('kind=mail'), 1);
  assert.equal(await total('kind=teams'), 1);
  assert.equal(await total(`scheduleId=${policy.id}`), 3);
  assert.equal(await total(`scheduleId=${types.id}`), 1);
  assert.equal(await total('scheduleId=none'), 4);
  assert.equal(await total('repository=repo-local'), 3);
  assert.equal(await total('repository=repo-fin'), 1);
  assert.equal(await total('repository=repo-sp'), 1);
  assert.equal(await total('repository=repo-od'), 1);
  assert.equal(await total(`scanId=${scanB.id}`), 2);
  assert.equal(await total('q=velho'), 1);
  assert.equal(await total('q=LIMPEZA%20anual'), 3);
  assert.equal(await total(`q=${encodeURIComponent('analise FINANCEIRA')}`), 2, 'sem diferenciar acentos e maiúsculas');
  assert.equal(await total('q=proibido'), 1);
  assert.equal(await total(`q=${encodeURIComponent('Caixa de Entrada')}`), 1);
  assert.equal(await total('mode=retention&kind=local&from=2026-10-02'), 2);

  // Paginação.
  const page1 = await get('/api/deletions?pageSize=4');
  assert.equal(page1.data.pages, 3);
  assert.deepEqual(
    page1.data.items.map((n) => n.recordId),
    [77, 1, 6, 5],
  );
  const page3 = await get('/api/deletions?pageSize=4&page=3');
  assert.deepEqual(
    page3.data.items.map((n) => n.recordId),
    [1],
  );
  assert.equal((await get('/api/deletions?pageSize=4&page=99')).data.page, 3);
  assert.equal((await get('/api/deletions?pageSize=9999')).data.pageSize, 500);
  assert.equal((await get('/api/deletions?pageSize=abc&page=-2')).data.pageSize, 50);

  // Linhas novas e antigas, no formato de exibição.
  const byKey = new Map(all.data.items.map((n) => [`${n.scanId}:${n.recordId}:${n.status}`, n]));
  const first = byKey.get(`${scanA.id}:1:deleted`);
  assert.equal(first.kind, 'local');
  assert.equal(first.kindLabel, 'Pasta do Windows');
  assert.equal(first.repositoryName, 'Arquivos RH');
  assert.equal(first.folder, 'a');
  assert.equal(first.size, 2048);
  assert.equal(first.scheduleName, 'Limpeza anual');
  assert.equal(first.purposeLabel, 'Política de retenção');
  assert.equal(first.originLabel, 'Automática (política)');
  assert.equal(first.cutoff, '2021-10-01T03:00:00.000Z');
  assert.equal(first.criterionDate, '2019-05-01T10:00:00.000Z');
  assert.equal(first.statusLabel, 'Excluído');
  assert.equal(first.scanExists, true);
  const legacyLocal = byKey.get(`${scanB.id}:10:deleted`);
  assert.equal(legacyLocal.kind, 'local');
  assert.equal(legacyLocal.repositoryId, 'repo-fin');
  assert.equal(legacyLocal.repositoryName, 'Financeiro');
  assert.equal(legacyLocal.relativePath, '2020\\relatorio.xlsx');
  assert.equal(legacyLocal.purpose, 'terms');
  assert.equal(legacyLocal.scheduleId, null);
  assert.equal(legacyLocal.originLabel, 'Manual');
  assert.equal(legacyLocal.size, null);
  const legacyCloud = byKey.get(`${scanB.id}:11:missing`);
  assert.equal(legacyCloud.kind, 'sharepoint');
  assert.equal(legacyCloud.repositoryName, 'Site Financeiro');
  assert.equal(legacyCloud.methodLabel, 'Para a lixeira');
  const legacyMail = byKey.get(`${GONE}:5:deleted`);
  assert.equal(legacyMail.kind, 'mail');
  assert.equal(legacyMail.statusLabel, 'Excluída');
  assert.equal(legacyMail.mailbox, 'ana@contoso.com');
  assert.equal(legacyMail.folder, 'Caixa de Entrada');
  assert.equal(legacyMail.scheduleName, 'E-mails antigos');
  assert.equal(legacyMail.purpose, 'retention');
  assert.equal(legacyMail.repositoryName, UNKNOWN_REPOSITORY);
  assert.equal(legacyMail.scanExists, false);
  assert.equal(legacyMail.scanName, 'Velha');
  const legacyTeams = byKey.get(`${GONE_TEAMS}:6:failed`);
  assert.equal(legacyTeams.kind, 'teams');
  assert.equal(legacyTeams.folder, 'Equipe › Geral');
  assert.equal(legacyTeams.detail, 'Proibido');
  const onedrive = byKey.get(`${all.data.items[1].scanId}:1:deleted`);
  assert.equal(onedrive.kind, 'onedrive');
  assert.equal(onedrive.folder, 'Ana › Documentos › Vídeos');
  assert.equal(byKey.get(`${GONE}:77:deleted`).kind, 'local');

  // Resumo: itens únicos (a falha seguida da exclusão conta uma vez) e espaço liberado.
  const summary = (await get('/api/deletions/summary')).data;
  assert.equal(summary.total, 9);
  assert.equal(summary.totals.attempts, 9);
  assert.equal(summary.totals.deleted, 6);
  assert.equal(summary.totals.failed, 2);
  assert.equal(summary.totals.missing, 1);
  assert.equal(summary.totals.changed, 0);
  assert.equal(summary.totals.uniqueDeleted, 6);
  assert.equal(summary.totals.bytesDeleted, 2048 + 1000 + 500);
  assert.equal(summary.totals.bytesUnknown, 3);
  assert.equal(summary.totals.uniquePending, 1, 'só a mensagem do Teams continua no lugar');
  assert.equal(summary.range.first, local(2026, 9, 30, 23, 59, 59));
  assert.equal(summary.range.last, local(2026, 10, 16));
  assert.deepEqual(
    summary.byDay.map((d) => d.day),
    ['2026-09-30', '2026-10-01', '2026-10-05', '2026-10-06', '2026-10-10', '2026-10-11', '2026-10-12', '2026-10-15', '2026-10-16'],
  );
  const pol = summary.byPolicy.find((g) => g.key === policy.id);
  assert.equal(pol.name, 'Limpeza anual');
  assert.equal(pol.attempts, 3);
  assert.equal(pol.uniqueDeleted, 2);
  assert.equal(pol.bytesDeleted, 3048);
  assert.equal(pol.failed, 1);
  assert.equal(pol.purpose, 'retention');
  const none = summary.byPolicy.find((g) => g.key === null);
  assert.equal(none.name, NO_POLICY);
  assert.equal(none.attempts, 4);
  const byName = summary.byPolicy.find((g) => g.key === 'nome:E-mails antigos');
  assert.equal(byName.name, 'E-mails antigos');
  assert.equal(byName.purpose, 'retention');
  const repo = summary.byRepository.find((g) => g.key === 'repo-local');
  assert.equal(repo.name, 'Arquivos RH');
  assert.equal(repo.attempts, 3);
  assert.equal(repo.uniqueDeleted, 2);
  assert.equal(summary.byRepository.find((g) => g.key === null).attempts, 3);
  assert.deepEqual(
    summary.byStatus.map((s) => [s.status, s.count]),
    [
      ['deleted', 6],
      ['missing', 1],
      ['changed', 0],
      ['failed', 2],
    ],
  );
  assert.equal(summary.byKind.find((k) => k.kind === 'local').attempts, 5);
  // Opções: as políticas de retenção primeiro, os agendamentos do cadastro e os só citados no registro.
  assert.equal(summary.options.schedules[0].purpose, 'retention');
  assert.ok(summary.options.schedules.some((s) => s.id === types.id && s.exists));
  assert.ok(summary.options.schedules.some((s) => s.id === 'nome:E-mails antigos' && !s.exists));
  assert.deepEqual(summary.options.repositories.map((r) => r.id).sort(), ['repo-fin', 'repo-local', 'repo-od', 'repo-sp']);
  assert.equal(summary.options.withoutSchedule, 4);

  // O resumo segue os filtros; as opções, não.
  const filtered = (await get(`/api/deletions/summary?scheduleId=${policy.id}`)).data;
  assert.equal(filtered.totals.attempts, 3);
  assert.equal(filtered.totals.uniqueDeleted, 2);
  assert.equal(filtered.options.repositories.length, 4);
  assert.equal((await get('/api/deletions/summary?status=failed')).data.totals.uniqueDeleted, 0);

  // A última linha termina de ser gravada: passa a contar (o cache acompanha o arquivo).
  fs.appendFileSync(store.deletionLog, '\n');
  assert.equal(await total(''), 10);
});

test('API /api/deletions: registro ausente ou substituído por uma pasta', async () => {
  const store = await newStore();
  const { get } = await startApp(store);
  const check = async () => {
    const list = await get('/api/deletions');
    assert.equal(list.status, 200);
    assert.deepEqual(list.data.items, []);
    assert.equal(list.data.total, 0);
    assert.equal(list.data.pages, 1);
    const summary = await get('/api/deletions/summary');
    assert.equal(summary.status, 200);
    assert.equal(summary.data.totals.attempts, 0);
    assert.deepEqual(summary.data.byPolicy, []);
    for (const ext of ['xlsx', 'csv', 'html']) assert.equal((await get(`/api/deletions/export.${ext}`)).status, 200, ext);
    const csv = (await get('/api/deletions/export.csv')).data.toString('utf8');
    assert.equal(csv.split('\r\n').filter(Boolean).length, 1, 'só o cabeçalho');
  };
  await check();
  fs.mkdirSync(store.deletionLog);
  await check();
});

test('exportações do registro: CSV, HTML e Excel', async () => {
  const { store } = await seed();
  const { get } = await startApp(store);

  const csv = await get('/api/deletions/export.csv');
  assert.equal(csv.status, 200);
  assert.match(csv.headers.get('content-type'), /^text\/csv/);
  assert.equal(csv.headers.get('content-disposition'), 'attachment; filename="registro-de-exclusoes-2026-09-30-a-2026-10-16.csv"');
  const text = csv.data.toString('utf8');
  assert.ok(text.startsWith('\uFEFF'));
  const rows = text.slice(1).split('\r\n');
  assert.equal(rows.at(-1), '', 'termina com CRLF');
  assert.equal(rows.length, 1 + 9 + 1);
  assert.ok(!/[^\r]\n/.test(text), 'todas as linhas terminam com CRLF');
  assert.equal(rows[0].split(';')[0], 'Quando');
  assert.equal(rows[0].split(';').length, 17);
  assert.ok(rows[0].includes('Política/Agendamento') && rows[0].includes('ID da análise'));
  assert.ok(text.includes(";'=1+1.tmp;"), 'nome de arquivo neutralizado (não vira fórmula)');
  assert.ok(!text.includes(';=1+1.tmp'));

  const filteredCsv = await get('/api/deletions/export.csv?from=2026-10-01&to=2026-10-10&kind=local');
  assert.equal(filteredCsv.headers.get('content-disposition'), 'attachment; filename="registro-de-exclusoes-2026-10-01-a-2026-10-10.csv"');
  assert.equal(filteredCsv.data.toString('utf8').split('\r\n').filter(Boolean).length, 1 + 4);

  const html = await get('/api/deletions/export.html');
  assert.equal(html.status, 200);
  assert.match(html.headers.get('content-type'), /^text\/html/);
  assert.match(html.headers.get('content-disposition'), /^attachment; filename="registro-de-exclusoes-.*\.html"$/);
  const page = html.data.toString('utf8');
  assert.ok(page.includes('&lt;b&gt;x&lt;/b&gt;'));
  assert.ok(!page.includes('<b>x</b>'));
  assert.ok(page.includes('Limpeza anual'));
  assert.ok(!page.includes('exporte em Excel para todas'));

  const xlsx = await get('/api/deletions/export.xlsx?from=2026-10-01&to=2026-10-10');
  assert.equal(xlsx.status, 200);
  assert.equal(xlsx.headers.get('content-type'), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  assert.equal(xlsx.headers.get('content-disposition'), 'attachment; filename="registro-de-exclusoes-2026-10-01-a-2026-10-10.xlsx"');
  const files = unzipSync(new Uint8Array(xlsx.data));
  const workbook = strFromU8(files['xl/workbook.xml']);
  assert.deepEqual(
    [...workbook.matchAll(/<sheet name="([^"]+)"/g)].map((m) => m[1]),
    ['Resumo', 'Exclusões'],
  );
  const resumo = strFromU8(files['xl/worksheets/sheet1.xml']);
  assert.ok(resumo.includes('de 01/10/2026 a 10/10/2026'));
  assert.ok(resumo.includes('Por política ou agendamento') && resumo.includes('Por repositório ou conexão') && resumo.includes('Por dia'));
  const sheet = strFromU8(files['xl/worksheets/sheet2.xml']);
  assert.equal((sheet.match(/<row /g) || []).length, 1 + 4);
  assert.ok(sheet.includes('C:\\RH\\a\\velho.doc'));
  assert.ok(!sheet.includes('Caixa de Entrada'), 'fora do período');
});

test('exportação HTML limitada às 10.000 exclusões mais recentes', async () => {
  const store = await newStore();
  const at = new Date().toISOString();
  const lines = [];
  for (let i = 1; i <= 10001; i++) lines.push(JSON.stringify({ scanId: GONE, scanName: 'Grande', recordId: i, status: 'deleted', mode: 'manual', method: 'file', by: 'acesso local', item: `C:\\x\\${i}.tmp`, at }));
  fs.writeFileSync(store.deletionLog, `${lines.join('\n')}\n`);
  const { get } = await startApp(store);
  const page = (await get('/api/deletions/export.html')).data.toString('utf8');
  assert.ok(page.includes('exporte em Excel para todas'));
  assert.equal((page.match(/<div class="path">/g) || []).length, 10000);
  assert.ok(page.includes('C:\\x\\10001.tmp'), 'as mais recentes');
  assert.ok(!page.includes('C:\\x\\1.tmp<'), 'a mais antiga fica de fora');
  const csv = (await get('/api/deletions/export.csv')).data.toString('utf8');
  assert.equal(csv.split('\r\n').filter(Boolean).length, 1 + 10001, 'o CSV traz todas');
});

test('DeletionLog: o cache acompanha o arquivo e o cadastro', async () => {
  const { store, scanA } = await seed();
  const log = new DeletionLog(store);
  const a = await log.query({});
  assert.equal(await log.query({}), a, 'mesma consulta: resultado do cache');
  assert.notEqual(await log.query({ status: 'deleted' }), a);
  appendLines(store, ['']);
  const b = await log.query({});
  assert.notEqual(b, a);
  assert.equal(b.total, 10);
  // Renomear a política muda o nome no resumo.
  store.updateSchedule(scanA.scheduleId, { name: 'Limpeza anual (nova)' });
  const c = await log.query({});
  assert.notEqual(c, b);
  assert.equal(c.summary.byPolicy.find((g) => g.key === scanA.scheduleId).name, 'Limpeza anual (nova)');
  // Página e eventos (das mais recentes para as mais antigas).
  const items = await log.page(c, 2, 3);
  assert.equal(items.length, 3);
  const events = [];
  for await (const n of log.events(c, { limit: 4 })) events.push(n.recordId);
  assert.deepEqual(events, [500, 77, 1, 6]);
});
