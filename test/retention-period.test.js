// Retenção por período: a data do critério (última abertura, modificação, as duas, criação ou uso)
// entre uma data inicial e uma final, os dois dias inteiros. Validação, limites do período no motor,
// reconferência na exclusão (automática e manual), relatório, exportações e registro de exclusões.
import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { sanitizeRetention, cutoffDate, describeRetention, amountText, expiryWindow, inWindow, RetentionError, FILE_CRITERIA } from '../src/retention/policy.js';
import { cutoffPreview, periodPreview, describeRetention as describeClient, amountText as amountClient } from '../public/js/retention.js';
import { Scanner } from '../src/scan/scanner.js';
import { Store } from '../src/store.js';
import { ScanManager } from '../src/scan/manager.js';
import { Scheduler } from '../src/schedule/scheduler.js';
import { createApp } from '../src/app.js';

let root;
const stores = [];

before(() => {
  root = fs.mkdtempSync(path.join(os.tmpdir(), 'clean-ret-period-'));
});

after(async () => {
  await Promise.all(stores.map((s) => s.close()));
  fs.rmSync(root, { recursive: true, force: true });
});

const fails = (fn, pattern) => assert.throws(fn, (err) => err instanceof RetentionError && pattern.test(err.message));

/** Pasta com arquivos; put(rel, mtime, atime) ajusta as datas (a ordem do utimes é atime, mtime). */
function repoDir() {
  const dir = fs.mkdtempSync(path.join(root, 'repo-'));
  const put = (rel, mtime, atime = mtime) => {
    const target = path.join(dir, rel);
    fs.mkdirSync(path.dirname(target), { recursive: true });
    fs.writeFileSync(target, rel);
    fs.utimesSync(target, atime, mtime);
    return target;
  };
  return { dir, put, exists: (rel) => fs.existsSync(path.join(dir, rel)) };
}

const policyFor = (input) => {
  const r = sanitizeRetention(input, 'files');
  return { ...r, cutoff: cutoffDate(r).toISOString() };
};

async function runFiles(dir, retention, { deleteMatches = true, ownerResolver = null } = {}) {
  const messages = [];
  const scanner = new Scanner(
    {
      repositories: [{ id: 'r1', name: 'Dados', path: dir, exclude: [], allowDelete: true, keep: [] }],
      terms: [],
      options: { deleteMatches, resolveOwner: Boolean(ownerResolver), checkContent: false, concurrency: 1 },
      retention,
      startedBy: 'política "Período"',
    },
    (m) => messages.push(m),
    ownerResolver ? { ownerResolver } : {},
  );
  const stats = await scanner.run();
  const records = messages.filter((m) => m.type === 'results').flatMap((m) => m.records);
  return {
    stats,
    records,
    names: records.map((r) => r.name).sort(),
    events: messages.filter((m) => m.type === 'deletions').flatMap((m) => m.items),
    logs: messages.filter((m) => m.type === 'log').map((m) => m.message),
  };
}

// ---------------------------------------------------------------------------------------------
// Política

test('período: validação, limites (os dois dias inteiros), idempotência e descrição', () => {
  const p = sanitizeRetention({ mode: 'period', from: '2015-01-01', to: '2018-12-31', criterion: 'opened-or-modified' }, 'files');
  assert.deepEqual(p, { criterion: 'opened-or-modified', from: '2015-01-01', to: '2018-12-31', maxDeletions: 1000, deleteMode: 'permanent', patterns: [] });
  assert.deepEqual(sanitizeRetention(p, 'files'), p, 'validar de novo devolve o mesmo (sem "mode", o período salvo vale)');
  assert.equal(+cutoffDate(p), +new Date(2019, 0, 1), 'o corte é o fim do dia final (00:00 do dia seguinte)');
  const w = expiryWindow({ ...p, cutoff: cutoffDate(p).toISOString() });
  assert.equal(w.floorMs, +new Date(2015, 0, 1));
  assert.equal(inWindow(+new Date(2015, 0, 1), w), true, '00:00 do dia inicial entra');
  assert.equal(inWindow(+new Date(2015, 0, 1) - 1, w), false, 'antes do dia inicial fica');
  assert.equal(inWindow(+new Date(2018, 11, 31, 23, 59, 59), w), true, 'o dia final inteiro entra');
  assert.equal(inWindow(+new Date(2019, 0, 1), w), false, 'o dia seguinte fica');
  assert.equal(inWindow(null, w), false, 'sem data: nunca expira');

  assert.equal(describeRetention(p, 'files'), 'Arquivos com a última abertura ou modificação entre 01/01/2015 e 31/12/2018');
  assert.equal(amountText(p), 'de 01/01/2015 a 31/12/2018');
  for (const criterion of Object.keys(FILE_CRITERIA)) {
    const r = sanitizeRetention({ mode: 'period', from: '2015-01-01', to: '2018-12-31', criterion, patterns: ['*.pdf'] }, 'files');
    assert.equal(describeClient(r, 'files'), describeRetention(r, 'files'), `${criterion}: a tela descreve como o servidor`);
    assert.match(describeRetention(r, 'files'), /entre 01\/01\/2015 e 31\/12\/2018 \(somente \*\.pdf\)$/);
  }
  assert.equal(amountClient(p), amountText(p));
  // Prévia da tela: o mesmo corte; período inválido não tem prévia.
  assert.equal(+cutoffPreview({ mode: 'period', from: '2015-01-01', to: '2018-12-31' }), +new Date(2018, 11, 31));
  assert.equal(periodPreview({ from: '2019-01-01', to: '2018-12-31' }), null);
  assert.equal(cutoffPreview({ mode: 'period', from: '', to: '' }), null, 'sem nenhuma das datas: inválido');
  assert.equal(cutoffPreview({ mode: 'period', from: '2015-01-01', to: '2999-01-01' }), null);
  // Um dia só (inicial = final) vale.
  assert.equal(sanitizeRetention({ mode: 'period', from: '2020-05-10', to: '2020-05-10' }, 'files').to, '2020-05-10');

  fails(() => sanitizeRetention({ mode: 'period', from: '', to: ' ' }, 'files'), /a data "a partir de", a data "até" ou as duas/);
  fails(() => sanitizeRetention({ mode: 'period', from: '2019-01-01', to: '2018-12-31' }, 'files'), /posterior à data final/);
  fails(() => sanitizeRetention({ mode: 'period', from: '2015-02-30', to: '2018-12-31' }, 'files'), /inicial do período inválida/);
  fails(() => sanitizeRetention({ mode: 'period', from: '1975-01-01', to: '2018-12-31' }, 'files'), /02\/01\/1980/);
  fails(() => sanitizeRetention({ mode: 'period', from: '2026-01-01', to: '2026-10-23' }, 'files', { now: new Date(2026, 9, 22, 12) }), /futuro/);
  fails(() => sanitizeRetention({ mode: 'period', from: '2015-01-01', to: '2018-12-31' }, 'mail'), /vale para arquivos/);
  // As outras formas ignoram um período esquecido no formulário (e o período ignora a data fixa).
  assert.equal(sanitizeRetention({ mode: 'date', since: '2020-01-01', from: '2015-01-01', to: '2018-12-31' }, 'files').from, undefined);
  assert.equal(sanitizeRetention({ mode: 'age', amount: 2, from: '2015-01-01', to: '2018-12-31' }, 'files').to, undefined);
  assert.equal(sanitizeRetention({ mode: 'period', since: '2020-01-01', from: '2015-01-01', to: '2018-12-31' }, 'files').since, undefined);
  // Políticas antigas (idade máxima, data fixa) continuam com o mesmo JSON.
  assert.deepEqual(Object.keys(sanitizeRetention({ amount: 5, unit: 'years' }, 'files')), ['criterion', 'amount', 'unit', 'maxDeletions', 'deleteMode', 'patterns']);
});

// ---------------------------------------------------------------------------------------------
// Motor

test('período no motor: expiram só as datas dentro dele (os dois dias inteiros)', async () => {
  const { dir, put, exists } = repoDir();
  put('antes.tmp', new Date(new Date(2015, 0, 1) - 1000));
  put('inicio.tmp', new Date(2015, 0, 1));
  put('meio.pdf', new Date(2017, 5, 15, 10));
  put('fim.docx', new Date(2018, 11, 31, 23, 59, 58));
  put('depois.tmp', new Date(2019, 0, 1));
  put('recente.xlsx', new Date(Date.now() - 86400000));
  const run = await runFiles(dir, policyFor({ mode: 'period', from: '2015-01-01', to: '2018-12-31', criterion: 'modified', maxDeletions: 0 }));
  assert.deepEqual(run.names, ['fim.docx', 'inicio.tmp', 'meio.pdf']);
  assert.equal(run.stats.deleted, 3);
  assert.ok(exists('antes.tmp') && exists('depois.tmp') && exists('recente.xlsx'));
  assert.ok(!exists('inicio.tmp') && !exists('meio.pdf') && !exists('fim.docx'));
  assert.match(run.logs[0], /entre 01\/01\/2015 e 31\/12\/2018 \(período/);
});

test('período com "sem abertura nem modificação": vale a mais recente das duas datas', async () => {
  const { dir, put, exists } = repoDir();
  const inside = new Date(2016, 3, 10);
  const before2015 = new Date(2012, 0, 1);
  const after2018 = new Date(2021, 0, 1);
  put('modificado-no-periodo.docx', inside, before2015); // aberto antes, modificado dentro: expira
  put('aberto-no-periodo.docx', before2015, inside); // modificado antes, aberto dentro: expira
  put('aberto-depois.docx', inside, after2018); // modificado dentro, mas aberto depois do período: fica
  put('parado-antes.docx', before2015, before2015); // nada no período nem depois: anterior ao início, fica
  const run = await runFiles(dir, policyFor({ mode: 'period', from: '2015-01-01', to: '2018-12-31', criterion: 'opened-or-modified', maxDeletions: 0 }), { deleteMatches: false });
  assert.deepEqual(run.names, ['aberto-no-periodo.docx', 'modificado-no-periodo.docx']);
  assert.ok(exists('aberto-depois.docx') && exists('parado-antes.docx'));
});

test('data que sai do período entre a listagem e a exclusão: o arquivo é mantido', async () => {
  const { dir, put, exists } = repoDir();
  put('volta.tmp', new Date(2016, 0, 1));
  put('aberto.tmp', new Date(2016, 0, 1));
  // Durante a identificação dos proprietários: um arquivo tem a data levada para antes do período
  // (restaurado de um backup) e o outro é modificado agora.
  const ownerResolver = {
    async resolve(paths) {
      fs.utimesSync(path.join(dir, 'volta.tmp'), new Date(2010, 0, 1), new Date(2010, 0, 1));
      fs.utimesSync(path.join(dir, 'aberto.tmp'), new Date(), new Date());
      return new Map(paths.map((p) => [p, null]));
    },
  };
  const run = await runFiles(dir, policyFor({ mode: 'period', from: '2015-01-01', to: '2018-12-31', criterion: 'modified', maxDeletions: 0 }), { ownerResolver });
  assert.deepEqual(run.events.map((e) => e.status).sort(), ['changed', 'changed']);
  assert.ok(exists('volta.tmp') && exists('aberto.tmp'));
});

// ---------------------------------------------------------------------------------------------
// API: política, execução, exclusão manual, relatório, exportações e registro

test('API: política por período, simulação, exclusão manual com reconferência, exportações e registro', async () => {
  const store = await new Store(path.join(root, `data-${Math.random().toString(36).slice(2)}`)).init();
  stores.push(store);
  const manager = new ScanManager(store);
  const scheduler = new Scheduler({ store, manager });
  const app = createApp({ store, manager, scheduler, config: { authUser: '', authPassword: '' } });
  const server = await new Promise((resolve) => {
    const s = app.listen(0, '127.0.0.1', () => resolve(s));
  });
  const base = `http://127.0.0.1:${server.address().port}`;
  const api = async (method, url, body) => {
    const res = await fetch(base + url, { method, headers: { 'Content-Type': 'application/json', 'X-CLEAN': '1' }, body: body === undefined ? undefined : JSON.stringify(body) });
    return { status: res.status, data: res.status === 204 ? null : await res.json() };
  };
  const wait = async (id) => {
    for (let i = 0; i < 200; i++) {
      const { data } = await api('GET', `/api/scans/${id}`);
      if (!['queued', 'running'].includes(data.status)) return data;
      await new Promise((r) => setTimeout(r, 50));
    }
    throw new Error('não terminou');
  };
  try {
    const { dir, put, exists } = repoDir();
    put('a.pdf', new Date(2016, 1, 1));
    put('b.xlsx', new Date(2017, 1, 1));
    put('antigo.doc', new Date(2010, 1, 1));
    put('novo.doc', new Date(2022, 1, 1));
    const repoId = (await api('POST', '/api/repositories', { name: 'Dados', path: dir, allowDelete: true })).data.id;
    const body = {
      purpose: 'retention',
      kind: 'files',
      name: 'Sem uso 2015–2018',
      repositoryIds: [repoId],
      retention: { mode: 'period', from: '2015-01-01', to: '2018-12-31', criterion: 'opened-or-modified', maxDeletions: 0 },
      options: { resolveOwner: false },
      action: 'delete',
      confirmDelete: 'EXCLUIR',
      rule: null,
    };
    const bad = await api('POST', '/api/schedules', { ...body, retention: { ...body.retention, from: '2019-01-01' } });
    assert.equal(bad.status, 400);
    assert.match(bad.data.error, /posterior à data final/);

    const created = await api('POST', '/api/schedules', body);
    assert.equal(created.status, 201, JSON.stringify(created.data));
    assert.equal(created.data.retentionText, 'Arquivos com a última abertura ou modificação entre 01/01/2015 e 31/12/2018');
    assert.equal(created.data.retention.from, '2015-01-01');

    // Simulação: lista, não exclui.
    const sim = await api('POST', `/api/schedules/${created.data.id}/run`, { simulate: true });
    const simScan = await wait(sim.data.scan.id);
    assert.equal(simScan.status, 'completed');
    assert.equal(simScan.retention.cutoff, new Date(2019, 0, 1).toISOString());
    assert.equal(simScan.stats.filesMatched, 2);
    assert.ok(exists('a.pdf') && exists('b.xlsx'));

    // Exclusão manual pelo relatório: um arquivo que saiu do período (aberto agora, sem ser modificado)
    // não é excluído sem nova confirmação.
    const results = (await api('GET', `/api/scans/${simScan.id}/results?pageSize=50`)).data;
    const items = results.items || results.results || results;
    const byName = Object.fromEntries(items.map((r) => [r.name, r]));
    fs.utimesSync(path.join(dir, 'b.xlsx'), new Date(), new Date(2017, 1, 1));
    const out = await api('POST', `/api/scans/${simScan.id}/results/${byName['b.xlsx'].id}/delete`, { confirm: true });
    assert.equal(out.status, 409, JSON.stringify(out.data));
    assert.match(out.data.error, /não está mais expirado/);
    assert.ok(exists('b.xlsx'));

    // Execução com exclusão: só o que continua no período.
    const run = await api('POST', `/api/schedules/${created.data.id}/run`, { confirm: true });
    const scan = await wait(run.data.scan.id);
    assert.equal(scan.status, 'completed', JSON.stringify(scan.log));
    assert.equal(scan.stats.deleted, 1);
    assert.ok(!exists('a.pdf') && exists('b.xlsx') && exists('antigo.doc') && exists('novo.doc'));

    // Exportação da execução: o período no resumo.
    const html = await (await fetch(`${base}/api/scans/${scan.id}/export.html`)).text();
    assert.match(html, /Período \(data do critério\)/);
    assert.match(html, /de 01\/01\/2015 a 31\/12\/2018/);
    assert.doesNotMatch(html, /undefined|NaN/);

    // Registro de exclusões: a política com o período, e a coluna do período na exportação.
    await store.flushDeletionLog();
    const lines = fs.readFileSync(store.deletionLog, 'utf8').trim().split('\n').map((l) => JSON.parse(l));
    const auto = lines.find((e) => e.mode === 'retention');
    assert.deepEqual([auto.policy.from, auto.policy.to], ['2015-01-01', '2018-12-31']);
    assert.equal(auto.policy.since, undefined);
    const list = (await api('GET', `/api/deletions?scheduleId=${created.data.id}`)).data;
    const event = list.items.find((n) => n.mode === 'retention');
    assert.match(event.policyRule, /Sem abertura nem modificação: entre 01\/01\/2015 e 31\/12\/2018/);
    assert.deepEqual([event.periodFrom, event.periodTo], ['2015-01-01', '2018-12-31']);
    const csv = await (await fetch(`${base}/api/deletions/export.csv?scheduleId=${created.data.id}`)).text();
    assert.match(csv, /01\/01\/2015 a 31\/12\/2018/);
  } finally {
    server.close();
    await scheduler.stop();
    await manager.shutdown?.();
  }
});

// ---------------------------------------------------------------------------------------------
// Reconferência

test('fim do período num dia sem 00:00 (horário de verão à meia-noite): o corte não escorrega', () => {
  // Em São Paulo, o horário de verão de 2018 começou em 04/11 às 00:00 (que não existiu).
  const code = `
    import { sanitizeRetention, cutoffDate, expiryWindow, inWindow } from ${JSON.stringify(new URL('../src/retention/policy.js', import.meta.url).href)};
    const r = sanitizeRetention({ mode: 'period', from: '2018-11-01', to: '2018-11-04', criterion: 'modified' }, 'files');
    const cutoff = cutoffDate(r);
    const w = expiryWindow({ ...r, cutoff: cutoff.toISOString() });
    console.log(JSON.stringify({ cutoff: cutoff.toString(), next: inWindow(+new Date(2018, 10, 5, 0, 30), w), last: inWindow(+new Date(2018, 10, 4, 23, 30), w), first: inWindow(+new Date(2018, 10, 1, 0, 0), w) }));
  `;
  const out = JSON.parse(execFileSync(process.execPath, ['--input-type=module', '-e', code], { env: { ...process.env, TZ: 'America/Sao_Paulo' } }).toString());
  assert.match(out.cutoff, /Nov 05 2018 00:00:00/);
  assert.deepEqual([out.next, out.last, out.first], [false, true, true]);
});

test('período até hoje: o fim é o início da execução — um arquivo aberto durante a execução fica', async () => {
  const { dir, put, exists } = repoDir();
  const d = new Date();
  const today = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
  const yesterday = new Date(Date.now() - 86400000);
  put('aberto-agora.docx', yesterday, yesterday);
  put('parado.docx', yesterday, yesterday);
  const retention = policyFor({ mode: 'period', from: '2020-01-01', to: today, criterion: 'opened-or-modified', maxDeletions: 0 });
  assert.ok(Date.parse(retention.cutoff) <= Date.now(), 'o corte não fica no futuro');
  // Entre a listagem e a exclusão, um dos arquivos é aberto.
  const ownerResolver = {
    async resolve(paths) {
      await new Promise((r) => setTimeout(r, 20));
      const st = fs.statSync(path.join(dir, 'aberto-agora.docx'));
      fs.utimesSync(path.join(dir, 'aberto-agora.docx'), new Date(), st.mtime);
      return new Map(paths.map((p) => [p, null]));
    },
  };
  const run = await runFiles(dir, retention, { ownerResolver });
  const byStatus = Object.fromEntries(run.events.map((e) => [run.records.find((r) => r.id === e.recordId).name, e.status]));
  assert.deepEqual(byStatus, { 'aberto-agora.docx': 'changed', 'parado.docx': 'deleted' });
  assert.ok(exists('aberto-agora.docx') && !exists('parado.docx'));
});


// ---------------------------------------------------------------------------------------------
// Só "a partir de" ou só "até"

test('período com uma data só: "a partir de" (até o início da execução) ou "até" (sem limite no início)', () => {
  const now = new Date(2026, 9, 22, 15, 30);
  const after = sanitizeRetention({ mode: 'period', from: '2020-01-01', to: '', criterion: 'modified' }, 'files', { now });
  assert.deepEqual(after, { criterion: 'modified', from: '2020-01-01', maxDeletions: 1000, deleteMode: 'permanent', patterns: [] });
  assert.deepEqual(sanitizeRetention(after, 'files', { now }), after, 'idempotente (sem "mode", "from" sozinho continua período)');
  assert.equal(+cutoffDate(after, now), +now, 'sem "até": o fim é o início da execução');
  assert.equal(describeRetention(after, 'files'), 'Arquivos modificados pela última vez a partir de 01/01/2020');
  assert.equal(amountText(after), 'a partir de 01/01/2020');
  assert.equal(describeClient(after, 'files'), describeRetention(after, 'files'));
  assert.equal(amountClient(after), amountText(after));

  const until = sanitizeRetention({ mode: 'period', from: '', to: '2018-12-31', criterion: 'modified' }, 'files', { now });
  assert.deepEqual(until, { criterion: 'modified', to: '2018-12-31', maxDeletions: 1000, deleteMode: 'permanent', patterns: [] });
  assert.equal(+cutoffDate(until, now), +new Date(2019, 0, 1));
  const w = expiryWindow({ ...until, cutoff: cutoffDate(until, now).toISOString() });
  assert.equal(w.floorMs, null, 'sem limite no início');
  assert.equal(inWindow(+new Date(1999, 0, 1), w), true);
  assert.equal(inWindow(+new Date(2019, 0, 1), w), false);
  assert.equal(describeRetention(until, 'files'), 'Arquivos modificados pela última vez até 31/12/2018');
  assert.equal(describeClient(until, 'files'), describeRetention(until, 'files'));
  // Prévia da tela: cada lado pode ficar em branco.
  assert.deepEqual(periodPreview({ from: '2020-01-01', to: '' }, now), { from: new Date(2020, 0, 1), to: null });
  assert.deepEqual(periodPreview({ from: '', to: '2018-12-31' }, now), { from: null, to: new Date(2018, 11, 31) });
  assert.equal(periodPreview({ from: 'xx', to: '2018-12-31' }, now), null, 'uma data preenchida inválida');
});

test('motor: só "até" e só "a partir de" (o arquivo usado durante a execução fica)', async () => {
  const { dir, put, exists } = repoDir();
  put('1999.tmp', new Date(1999, 5, 1));
  put('2018.tmp', new Date(2018, 11, 31, 22));
  put('2019.tmp', new Date(2019, 0, 1, 0, 0, 1));
  const until = await runFiles(dir, policyFor({ mode: 'period', to: '2018-12-31', criterion: 'modified', maxDeletions: 0 }), { deleteMatches: false });
  assert.deepEqual(until.names, ['1999.tmp', '2018.tmp']);
  assert.match(until.logs[0], /até 31\/12\/2018, inclusive/);
  const after = await runFiles(dir, policyFor({ mode: 'period', from: '2019-01-01', criterion: 'modified', maxDeletions: 0 }), { deleteMatches: false });
  assert.deepEqual(after.names, ['2019.tmp']);
  assert.match(after.logs[0], /a partir de 01\/01\/2019 até a criação desta execução/);
  assert.ok(exists('1999.tmp') && exists('2018.tmp') && exists('2019.tmp'));
});

test('análise de arquivos: "alterados a partir de" e "alterados até" (o dia final inteiro)', async () => {
  const { sanitizeOptions } = await import('../src/scan/manager.js');
  const scannerWith = (options) => new Scanner({ repositories: [], terms: [{ id: 't', type: 'text', value: 'x', listName: 'L' }], options: sanitizeOptions({ checkContent: false, ...options }) }, () => {});
  // A tela manda "até" como 00:00 do dia seguinte (o dia escolhido conta inteiro).
  const both = scannerWith({ modifiedAfter: '2015-01-01T00:00:00', modifiedBefore: '2019-01-01T00:00:00' });
  assert.equal(both.outsideChangedRange(+new Date(2014, 11, 31, 23, 59)), true, 'antes do início');
  assert.equal(both.outsideChangedRange(+new Date(2015, 0, 1)), false, '00:00 do dia inicial');
  assert.equal(both.outsideChangedRange(+new Date(2018, 11, 31, 23, 59, 59)), false, 'o dia final inteiro');
  assert.equal(both.outsideChangedRange(+new Date(2019, 0, 1)), true, 'o dia seguinte');
  const onlyUntil = scannerWith({ modifiedBefore: '2015-01-01T00:00:00' });
  assert.equal(onlyUntil.outsideChangedRange(+new Date(1999, 0, 1)), false);
  assert.equal(onlyUntil.outsideChangedRange(+new Date(2015, 0, 1)), true);
  assert.equal(scannerWith({}).outsideChangedRange(+new Date(1999, 0, 1)), false, 'sem filtro');
  assert.throws(() => sanitizeOptions({ modifiedAfter: '2019-01-01T00:00:00', modifiedBefore: '2018-01-01T00:00:00' }), /anterior à data "alterados a partir de"/);
  assert.throws(() => sanitizeOptions({ modifiedBefore: 'ontem' }), /inválida/);
  assert.equal(Object.hasOwn(sanitizeOptions({}), 'modifiedBefore'), false, 'sem "até": as opções continuam iguais às de antes');

  // Na varredura vale a data mais recente entre a modificação e a criação (como em "a partir de"): um
  // arquivo antigo copiado agora para o repositório conta pela data da cópia.
  const { dir, put } = repoDir();
  put('antigo.txt', new Date(2014, 5, 1));
  const st = fs.statSync(path.join(dir, 'antigo.txt'));
  const run = async (options) => {
    const messages = [];
    await new Scanner(
      { repositories: [{ id: 'r1', name: 'Dados', path: dir, exclude: [], allowDelete: false, keep: [] }], terms: [{ id: 't', type: 'text', value: 'antigo', listName: 'L' }], options: sanitizeOptions({ checkContent: false, ...options }) },
      (m) => messages.push(m),
    ).run();
    return { names: messages.filter((m) => m.type === 'results').flatMap((m) => m.records).map((r) => r.name), stats: messages.at(-1)?.stats };
  };
  const tomorrow = new Date(Date.now() + 86400000);
  const day = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}T00:00:00`;
  assert.deepEqual((await run({ modifiedBefore: day(tomorrow) })).names, ['antigo.txt'], 'até hoje: entra');
  const copiedNow = st.birthtimeMs > Date.now() - 86400000;
  assert.deepEqual((await run({ modifiedBefore: '2015-01-01T00:00:00' })).names, copiedNow ? [] : ['antigo.txt'], 'copiado agora: conta a data da cópia');
});

test('período: data que não é texto (número, objeto, lista) é recusada — nunca vira um período aberto', () => {
  for (const bad of [20181231, true, { d: 1 }, ['2018-12-31']]) {
    fails(() => sanitizeRetention({ mode: 'period', from: '2015-01-01', to: bad, criterion: 'modified' }, 'files'), /data final do período/);
    fails(() => sanitizeRetention({ from: bad, to: '2018-12-31', criterion: 'modified' }, 'files'), /data inicial do período/);
  }
  // null ou em branco: ausente (período de um lado só).
  assert.equal(sanitizeRetention({ mode: 'period', from: '2015-01-01', to: null }, 'files').to, undefined);
  fails(() => sanitizeRetention({ mode: 'period', from: '2015-01-01', to: 'incompleta' }, 'files'), /data final do período/);
});
