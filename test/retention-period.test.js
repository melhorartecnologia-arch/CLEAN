// Retenção por período: a data do critério (última abertura, modificação, as duas, criação ou uso)
// entre uma data inicial e uma final, os dois dias inteiros. Validação, limites do período no motor,
// reconferência na exclusão (automática e manual), relatório, exportações e registro de exclusões.
import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
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
  assert.equal(cutoffPreview({ mode: 'period', from: '2015-01-01', to: '' }), null);
  assert.equal(cutoffPreview({ mode: 'period', from: '2015-01-01', to: '2999-01-01' }), null);
  // Um dia só (inicial = final) vale.
  assert.equal(sanitizeRetention({ mode: 'period', from: '2020-05-10', to: '2020-05-10' }, 'files').to, '2020-05-10');

  fails(() => sanitizeRetention({ mode: 'period', from: '2015-01-01' }, 'files'), /Informe a data final do período/);
  fails(() => sanitizeRetention({ mode: 'period', to: '2015-01-01' }, 'files'), /Informe a data inicial do período/);
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
