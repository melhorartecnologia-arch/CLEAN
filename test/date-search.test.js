// Busca só pela data de alteração (Nova análise › "Somente pela data de alteração"): todos os
// arquivos, de qualquer tipo, dentro de "alterados a partir de" e/ou "alterados até", sem listas de
// referência nem tipos — com a exclusão automática (até o limite) ou a revisão pelo relatório.
import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { sanitizeFileTypes, typeMatcher, describeFileTypes, categoryLabel } from '../src/types/catalog.js';
import { describeFileTypes as describeClient } from '../public/js/file-types.js';
import { Store } from '../src/store.js';
import { ScanManager } from '../src/scan/manager.js';
import { Scheduler } from '../src/schedule/scheduler.js';
import { createApp } from '../src/app.js';

let root;
const stores = [];

before(() => {
  root = fs.mkdtempSync(path.join(os.tmpdir(), 'clean-date-search-'));
});

after(async () => {
  await Promise.all(stores.map((s) => s.close()));
  fs.rmSync(root, { recursive: true, force: true });
});

/** "AAAA-MM-DDT00:00:00" (hora do servidor) de uma data. */
const day = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}T00:00:00`;

test('catálogo: "todos os arquivos" — qualquer tipo combina, com a categoria pela extensão ou "outros"', () => {
  const t = sanitizeFileTypes({ all: true, maxDeletions: 50, categories: ['video'], extensions: '.xyz', checkContent: true, minSizeMB: 10 });
  assert.deepEqual(t, { all: true, categories: [], extensions: [], checkContent: false, minSizeMB: 0, maxDeletions: 50 }, 'só a data escolhe');
  assert.equal(sanitizeFileTypes({ all: true }).maxDeletions, 1000);
  assert.throws(() => sanitizeFileTypes({ all: true, maxDeletions: -1 }), /limite de exclusões/);
  const match = typeMatcher(t);
  assert.deepEqual(match('Relatório.PDF'), { category: 'document', extension: '.pdf' });
  assert.deepEqual(match('desenho.dwg'), { category: 'other', extension: '.dwg' });
  assert.deepEqual(match('LEIAME'), { category: 'other', extension: '' });
  assert.equal(categoryLabel('other'), 'Outros arquivos');
  assert.equal(describeFileTypes(t), 'Todos os arquivos (somente pela data de alteração)');
  assert.equal(describeClient(t), describeFileTypes(t));
  // all só com o valor exato: { all: 'sim' } continua exigindo os tipos.
  assert.throws(() => sanitizeFileTypes({ all: 'sim' }), /Escolha ao menos um tipo/);
});

test('API: busca só pela data — exige uma data, não pode ser agendada, lista e exclui pelo período', async () => {
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
    const dir = fs.mkdtempSync(path.join(root, 'repo-'));
    for (const name of ['a.pdf', 'b.dwg', 'sem-extensao', 'pasta/c.mp4']) {
      fs.mkdirSync(path.dirname(path.join(dir, name)), { recursive: true });
      fs.writeFileSync(path.join(dir, name), name);
    }
    const repo = (await api('POST', '/api/repositories', { name: 'Dados', path: dir, allowDelete: true })).data;
    const tomorrow = day(new Date(Date.now() + 86400000));

    // Sem nenhuma das datas: recusada (seriam todos os arquivos do repositório).
    const noDate = await api('POST', '/api/scans', { repositoryIds: [repo.id], fileTypes: { all: true }, options: { resolveOwner: false } });
    assert.equal(noDate.status, 400);
    assert.match(noDate.data.error, /informe "Somente arquivos alterados a partir de"/);

    // Não pode ser agendada (o período de cada execução substituiria as datas).
    const sched = await api('POST', '/api/schedules', { purpose: 'types', kind: 'files', name: 'Por data', repositoryIds: [repo.id], fileTypes: { all: true }, options: {}, rule: null });
    assert.equal(sched.status, 400);
    assert.match(sched.data.error, /não pode ser agendada.*política de retenção/);

    // Exclusão exige a confirmação EXCLUIR.
    const unconfirmed = await api('POST', '/api/scans', { repositoryIds: [repo.id], fileTypes: { all: true }, options: { modifiedBefore: tomorrow, deleteMatches: true, resolveOwner: false } });
    assert.equal(unconfirmed.status, 400);

    // Só listar: todos os arquivos alterados até hoje (criados agora).
    const listed = await api('POST', '/api/scans', { name: '', repositoryIds: [repo.id], fileTypes: { all: true }, options: { modifiedBefore: tomorrow, resolveOwner: false } });
    assert.equal(listed.status, 201, JSON.stringify(listed.data));
    const scan = await wait(listed.data.id);
    assert.equal(scan.status, 'completed');
    assert.match(scan.name, /^Busca por data/);
    assert.equal(scan.stats.filesMatched, 4);
    const summary = (await api('GET', `/api/scans/${scan.id}/summary`)).data;
    assert.deepEqual(summary.types.byType.map((g) => g.key).sort(), ['document', 'other', 'video']);
    for (const name of ['a.pdf', 'b.dwg', 'sem-extensao', 'pasta/c.mp4']) assert.ok(fs.existsSync(path.join(dir, name)), 'nada excluído');

    // Fora do período (alterados até 2015; os arquivos foram criados agora): nada entra.
    const none = await wait((await api('POST', '/api/scans', { repositoryIds: [repo.id], fileTypes: { all: true }, options: { modifiedBefore: '2015-01-01T00:00:00', resolveOwner: false } })).data.id);
    assert.equal(none.stats.filesMatched, 0);
    assert.equal(none.stats.filesSkippedByDate, 4);

    // Excluir automaticamente, com limite de 3 por execução.
    const del = await api('POST', '/api/scans', {
      repositoryIds: [repo.id],
      fileTypes: { all: true, maxDeletions: 3 },
      options: { modifiedAfter: '2020-01-01T00:00:00', modifiedBefore: tomorrow, deleteMatches: true, resolveOwner: false },
      confirmDelete: 'EXCLUIR',
    });
    assert.equal(del.status, 201, JSON.stringify(del.data));
    const deleted = await wait(del.data.id);
    assert.equal(deleted.stats.deleted, 3);
    const left = ['a.pdf', 'b.dwg', 'sem-extensao', 'pasta/c.mp4'].filter((name) => fs.existsSync(path.join(dir, name)));
    assert.equal(left.length, 1, 'o limite segurou um');

    // Registro de exclusões: finalidade "Busca por data de alteração".
    await store.flushDeletionLog();
    const log = (await api('GET', '/api/deletions?purpose=date')).data;
    assert.equal(log.total, 3);
    assert.ok(log.items.every((n) => n.purposeLabel === 'Busca por data de alteração'));

    // Exportação: a busca descrita, o período e o tipo "Outros arquivos".
    const html = await (await fetch(`${base}/api/scans/${scan.id}/export.html`)).text();
    assert.match(html, /Todos os arquivos \(somente pela data de alteração\)/);
    assert.match(html, /Relatório CLEAN – busca por data de alteração/);
    assert.match(html, /Somente listar \(revisão no relatório\)/);
    const csv = await (await fetch(`${base}/api/scans/${scan.id}/export.csv`)).text();
    assert.match(csv, /Data de alteração/, 'encontrado pela data');
    const noneHtml = await (await fetch(`${base}/api/scans/${none.id}/export.html`)).text();
    assert.match(noneHtml, /Fora do período \(não listados\)<\/td><td>4/);
    assert.match(html, /Outros arquivos/);
    assert.doesNotMatch(html, /undefined|NaN/);
    const xlsx = await fetch(`${base}/api/scans/${scan.id}/export.xlsx`);
    assert.equal(xlsx.status, 200);
  } finally {
    server.close();
    await scheduler.stop();
    await manager.shutdown?.();
  }
});

test('datas do filtro: só texto de data, ano a partir de 1980 (24 digitado vira 0024 e é recusado)', async () => {
  const { sanitizeOptions } = await import('../src/scan/manager.js');
  for (const bad of ['0024-01-01T00:00:00', '1975-06-01T00:00:00', true, 1, '1', 'ontem', ['2024-01-01']]) {
    assert.throws(() => sanitizeOptions({ modifiedAfter: bad }), /inválida/, JSON.stringify(bad));
    assert.throws(() => sanitizeOptions({ modifiedBefore: bad }), /inválida/, JSON.stringify(bad));
  }
  assert.equal(sanitizeOptions({ modifiedAfter: '2024-01-01T00:00:00' }).modifiedAfter, new Date(2024, 0, 1).toISOString());
  assert.equal(sanitizeOptions({ modifiedAfter: '2026-10-15T12:34:56.789Z' }).modifiedAfter, '2026-10-15T12:34:56.789Z', 'o formato do agendador');
});
