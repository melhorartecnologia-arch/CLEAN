import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { unzipSync, strFromU8 } from 'fflate';
import { createApp } from '../src/app.js';
import { Store } from '../src/store.js';
import { ScanManager } from '../src/scan/manager.js';

const FIXTURES = path.join(import.meta.dirname, 'fixtures');
let root;
let server;
let base;
let store;
const stores = [];

async function startServer(config = {}) {
  const s = await new Store(path.join(root, `data-${Math.random().toString(36).slice(2)}`)).init();
  stores.push(s);
  const manager = new ScanManager(s);
  const app = createApp({ store: s, manager, config: { authUser: '', authPassword: '', ...config } });
  const srv = await new Promise((resolve) => {
    const x = app.listen(0, '127.0.0.1', () => resolve(x));
  });
  return { srv, store: s, url: `http://127.0.0.1:${srv.address().port}` };
}

async function api(method, url, body, headers = {}) {
  const res = await fetch(base + url, {
    method,
    headers: { 'Content-Type': 'application/json', 'X-CLEAN': '1', ...headers },
    body: body === undefined ? undefined : JSON.stringify(body),
  });
  const type = res.headers.get('content-type') || '';
  const data = type.includes('application/json') ? await res.json() : await res.arrayBuffer();
  return { status: res.status, data, headers: res.headers };
}

before(async () => {
  root = fs.mkdtempSync(path.join(os.tmpdir(), 'clean-api-'));
  const repo = path.join(root, 'Arquivos');
  fs.mkdirSync(path.join(repo, 'RH'), { recursive: true });
  fs.copyFileSync(path.join(FIXTURES, 'plan.xlsx'), path.join(repo, 'RH', 'folha_salarios.xlsx'));
  fs.copyFileSync(path.join(FIXTURES, 'doc.pdf'), path.join(repo, 'contrato.pdf'));
  fs.writeFileSync(path.join(repo, 'RH', 'anotacoes.txt'), '=HYPERLINK("x") confidencial\nsem mais');
  fs.writeFileSync(path.join(repo, 'leia-me.txt'), 'nada aqui');
  ({ srv: server, store, url: base } = await startServer());
});

after(async () => {
  server?.close();
  await Promise.all(stores.map((s) => s.close()));
  fs.rmSync(root, { recursive: true, force: true });
});

test('fluxo completo pela API', async () => {
  const info = await api('GET', '/api/info');
  assert.equal(info.status, 200);
  assert.ok(info.data.presets.some((p) => p.id === 'cpf'));

  const bad = await api('POST', '/api/repositories', { name: '', path: 'relativo' });
  assert.equal(bad.status, 400);
  assert.match(bad.data.error, /Informe/);

  const tested = await api('POST', '/api/repositories/test', { path: path.join(root, 'Arquivos') });
  assert.equal(tested.data.ok, true);
  assert.match(tested.data.message, /1 pasta\(s\) e 2 arquivo\(s\)/);
  const missing = await api('POST', '/api/repositories/test', { path: path.join(root, 'nada') });
  assert.equal(missing.data.ok, false);

  const repo = await api('POST', '/api/repositories', {
    name: 'Arquivos',
    path: `${path.join(root, 'Arquivos')}/`,
    exclude: 'Temp\n\n*.bak',
    audit: { enabled: false, days: 900 },
  });
  assert.equal(repo.status, 201);
  assert.equal(repo.data.path, path.join(root, 'Arquivos'));
  assert.deepEqual(repo.data.exclude, ['Temp', '*.bak']);
  assert.equal(repo.data.audit.days, 365);

  const badList = await api('POST', '/api/lists', { name: 'X', terms: [{ type: 'regex', value: '(' }] });
  assert.equal(badList.status, 400);
  assert.match(badList.data.error, /Termo 1/);

  const list = await api('POST', '/api/lists', {
    name: 'Sensíveis',
    terms: [
      { type: 'text', value: 'salário' },
      { type: 'text', value: 'SALARIO' }, // duplicado (mesmo termo sem acento) é descartado
      { type: 'text', value: 'confidencial', wholeWord: true },
      { type: 'regex', value: String.raw`\b\d{3}\.?\d{3}\.?\d{3}-?\d{2}\b`, validator: 'cpf', label: 'CPF' },
    ],
  });
  assert.equal(list.status, 201);
  assert.equal(list.data.terms.length, 3);
  const lists = await api('GET', '/api/lists');
  assert.equal(lists.data[0].termCount, 3);

  const tried = await api('POST', '/api/lists/test', { terms: list.data.terms, text: 'CPF 529.982.247-25 e 111.111.111-11; salario' });
  assert.deepEqual(tried.data.matches.map((m) => [m.term, m.count]).sort(), [['CPF', 1], ['salário', 1]]);
  const slow = await api('POST', '/api/lists/test', { terms: [{ type: 'regex', value: '(a+)+$' }], text: `${'a'.repeat(40)}!` });
  assert.equal(slow.status, 422);

  const scan = await api('POST', '/api/scans', { repositoryIds: [repo.data.id], listIds: [list.data.id], options: { checkName: true } });
  assert.equal(scan.status, 201);
  let current;
  for (let i = 0; i < 200; i++) {
    current = await api('GET', `/api/scans/${scan.data.id}`);
    if (!['queued', 'running'].includes(current.data.status)) break;
    await new Promise((r) => setTimeout(r, 100));
  }
  assert.equal(current.data.status, 'completed');
  assert.equal(current.data.stats.filesMatched, 3);

  const results = await api('GET', `/api/scans/${scan.data.id}/results?sort=name`);
  assert.equal(results.data.total, 3);
  assert.deepEqual(results.data.items.map((r) => r.name), ['anotacoes.txt', 'contrato.pdf', 'folha_salarios.xlsx']);
  assert.ok(!('_search' in results.data.items[0]));

  const filtered = await api('GET', `/api/scans/${scan.data.id}/results?q=${encodeURIComponent('carlos financeiro')}`);
  assert.deepEqual(filtered.data.items.map((r) => r.name), ['folha_salarios.xlsx']);
  const byTerm = await api('GET', `/api/scans/${scan.data.id}/results?term=CPF&location=content`);
  assert.deepEqual(byTerm.data.items.map((r) => r.name).sort(), ['contrato.pdf', 'folha_salarios.xlsx']);
  const paged = await api('GET', `/api/scans/${scan.data.id}/results?pageSize=1&page=2&sort=name`);
  assert.equal(paged.data.pages, 3);
  assert.equal(paged.data.items[0].name, 'contrato.pdf');

  const summary = await api('GET', `/api/scans/${scan.data.id}/summary`);
  assert.equal(summary.data.files, 3);
  assert.deepEqual(summary.data.options.terms, ['confidencial', 'CPF', 'salário']);
  assert.deepEqual(summary.data.options.extensions, ['.pdf', '.txt', '.xlsx']);
  const filteredSummary = await api('GET', `/api/scans/${scan.data.id}/summary?extension=.pdf`);
  assert.equal(filteredSummary.data.files, 1);
  assert.equal(filteredSummary.data.options.terms.length, 3, 'opções de filtro consideram todos os resultados');
  const cpf = summary.data.byTerm.find((t) => t.term === 'CPF');
  assert.equal(cpf.files, 2);
  assert.ok(summary.data.byUser.some((u) => u.user === 'Carlos Financeiro'));

  // Exportações
  const xlsx = await api('GET', `/api/scans/${scan.data.id}/export.xlsx`);
  assert.equal(xlsx.status, 200);
  assert.match(xlsx.headers.get('content-disposition'), /attachment; filename="relatorio-Analise-de-/);
  const parts = unzipSync(new Uint8Array(xlsx.data));
  assert.ok(parts['xl/worksheets/sheet2.xml']);
  const arquivos = strFromU8(parts['xl/worksheets/sheet2.xml']);
  assert.ok(arquivos.includes('folha_salarios.xlsx'));
  assert.ok(arquivos.includes('Carlos Financeiro'));
  assert.ok(strFromU8(parts['xl/workbook.xml']).includes('name="Ocorrências"'));

  const csv = await api('GET', `/api/scans/${scan.data.id}/export.csv`);
  const csvText = Buffer.from(csv.data).toString('utf8');
  assert.ok(csvText.startsWith('\uFEFFRepositório;Arquivo;Termo'));
  assert.ok(csvText.includes("'=HYPERLINK"), 'fórmulas são neutralizadas no CSV');
  const html = await api('GET', `/api/scans/${scan.data.id}/export.html`);
  const htmlText = Buffer.from(html.data).toString('utf8');
  assert.ok(htmlText.includes('<mark>confidencial</mark>'));
  assert.ok(!htmlText.includes('<script'));
  const json = await api('GET', `/api/scans/${scan.data.id}/export.json`);
  assert.equal(json.data.results.length, 3);
  assert.match(json.headers.get('content-disposition'), /\.json"$/);

  const errors = await api('GET', `/api/scans/${scan.data.id}/errors`);
  assert.equal(errors.data.total, 0);

  // Cancelar análise já concluída
  const cancel = await api('POST', `/api/scans/${scan.data.id}/cancel`);
  assert.equal(cancel.status, 409);

  const del = await api('DELETE', `/api/scans/${scan.data.id}`);
  assert.equal(del.status, 204);
  assert.equal((await api('GET', `/api/scans/${scan.data.id}`)).status, 404);
  assert.equal((await api('DELETE', `/api/repositories/${repo.data.id}`)).status, 204);
  assert.equal((await api('DELETE', `/api/lists/${list.data.id}`)).status, 204);
});

test('proteção contra CSRF e JSON inválido', async () => {
  const noHeader = await fetch(`${base}/api/lists`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: '{}' });
  assert.equal(noHeader.status, 403);
  const otherOrigin = await api('POST', '/api/lists', { name: 'x', terms: [] }, { Origin: 'http://malicioso.example' });
  assert.equal(otherOrigin.status, 403);
  const invalid = await fetch(`${base}/api/lists`, { method: 'POST', headers: { 'Content-Type': 'application/json', 'X-CLEAN': '1' }, body: '{x' });
  assert.equal(invalid.status, 400);
  assert.equal((await invalid.json()).error, 'JSON inválido.');
  const notFound = await api('GET', '/api/nada');
  assert.equal(notFound.status, 404);
  const info = await fetch(`${base}/api/info`);
  assert.match(info.headers.get('content-security-policy'), /default-src 'self'/);
  assert.equal(info.headers.get('x-frame-options'), 'DENY');
});

test('autenticação básica quando configurada', async () => {
  const { srv, url } = await startServer({ authUser: 'admin', authPassword: 's3nh@' });
  try {
    assert.equal((await fetch(`${url}/api/info`)).status, 401);
    const wrong = await fetch(`${url}/api/info`, { headers: { Authorization: `Basic ${Buffer.from('admin:errada').toString('base64')}` } });
    assert.equal(wrong.status, 401);
    const ok = await fetch(`${url}/api/info`, { headers: { Authorization: `Basic ${Buffer.from('admin:s3nh@').toString('base64')}` } });
    assert.equal(ok.status, 200);
    assert.equal((await ok.json()).auth, true);
  } finally {
    srv.close();
  }
});
