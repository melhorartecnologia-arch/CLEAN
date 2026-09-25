import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { Scanner } from '../src/scan/scanner.js';
import { ScanManager, sanitizeOptions } from '../src/scan/manager.js';
import { Store } from '../src/store.js';
import { PRESETS } from '../src/scan/presets.js';

const FIXTURES = path.join(import.meta.dirname, 'fixtures');
let root;
let repoDir;

const cpf = PRESETS.find((p) => p.id === 'cpf');
const TERMS = [
  { id: 't1', type: 'text', value: 'salário', listName: 'RH' },
  { id: 't2', type: 'text', value: 'confidencial', listName: 'RH' },
  { id: 't3', type: 'text', value: 'demissão', listName: 'RH' },
  { id: 't4', type: 'regex', value: cpf.value, validator: 'cpf', label: 'CPF', listName: 'LGPD' },
  { id: 't5', type: 'text', value: 'senha', wholeWord: true, listName: 'Segurança' },
];

before(() => {
  root = fs.mkdtempSync(path.join(os.tmpdir(), 'clean-scan-'));
  repoDir = path.join(root, 'Compartilhado');
  const put = (rel, content) => {
    const file = path.join(repoDir, rel);
    fs.mkdirSync(path.dirname(file), { recursive: true });
    if (Buffer.isBuffer(content)) fs.writeFileSync(file, content);
    else fs.writeFileSync(file, content, 'utf8');
    return file;
  };
  put('Financeiro/salarios_2025.xlsx', fs.readFileSync(path.join(FIXTURES, 'plan.xlsx')));
  put('Financeiro/relatorio.docx', fs.readFileSync(path.join(FIXTURES, 'doc.docx')));
  put('Financeiro/protegido.docx', fs.readFileSync(path.join(FIXTURES, 'senha.docx')));
  put('RH/demissao.txt', 'Plano de demissão de João.\nCPF 111.444.777-35\n');
  put('RH/limpo.txt', 'nada relevante aqui');
  put('RH/~$lock.docx', 'arquivo de bloqueio do Office confidencial');
  put('Antigo/velho.txt', 'material confidencial antigo');
  const old = put('RH/antigo-salario.txt', 'salário de 2010');
  const past = new Date('2010-01-01T00:00:00Z');
  fs.utimesSync(old, past, past);
});

after(() => fs.rmSync(root, { recursive: true, force: true }));

async function runScanner(options = {}, extra = {}) {
  const messages = [];
  const repo = { id: 'r1', name: 'Compartilhado', path: repoDir, exclude: ['Antigo'], ...(extra.repo || {}) };
  const scanner = new Scanner({ repositories: [repo], terms: TERMS, options }, (m) => messages.push(m), extra.deps);
  if (extra.onCreate) extra.onCreate(scanner);
  const stats = await scanner.run();
  const records = messages.filter((m) => m.type === 'results').flatMap((m) => m.records);
  const byName = Object.fromEntries(records.map((r) => [r.name, r]));
  return { stats, records, byName, messages };
}

test('análise completa: nome, conteúdo, exclusões e último usuário', async () => {
  const { stats, byName, messages } = await runScanner();
  assert.deepEqual(Object.keys(byName).sort(), ['antigo-salario.txt', 'demissao.txt', 'relatorio.docx', 'salarios_2025.xlsx']);
  assert.equal(stats.filesSeen, 6); // ~$lock.docx e a pasta Antigo são ignorados
  assert.equal(stats.filesMatched, 4);
  assert.equal(stats.contentEncrypted, 1);
  assert.equal(stats.errors, 0);

  const xlsx = byName['salarios_2025.xlsx'];
  const nameHit = xlsx.matches.find((m) => m.location === 'name');
  assert.equal(nameHit.term, 'salário');
  const cpfHit = xlsx.matches.find((m) => m.term === 'CPF');
  assert.equal(cpfHit.count, 2);
  assert.deepEqual(cpfHit.values, ['52998224725', '111.444.777-35']);
  assert.match(cpfHit.samples[0].where, /Planilha "Folha de Pagamento", linha 2/);
  assert.equal(xlsx.lastUser, 'Carlos Financeiro');
  assert.equal(xlsx.lastUserSource, 'metadata');
  assert.equal(xlsx.metadata.author, 'Ana Planilha');
  assert.ok(xlsx.owner, 'proprietário do arquivo resolvido');
  assert.equal(xlsx.relativePath, path.join('Financeiro', 'salarios_2025.xlsx'));

  const txt = byName['demissao.txt'];
  assert.deepEqual(txt.terms.sort(), ['CPF', 'demissão']);
  assert.equal(txt.matches.find((m) => m.location === 'name').term, 'demissão');
  assert.equal(txt.lastUserSource, 'owner');
  assert.equal(txt.lastUser, txt.owner);

  const docx = byName['relatorio.docx'];
  assert.ok(docx.terms.includes('confidencial'));
  assert.ok(docx.terms.includes('senha'));
  assert.equal(docx.lastUser, 'Marcos Revisor');
  assert.equal(messages.at(-1).type, 'done');
});

test('filtro por data de modificação e somente nomes', async () => {
  const { byName, stats } = await runScanner({ checkContent: false, modifiedAfter: '2020-01-01' });
  assert.equal(stats.filesSkippedByDate, 1);
  assert.deepEqual(Object.keys(byName).sort(), ['demissao.txt', 'salarios_2025.xlsx']);
  // Mesmo sem ler o conteúdo, os metadados do documento informam o último usuário
  assert.equal(byName['salarios_2025.xlsx'].lastUser, 'Carlos Financeiro');
  assert.equal(byName['salarios_2025.xlsx'].contentStatus, 'not-requested');
});

test('caminho completo como nome e sem proprietário', async () => {
  const { byName } = await runScanner({ nameTarget: 'path', checkContent: false, resolveOwner: false }, { repo: { exclude: [] } });
  // A pasta "Antigo" não é excluída aqui, mas não contém termos no caminho
  assert.ok(byName['relatorio.docx'] === undefined);
  assert.equal(byName['salarios_2025.xlsx'].owner, null);
  assert.equal(byName['salarios_2025.xlsx'].matches[0].samples[0].where, 'Caminho');
});

test('log de auditoria tem prioridade sobre metadados e proprietário', async () => {
  const event = { t: '2026-09-20T10:00:00.000Z', id: 4663, u: 'bruno', d: 'EMPRESA', f: path.join(repoDir, 'Financeiro', 'relatorio.docx'), m: '0x2', a: 1, w: 1 };
  const calls = [];
  const auditQuery = async (params) => {
    calls.push(params);
    return [event];
  };
  const { byName } = await runScanner({}, { repo: { audit: { enabled: true, days: 7, ignoreUsers: ['svc-backup'] } }, deps: { auditQuery } });
  assert.equal(calls.length, 1);
  assert.equal(calls[0].days, 7);
  assert.deepEqual(calls[0].ignoreUsers, ['svc-backup']);
  const docx = byName['relatorio.docx'];
  assert.equal(docx.lastUser, 'EMPRESA\\bruno');
  assert.equal(docx.lastUserSource, 'audit');
  assert.equal(docx.audit.action, 'Gravação');
  assert.equal(byName['demissao.txt'].lastUserSource, 'owner');
});

test('falha na auditoria vira aviso e a análise continua', async () => {
  const auditQuery = async () => {
    throw new Error('acesso negado ao log');
  };
  const { byName, messages } = await runScanner({}, { repo: { audit: { enabled: true } }, deps: { auditQuery } });
  assert.ok(messages.some((m) => m.type === 'log' && m.level === 'warn' && /acesso negado ao log/.test(m.message)));
  assert.equal(byName['relatorio.docx'].lastUserSource, 'metadata');
});

test('repositório inexistente gera erro sem interromper os demais', async () => {
  const messages = [];
  const scanner = new Scanner(
    {
      repositories: [
        { id: 'x', name: 'Inexistente', path: path.join(root, 'nao-existe') },
        { id: 'r1', name: 'Compartilhado', path: repoDir },
      ],
      terms: TERMS,
    },
    (m) => messages.push(m),
  );
  const stats = await scanner.run();
  assert.equal(stats.repositoriesDone, 2);
  assert.ok(stats.filesMatched > 0);
  const errors = messages.filter((m) => m.type === 'errors').flatMap((m) => m.items);
  assert.match(errors[0].message, /Não encontrado/);
});

test('cancelamento interrompe a análise', async () => {
  const { stats, messages } = await runScanner({ concurrency: 1 }, { onCreate: (s) => s.cancel() });
  assert.equal(stats.filesSeen, 0);
  assert.equal(messages.at(-1).cancelled, true);
});

test('opções inválidas são rejeitadas', () => {
  assert.throws(() => sanitizeOptions({ checkName: false, checkContent: false }), /ao menos uma/);
  assert.throws(() => sanitizeOptions({ modifiedAfter: 'ontem' }), /inválida/);
  assert.equal(sanitizeOptions({ concurrency: 99, maxFileSizeMB: '10' }).concurrency, 16);
  assert.equal(sanitizeOptions({ maxFileSizeMB: '10' }).maxFileSizeMB, 10);
});

test('gerenciador executa a análise em uma worker thread e grava os resultados', async () => {
  const store = await new Store(path.join(root, 'data')).init();
  const repo = store.createRepository({ name: 'Compartilhado', path: repoDir, exclude: ['Antigo'] });
  const list = store.createList({ name: 'RH', terms: TERMS.map(({ listName, ...t }) => t) });
  const manager = new ScanManager(store);
  const scan = await manager.start({ repositoryIds: [repo.id], listIds: [list.id], options: {} });
  assert.ok(['queued', 'running'].includes(store.getScan(scan.id).status));
  for (let i = 0; i < 200 && ['queued', 'running'].includes(store.getScan(scan.id).status); i++) {
    await new Promise((r) => setTimeout(r, 100));
  }
  const done = store.getScan(scan.id);
  assert.equal(done.status, 'completed', JSON.stringify(done.log));
  assert.equal(done.stats.filesMatched, 4);
  const results = await store.readResults(scan.id);
  assert.equal(results.length, 4);
  assert.ok(results.every((r) => r.matches[0].termId.startsWith(`${list.id}:`)));
  assert.ok(done.log.some((l) => /concluída/.test(l.message)));
  await assert.rejects(manager.start({ repositoryIds: [], listIds: [list.id] }), /repositórios/);
  await store.close();
  const reloaded = await new Store(path.join(root, 'data')).init();
  assert.equal(reloaded.getScan(scan.id).status, 'completed');
  await reloaded.close();
});

test('expressão regular com retrocesso excessivo não trava a análise', async () => {
  const dir = path.join(root, 'Lento');
  fs.mkdirSync(dir, { recursive: true });
  fs.writeFileSync(path.join(dir, 'a.txt'), `${'a'.repeat(40)}! confidencial`);
  const messages = [];
  const scanner = new Scanner(
    {
      repositories: [{ id: 'l', name: 'Lento', path: dir }],
      terms: [
        { id: 'lenta', type: 'regex', value: '(a+)+$', label: 'Lenta' },
        { id: 'conf', type: 'text', value: 'confidencial' },
      ],
      options: { checkName: false },
      regexTimeoutMs: 300,
    },
    (m) => messages.push(m),
  );
  const started = Date.now();
  const stats = await scanner.run();
  assert.ok(Date.now() - started < 10000, 'terminou dentro do tempo');
  assert.equal(stats.filesMatched, 1, 'os demais termos continuam sendo encontrados');
  const errors = messages.filter((m) => m.type === 'errors').flatMap((m) => m.items);
  assert.match(errors[0].message, /Tempo limite/);
});

test('a fila respeita o limite de análises simultâneas e permite cancelar as que aguardam', async () => {
  const store = await new Store(path.join(root, 'data-fila')).init();
  const repo = store.createRepository({ name: 'Compartilhado', path: repoDir, exclude: ['Antigo'] });
  const list = store.createList({ name: 'RH', terms: TERMS.map(({ listName, ...t }) => t) });
  const manager = new ScanManager(store, { maxConcurrent: 1 });
  const scans = [];
  for (let i = 0; i < 3; i++) scans.push(await manager.start({ repositoryIds: [repo.id], listIds: [list.id] }));
  assert.equal(manager.running.size, 1);
  assert.equal(manager.cancel(scans[2].id), true);
  let maxRunning = 0;
  for (let i = 0; i < 300; i++) {
    const statuses = scans.map((s) => store.getScan(s.id).status);
    maxRunning = Math.max(maxRunning, statuses.filter((st) => st === 'running').length);
    if (statuses.every((st) => !['queued', 'running'].includes(st))) break;
    await new Promise((r) => setTimeout(r, 50));
  }
  assert.equal(maxRunning, 1);
  assert.deepEqual(
    scans.map((s) => store.getScan(s.id).status),
    ['completed', 'completed', 'cancelled'],
  );
  await store.close();
});
