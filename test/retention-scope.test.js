// Retenção por data fixa ("sem uso desde"), critério "sem abertura nem modificação" e pastas
// escolhidas: limites da data de corte, critério sem a criação, varredura e exclusão só dentro das
// pastas (também pelo caminho real e com "excluir mesmo assim"), na nuvem, e a compatibilidade das
// políticas salvas antes (o agendador compara o JSON da política da execução com o salvo).
import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { sanitizeRetention, cutoffDate, fileDate, cloudDate, describeRetention, sanitizeFolders, RetentionError, recordInScope } from '../src/retention/policy.js';
import { cutoffPreview, describeRetention as describeClient } from '../public/js/retention.js';
import { Scanner } from '../src/scan/scanner.js';
import { deleteFile } from '../src/scan/delete.js';
import { Store } from '../src/store.js';
import { ScanManager } from '../src/scan/manager.js';
import { Scheduler, deletionPins, deletionCriteria } from '../src/schedule/scheduler.js';
import { createApp } from '../src/app.js';
import { startMockApis } from './helpers/mock-apis.js';
import { world, repo as cloudRepo, file, folder } from './helpers/cloud-world.js';

let root;
const stores = [];
const DAY = 86400000;
const OLD = new Date('2015-03-10T12:00:00Z');

before(() => {
  root = fs.mkdtempSync(path.join(os.tmpdir(), 'clean-ret-scope-'));
});

after(async () => {
  await Promise.all(stores.map((s) => s.close()));
  fs.rmSync(root, { recursive: true, force: true });
});

const fails = (fn, pattern) => assert.throws(fn, (err) => err instanceof RetentionError && pattern.test(err.message));

/** Pasta com arquivos; put(rel, mtime, atime) ajusta as datas (a ordem do utimes é atime, mtime). */
function repoDir() {
  const dir = fs.mkdtempSync(path.join(root, 'repo-'));
  const put = (rel, mtime = OLD, atime = mtime) => {
    const target = path.join(dir, rel);
    fs.mkdirSync(path.dirname(target), { recursive: true });
    fs.writeFileSync(target, rel);
    fs.utimesSync(target, atime, mtime);
    return target;
  };
  return { dir, put, exists: (rel) => fs.existsSync(path.join(dir, rel)) };
}

const policyFor = (input, repos = [{ id: 'r1', name: 'Dados', type: 'local' }]) => {
  const r = sanitizeRetention(input, 'files', { targets: repos });
  return { ...r, cutoff: cutoffDate(r).toISOString() };
};

async function runFiles(dir, retention, { deleteMatches = true, exclude = [], ownerResolver = null } = {}) {
  const messages = [];
  const scanner = new Scanner(
    {
      repositories: [{ id: 'r1', name: 'Dados', path: dir, exclude, allowDelete: true, keep: [] }],
      terms: [],
      options: { deleteMatches, resolveOwner: Boolean(ownerResolver), checkContent: false, concurrency: 1 },
      retention,
      startedBy: 'política "Limpeza"',
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
    errors: messages.filter((m) => m.type === 'errors').flatMap((m) => m.items),
    logs: messages.filter((m) => m.type === 'log').map((m) => m.message),
  };
}

// ---------------------------------------------------------------------------------------------
// Política: data fixa e critério

test('data fixa: validação, corte às 00:00 locais, idempotência e descrição', () => {
  const abs = sanitizeRetention({ mode: 'date', since: '2020-01-01', criterion: 'opened-or-modified' }, 'files');
  assert.deepEqual(abs, { criterion: 'opened-or-modified', since: '2020-01-01', maxDeletions: 1000, deleteMode: 'permanent', patterns: [] });
  assert.equal(abs.amount, undefined, 'sem idade máxima na data fixa');
  assert.deepEqual(sanitizeRetention(abs, 'files'), abs, 'validar de novo devolve o mesmo (o agendador compara o JSON)');
  assert.equal(+cutoffDate(abs, new Date(2026, 9, 21, 10)), +new Date(2020, 0, 1), '00:00 do dia, hora local');
  assert.equal(+cutoffDate(abs, new Date(2031, 0, 1)), +new Date(2020, 0, 1), 'não depende do "agora"');
  assert.equal(+cutoffPreview(abs), +cutoffDate(abs), 'prévia da tela igual ao servidor');
  assert.equal(describeRetention(abs, 'files'), 'Arquivos sem abertura nem modificação desde 01/01/2020');
  assert.equal(describeClient(abs, 'files'), describeRetention(abs, 'files'));
  assert.equal(describeRetention(sanitizeRetention({ since: '2020-01-01', criterion: 'created' }, 'files'), 'files'), 'Arquivos criados antes de 01/01/2020');
  assert.equal(describeRetention(sanitizeRetention({ since: '2019-06-30' }, 'mail'), 'mail'), 'Mensagens recebidas antes de 30/06/2019 (inclusive a Lixeira e o Lixo Eletrônico)');
  // "mode: age" ignora uma data esquecida no formulário; sem "mode", a data salva vale.
  assert.deepEqual(sanitizeRetention({ mode: 'age', since: '2020-01-01', amount: 2 }, 'files').since, undefined);
  fails(() => sanitizeRetention({ mode: 'date', since: '' }, 'files'), /Informe a data de corte/);
  fails(() => sanitizeRetention({ mode: 'date', since: '2020-02-30' }, 'files'), /inválida/);
  fails(() => sanitizeRetention({ mode: 'date', since: '1970-01-01' }, 'files'), /02\/01\/1980/);
  fails(() => sanitizeRetention({ mode: 'date', since: '2026-10-22' }, 'files', { now: new Date(2026, 9, 21, 23, 59) }), /futuro/);
  // Hoje vale (nada aberto ou modificado hoje).
  assert.equal(sanitizeRetention({ mode: 'date', since: '2026-10-21' }, 'files', { now: new Date(2026, 9, 21, 0, 0) }).since, '2026-10-21');
  assert.equal(cutoffPreview({ mode: 'date', since: '2999-01-01' }), null);
  // Idade máxima continua obrigatória no modo antigo.
  fails(() => sanitizeRetention({ amount: 0 }, 'files'), /idade máxima/);
});

test('critério "sem abertura nem modificação": a mais recente das duas, sem a criação; recusado na nuvem', () => {
  const [t1, t3, t5] = [Date.UTC(2015, 0, 1), Date.UTC(2016, 0, 1), Date.now()];
  assert.equal(fileDate({ mtimeMs: t1, atimeMs: t3, birthtimeMs: t5 }, 'opened-or-modified'), t3, 'a criação recente não conta');
  assert.equal(fileDate({ mtimeMs: t3, atimeMs: t1, birthtimeMs: t5 }, 'opened-or-modified'), t3);
  assert.equal(fileDate({ mtimeMs: t1, atimeMs: 0, birthtimeMs: t5 }, 'opened-or-modified'), t1, 'último acesso desconhecido: a modificação');
  assert.equal(fileDate({ mtimeMs: 0, atimeMs: 0, birthtimeMs: t5 }, 'opened-or-modified'), null, 'sem datas: nunca expira');
  assert.equal(fileDate({ mtimeMs: t1, atimeMs: t1, birthtimeMs: t5 }, 'used'), t5, '"sem uso" continua contando a criação');
  assert.equal(cloudDate({ lastModifiedDateTime: '2015-01-01T00:00:00Z' }, 'opened-or-modified'), null, 'na nuvem não há último acesso');
  fails(() => sanitizeRetention({ criterion: 'opened-or-modified', amount: 1 }, 'files', { cloud: true }), /OneDrive e o SharePoint não informam a última abertura/);
  fails(() => sanitizeRetention({ criterion: 'opened', amount: 1 }, 'files'), /critério/);
});

test('pastas: caminhos relativos ou completos, normalizados, sem repetir e sem sair do repositório', () => {
  const local = { id: 'r1', name: 'Dados', type: 'local', path: '/srv/dados' };
  const win = { id: 'r2', name: 'Compartilhado', type: 'local', path: 'D:\\Dados' };
  const sp = { id: 'r3', name: 'SharePoint', type: 'sharepoint', path: 'SharePoint: todos os sites' };
  const targets = [local, win, sp];
  assert.deepEqual(
    sanitizeFolders(
      [
        { repositoryId: 'r1', path: 'Fin' },
        { repositoryId: 'r1', path: '/srv/dados/Fin/sub' }, // dentro de Fin: já coberta
        { repositoryId: 'r1', path: '\\RH\\2019\\' },
        { repositoryId: 'r2', path: 'd:\\dados\\Compras' },
        { repositoryId: 'r2', path: 'COMPRAS\\Antigo' }, // Windows: sem diferenciar maiúsculas
        { repositoryId: 'r3', path: 'Documentos Compartilhados/Projetos' },
      ],
      { targets },
    ),
    [
      { repositoryId: 'r1', path: 'Fin' },
      { repositoryId: 'r1', path: 'RH/2019' },
      { repositoryId: 'r2', path: 'Compras' },
      { repositoryId: 'r3', path: 'Documentos Compartilhados/Projetos' },
    ],
  );
  // A raiz do repositório: o repositório inteiro (sem restrição).
  assert.deepEqual(sanitizeFolders([{ repositoryId: 'r1', path: '/srv/dados' }, { repositoryId: 'r1', path: 'Fin' }], { targets }), []);
  for (const bad of ['..', '../x', 'a/../b', '.', '/etc', '/srv/dadosX/a', 'C:\\x', 'a*', 'a:b', 'nome. ', '']) {
    assert.throws(() => sanitizeFolders([{ repositoryId: 'r1', path: bad }], { targets }), RetentionError, `recusa ${JSON.stringify(bad)}`);
  }
  assert.throws(() => sanitizeFolders([{ repositoryId: 'r2', path: 'E:\\Outro' }], { targets }), /não fica dentro/);
  assert.throws(() => sanitizeFolders([{ repositoryId: 'r3', path: 'https://contoso.sharepoint.com/sites/x' }], { targets }), /não o endereço/);
  assert.throws(() => sanitizeFolders([{ repositoryId: 'r9', path: 'Fin' }], { targets }), /não está marcado/);
  // Sem os repositórios para conferir: falha (nunca vale o repositório inteiro).
  assert.throws(() => sanitizeFolders([{ repositoryId: 'r1', path: 'Fin' }]), /não foi possível conferir/i);
  assert.throws(() => sanitizeRetention({ amount: 1, folders: [{ repositoryId: 'r1', path: 'Fin' }] }, 'files'), /não foi possível conferir/i);
  fails(() => sanitizeRetention({ amount: 1, folders: [{ repositoryId: 'r1', path: 'x' }] }, 'mail', { targets }), /valem para arquivos|vale para arquivos/);
  // Sem pastas: o JSON da política não ganha a chave (políticas antigas continuam iguais).
  assert.equal('folders' in sanitizeRetention({ amount: 1, folders: [] }, 'files', { targets }), false);
  const withFolders = sanitizeRetention({ amount: 1, folders: [{ repositoryId: 'r1', path: 'Fin' }] }, 'files', { targets });
  assert.deepEqual(sanitizeRetention(withFolders, 'files', { targets }), withFolders, 'idempotente');
  assert.match(describeRetention(withFolders, 'files'), /— somente na pasta Fin$/);
});

// ---------------------------------------------------------------------------------------------
// Motor: pastas do Windows

test('data fixa no motor: expira só o que é anterior às 00:00 do dia (o próprio dia fica)', async () => {
  const { dir, put, exists } = repoDir();
  const cut = new Date(2020, 0, 1);
  put('antes.tmp', new Date(+cut - 1000));
  put('dia-0h.tmp', cut);
  put('dia-12h.tmp', new Date(2020, 0, 1, 12));
  put('depois.tmp', new Date(2020, 0, 2));
  const run = await runFiles(dir, policyFor({ since: '2020-01-01', criterion: 'modified', maxDeletions: 0 }));
  assert.deepEqual(run.names, ['antes.tmp']);
  assert.equal(run.stats.deleted, 1);
  assert.ok(!exists('antes.tmp') && exists('dia-0h.tmp') && exists('dia-12h.tmp') && exists('depois.tmp'));
  assert.match(run.logs[0], /anterior a 01\/01\/2020.*\(data fixa\)/);
});

test('"sem abertura nem modificação" no motor: aberto ou modificado depois fica; criado agora e sem uso expira', async () => {
  const { dir, put, exists } = repoDir();
  const recent = new Date(Date.now() - DAY);
  put('parado.docx', OLD, OLD);
  put('aberto-ontem.docx', OLD, recent);
  put('modificado-ontem.docx', recent, OLD);
  const st = fs.statSync(path.join(dir, 'parado.docx'));
  const run = await runFiles(dir, policyFor({ criterion: 'opened-or-modified', since: '2020-01-01', maxDeletions: 0 }));
  assert.deepEqual(run.names, ['parado.docx'], 'a criação (agora) não segura o arquivo, ao contrário de "sem uso"');
  assert.equal(run.records[0].retention.criterion, 'opened-or-modified');
  assert.equal(run.records[0].retention.date, OLD.toISOString());
  assert.ok(!exists('parado.docx') && exists('aberto-ontem.docx') && exists('modificado-ontem.docx'));
  if (st.birthtimeMs > Date.now() - DAY) {
    const used = await runFiles(repoDir().dir, policyFor({ criterion: 'used', since: '2020-01-01' }), { deleteMatches: false });
    assert.equal(used.records.length, 0, 'controle: "sem uso" (com a criação) não lista nada');
  }
});

test('arquivo aberto entre a listagem e a exclusão (novo critério) não é excluído', async () => {
  const { dir, put, exists } = repoDir();
  put('contrato.docx');
  const ownerResolver = {
    async resolve(paths) {
      for (const p of paths) fs.utimesSync(p, new Date(), fs.statSync(p).mtime); // só o último acesso muda
      return new Map();
    },
  };
  const run = await runFiles(dir, policyFor({ criterion: 'opened-or-modified', since: '2020-01-01' }), { ownerResolver });
  assert.ok(exists('contrato.docx'));
  assert.deepEqual(run.events.map((e) => e.status), ['changed']);
  assert.match(run.events[0].error, /deixou de estar expirado/);
});

/** Árvore com nomes parecidos (Fin, Financeiro, Fin-2015, RH/Fin) para conferir os prefixos. */
function finTree() {
  const t = repoDir();
  for (const rel of ['Fin/a.tmp', 'Fin/sub/b.tmp', 'Financeiro/c.tmp', 'Fin-2015/d.tmp', 'RH/Fin/e.tmp', 'raiz.tmp', 'Outros/f.tmp']) t.put(rel);
  return t;
}

test('pastas escolhidas: só elas são varridas e excluídas (sem confundir prefixos parecidos)', async () => {
  const { dir, exists } = finTree();
  const repos = [{ id: 'r1', name: 'Dados', type: 'local', path: dir }];
  const run = await runFiles(dir, policyFor({ since: '2020-01-01', criterion: 'modified', maxDeletions: 0, folders: [{ repositoryId: 'r1', path: path.join(dir, 'Fin') }] }, repos));
  assert.deepEqual(run.names, ['a.tmp', 'b.tmp']);
  assert.equal(run.stats.filesSeen, 2, 'os arquivos de fora nem são listados');
  assert.deepEqual(run.records.map((r) => r.relativePath.split(path.sep).join('/')).sort(), ['Fin/a.tmp', 'Fin/sub/b.tmp'], 'caminho relativo à raiz do repositório');
  assert.ok(run.records.every((r) => r.retention.folder === 'Fin'));
  assert.ok(run.events.every((e) => e.status === 'deleted' && e.target.repositoryId === 'r1' && e.target.scopeFolder === 'Fin' && e.target.type === 'local'));
  assert.ok(!exists('Fin/a.tmp') && !exists('Fin/sub/b.tmp'));
  for (const rel of ['Financeiro/c.tmp', 'Fin-2015/d.tmp', 'RH/Fin/e.tmp', 'raiz.tmp', 'Outros/f.tmp']) assert.ok(exists(rel), `${rel} continua`);

  // Subpasta e pasta repetida/aninhada; pasta ignorada no repositório continua de fora.
  const t2 = finTree();
  const r2 = [{ id: 'r1', name: 'Dados', type: 'local', path: t2.dir }];
  const sub = await runFiles(t2.dir, policyFor({ since: '2020-01-01', criterion: 'modified', folders: [{ repositoryId: 'r1', path: 'Fin\\sub' }, { repositoryId: 'r1', path: 'RH' }] }, r2), { deleteMatches: false });
  assert.deepEqual(sub.names, ['b.tmp', 'e.tmp']);
  const nested = await runFiles(t2.dir, policyFor({ since: '2020-01-01', criterion: 'modified', folders: [{ repositoryId: 'r1', path: 'Fin' }, { repositoryId: 'r1', path: 'Fin/sub' }] }, r2), { deleteMatches: false, exclude: ['sub'] });
  assert.deepEqual(nested.names, ['a.tmp'], 'sem duplicar; "sub" é ignorada no repositório');
});

test('pasta escolhida que não existe, não é pasta, é ignorada ou é um atalho: nada é varrido no lugar dela', async () => {
  const { dir, put, exists } = finTree();
  const repos = [{ id: 'r1', name: 'Dados', type: 'local', path: dir }];
  const missing = await runFiles(dir, policyFor({ since: '2020-01-01', criterion: 'modified', folders: [{ repositoryId: 'r1', path: 'Nao-existe' }] }, repos));
  assert.equal(missing.stats.filesSeen, 0, 'nunca vale o repositório inteiro');
  assert.equal(missing.stats.deleted, 0);
  assert.equal(missing.stats.gaps, 1);
  assert.match(missing.errors[0].message, /não foi encontrada/);
  assert.ok(exists('raiz.tmp') && exists('Fin/a.tmp'));

  const notDir = await runFiles(dir, policyFor({ since: '2020-01-01', criterion: 'modified', folders: [{ repositoryId: 'r1', path: 'raiz.tmp' }] }, repos), { deleteMatches: false });
  assert.match(notDir.errors[0].message, /não é uma pasta/);
  const ignored = await runFiles(dir, policyFor({ since: '2020-01-01', criterion: 'modified', folders: [{ repositoryId: 'r1', path: 'Fin' }] }, repos), { deleteMatches: false, exclude: ['Fin'] });
  assert.equal(ignored.stats.filesSeen, 0);
  assert.match(ignored.errors[0].message, /ignoradas do repositório/);

  // Atalho (link simbólico/junção) para fora do repositório com arquivos antigos.
  const outside = fs.mkdtempSync(path.join(root, 'fora-'));
  fs.writeFileSync(path.join(outside, 'vitima.tmp'), 'x');
  fs.utimesSync(path.join(outside, 'vitima.tmp'), OLD, OLD);
  fs.symlinkSync(outside, path.join(dir, 'Atalho'), 'dir');
  put('Valida/g.tmp');
  const link = await runFiles(dir, policyFor({ since: '2020-01-01', criterion: 'modified', folders: [{ repositoryId: 'r1', path: 'Atalho' }, { repositoryId: 'r1', path: 'Valida' }] }, repos));
  assert.deepEqual(link.names, ['g.tmp'], 'só a pasta válida');
  assert.match(link.errors[0].message, /atalho, link simbólico ou junção/);
  assert.ok(fs.existsSync(path.join(outside, 'vitima.tmp')), 'nada fora do repositório é tocado');
});

test('pasta trocada por um atalho entre a listagem e a exclusão: recusada pelo caminho real', async () => {
  const { dir, put, exists } = repoDir();
  put('Fin/sub/b.tmp');
  put('Outro/c.tmp');
  const repos = [{ id: 'r1', name: 'Dados', type: 'local', path: dir }];
  // Durante a identificação dos proprietários, "Fin/sub" vira um atalho para "Outro" (dentro do
  // repositório, mas fora das pastas escolhidas): a conferência pela raiz não pega; a das pastas, sim.
  const ownerResolver = {
    async resolve(paths) {
      fs.renameSync(path.join(dir, 'Fin/sub'), path.join(dir, 'Fin/sub-original'));
      fs.symlinkSync(path.join(dir, 'Outro'), path.join(dir, 'Fin/sub'), 'dir');
      fs.copyFileSync(path.join(dir, 'Outro/c.tmp'), path.join(dir, 'Outro/b.tmp'));
      fs.utimesSync(path.join(dir, 'Outro/b.tmp'), OLD, OLD);
      return new Map(paths.map((p) => [p, null]));
    },
  };
  const run = await runFiles(dir, policyFor({ since: '2020-01-01', criterion: 'modified', folders: [{ repositoryId: 'r1', path: 'Fin' }] }, repos), { ownerResolver });
  assert.deepEqual(run.events.map((e) => e.status), ['failed']);
  assert.match(run.events[0].error, /pastas escolhidas/);
  assert.ok(exists('Outro/b.tmp'), 'o arquivo fora das pastas escolhidas não foi excluído');
});

test('deleteFile com pastas escolhidas: fora delas é recusado, também com "excluir mesmo assim"', async () => {
  const { dir, put } = repoDir();
  const inside = put('Fin/a.tmp');
  const outsideFile = put('Outro/b.tmp');
  const within = [path.join(dir, 'Fin')];
  const no = await deleteFile(outsideFile, { root: dir, within, force: true });
  assert.equal(no.status, 'failed');
  assert.match(no.error, /pastas escolhidas/);
  assert.ok(fs.existsSync(outsideFile));
  const prefix = put('Financeiro/c.tmp');
  assert.equal((await deleteFile(prefix, { root: dir, within, force: true })).status, 'failed', 'Fin ≠ Financeiro');
  assert.equal((await deleteFile(inside, { root: dir, within })).status, 'deleted');
  assert.equal(recordInScope({ folders: [{ repositoryId: 'r1', path: 'Fin' }] }, { repositoryId: 'r1', relativePath: 'Financeiro/c.tmp' }), false);
  assert.equal(recordInScope({ folders: [{ repositoryId: 'r1', path: 'Fin' }] }, { repositoryId: 'r2', relativePath: 'qualquer.tmp' }), true, 'outro repositório: sem restrição');
});

// ---------------------------------------------------------------------------------------------
// Motor: OneDrive e SharePoint

test('nuvem: só as pastas escolhidas são listadas e excluídas (com ou sem o nome da biblioteca)', async () => {
  const data = world();
  const old = { modified: '2015-01-01T00:00:00Z', created: '2015-01-01T00:00:00Z' };
  // Itens novos a cada rodada (o simulador tira da lista o que for excluído).
  const reset = () => {
    data.drives['d-fin'].items = [
      folder('f-fin', 'Fin', [file('i-a', 'a.txt', 'x', old)]),
      folder('f-financeiro', 'Financeiro', [file('i-b', 'b.txt', 'x', old)]),
      file('i-raiz', 'raiz.txt', 'x', old),
    ];
  };
  const mock = await startMockApis({ graph: data });
  try {
    const repository = cloudRepo('sharepoint', {
      allowDelete: true,
      deleteMode: 'permanent',
      exclude: [],
      cloud: { scope: 'list', accounts: [], sites: ['https://contoso.sharepoint.com/sites/Financeiro'], exclude: [] },
      keep: { all: null, accounts: [], sites: [] },
    });
    for (const spec of ['fin', 'Documentos Compartilhados/Fin', 'Fantasma']) {
      reset();
      data.driveDeleted = [];
      mock.calls.length = 0;
      const r = sanitizeRetention({ criterion: 'modified', since: '2020-01-01', maxDeletions: 0, folders: [{ repositoryId: repository.id, path: spec }] }, 'files', { cloud: true, targets: [repository] });
      const messages = [];
      const stats = await new Scanner(
        { repositories: [repository], terms: [], options: { deleteMatches: true, checkContent: false }, retention: { ...r, cutoff: cutoffDate(r).toISOString() }, endpoints: mock.endpoints },
        (m) => messages.push(m),
      ).run();
      const records = messages.filter((m) => m.type === 'results').flatMap((m) => m.records);
      assert.ok(!mock.calls.some((c) => c.includes('/items/f-financeiro/children')), `${spec}: a pasta Financeiro não é listada`);
      if (spec === 'Fantasma') {
        assert.equal(records.length, 0);
        assert.equal(stats.gaps, 1);
        const errors = messages.filter((m) => m.type === 'errors').flatMap((m) => m.items);
        assert.ok(errors.some((e) => /não foi encontrada em nenhuma biblioteca/.test(e.message)));
      } else {
        assert.deepEqual(records.map((x) => x.name), ['a.txt'], spec);
        assert.deepEqual(data.driveDeleted.map((d) => d.id), ['i-a'], spec);
        const events = messages.filter((m) => m.type === 'deletions').flatMap((m) => m.items);
        assert.equal(events[0].target.type, 'sharepoint');
        assert.equal(events[0].target.library, 'Documentos Compartilhados');
      }
    }
  } finally {
    await mock.close();
  }
});

// ---------------------------------------------------------------------------------------------
// Agendador: políticas salvas antes continuam valendo; mudar data/pastas suspende a exclusão em curso

test('compatibilidade: política antiga não é dada como "alterada"; mudar a data ou as pastas, sim', async () => {
  const store = await new Store(path.join(root, `data-${Math.random().toString(36).slice(2)}`)).init();
  stores.push(store);
  const { dir } = finTree();
  const repo = store.createRepository({ type: 'local', name: 'Dados', path: dir, exclude: [], allowDelete: true, audit: { enabled: false } });
  const legacy = { criterion: 'modified', amount: 5, unit: 'years', maxDeletions: 1000, deleteMode: 'permanent', patterns: [] };
  const data = { purpose: 'retention', kind: 'files', targetIds: [repo.id], listIds: [] };
  const schedule = store.createSchedule({
    ...data,
    name: 'Antiga',
    options: { resolveOwner: false },
    retention: legacy,
    action: 'delete',
    rule: null,
    period: { type: 'all' },
    keepLast: 0,
    enabled: true,
    history: [],
    runCount: 0,
    countDone: 0,
    nextRunAt: null,
    deleteConfirmation: { by: 'ana', at: new Date().toISOString(), targets: deletionPins('files', [repo]), criteria: deletionCriteria(store, data) },
  });
  const scheduler = new Scheduler({ store, manager: {} });
  const scanOf = (retention) => ({
    scheduleId: schedule.id,
    kind: 'files',
    repositoryIds: [repo.id],
    listIds: [],
    options: { deleteMatches: true },
    scheduleCriteria: store.getSchedule(schedule.id).deleteConfirmation.criteria,
    scheduleEnabled: true,
    // Como o gerenciador grava: a política validada de novo (com os repositórios) mais o corte.
    retention: { ...sanitizeRetention(retention, 'files', { targets: [repo] }), cutoff: new Date().toISOString() },
  });
  assert.equal(scheduler.deletionGuard(scanOf(legacy)), null, 'a política salva antes da mudança continua excluindo');

  const withDate = { criterion: 'modified', since: '2020-01-01', maxDeletions: 1000, deleteMode: 'permanent', patterns: [], folders: [{ repositoryId: repo.id, path: 'Fin' }] };
  store.updateSchedule(schedule.id, { retention: sanitizeRetention(withDate, 'files', { targets: [repo] }) });
  const running = scanOf(withDate);
  assert.equal(scheduler.deletionGuard(running), null);
  store.updateSchedule(schedule.id, { retention: sanitizeRetention({ ...withDate, since: '2021-01-01' }, 'files', { targets: [repo] }) });
  assert.match(scheduler.deletionGuard(running), /alterada/, 'nova data: a execução em curso deixa de excluir');
  store.updateSchedule(schedule.id, { retention: sanitizeRetention({ ...withDate, folders: [{ repositoryId: repo.id, path: 'RH' }] }, 'files', { targets: [repo] }) });
  assert.match(scheduler.deletionGuard(running), /alterada/, 'novas pastas: idem');
});

// ---------------------------------------------------------------------------------------------
// API: política por data fixa e pastas, execução, registro e relatório

test('API: política "sem abertura nem modificação desde" numa pasta, exclusão, registro e exportação', async () => {
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
    const { dir, exists } = finTree();
    const repoId = (await api('POST', '/api/repositories', { name: 'Dados', path: dir, allowDelete: true })).data.id;
    const body = {
      purpose: 'retention',
      kind: 'files',
      name: 'Pasta Fin desde 2020',
      repositoryIds: [repoId],
      retention: { mode: 'date', since: '2020-01-01', criterion: 'opened-or-modified', maxDeletions: 0, folders: [{ repositoryId: repoId, path: path.join(dir, 'Fin') }] },
      options: { resolveOwner: false },
      action: 'delete',
      confirmDelete: 'EXCLUIR',
      rule: null,
    };
    // Validações: pasta fora do repositório, de um repositório não marcado, data no futuro.
    const bad = async (retention, pattern) => {
      const res = await api('POST', '/api/schedules', { ...body, retention: { ...body.retention, ...retention } });
      assert.equal(res.status, 400, JSON.stringify(res.data));
      assert.match(res.data.error, pattern);
    };
    await bad({ folders: [{ repositoryId: repoId, path: path.join(root, 'outro') }] }, /não fica dentro/);
    await bad({ folders: [{ repositoryId: 'nao-marcado', path: 'Fin' }] }, /não está marcado/);
    await bad({ since: '2999-01-01' }, /futuro/);

    const created = await api('POST', '/api/schedules', body);
    assert.equal(created.status, 201, JSON.stringify(created.data));
    assert.deepEqual(created.data.retention.folders, [{ repositoryId: repoId, path: 'Fin' }], 'gravada relativa à raiz');
    assert.equal(created.data.retentionText, 'Arquivos sem abertura nem modificação desde 01/01/2020 — somente na pasta Fin');

    const run = await api('POST', `/api/schedules/${created.data.id}/run`, { confirm: true });
    const scan = await wait(run.data.scan.id);
    assert.equal(scan.status, 'completed', JSON.stringify(scan.log));
    assert.equal(scan.retention.cutoff, new Date(2020, 0, 1).toISOString());
    assert.equal(scan.retention.since, '2020-01-01');
    assert.equal(scan.stats.deleted, 2);
    assert.ok(!exists('Fin/a.tmp') && !exists('Fin/sub/b.tmp') && exists('Financeiro/c.tmp') && exists('raiz.tmp'));

    // Registro geral: cada exclusão com o repositório, a pasta da política e o tamanho.
    await store.flushDeletionLog?.();
    const lines = fs.readFileSync(path.join(store.dataDir, 'exclusoes.ndjson'), 'utf8').trim().split('\n').map((l) => JSON.parse(l));
    assert.equal(lines.length, 2);
    for (const e of lines) {
      assert.equal(e.mode, 'retention');
      assert.equal(e.target.repositoryId, repoId);
      assert.equal(e.target.repositoryName, 'Dados');
      assert.equal(e.target.scopeFolder, 'Fin');
      assert.ok(e.target.size > 0);
    }

    // Exportação: a regra com a data fixa e as pastas, sem "undefined".
    const html = await (await fetch(`${base}/api/scans/${scan.id}/export.html`)).text();
    assert.match(html, /sem abertura nem modificação desde 01\/01\/2020/);
    assert.match(html, /Pastas da política/);
    assert.doesNotMatch(html, /undefined|NaN/);
    const csv = await (await fetch(`${base}/api/scans/${scan.id}/export.csv`)).text();
    assert.doesNotMatch(csv, /undefined|NaN/);
  } finally {
    server.close();
    await scheduler.stop();
    await manager.shutdown?.();
  }
});
