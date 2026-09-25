// Salvaguardas da exclusão: conferência antes de excluir, pastas protegidas, permissões conferidas
// de novo, registro imediato (inclusive no cancelamento), pedidos simultâneos e servidores IMAP sem
// os recursos necessários para excluir com segurança.
import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import crypto from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { Scanner } from '../src/scan/scanner.js';
import { MailScanner } from '../src/mail/scanner.js';
import { ImapConnector } from '../src/mail/imap.js';
import { deleteFile, isInside, isWithin, cleanPaths } from '../src/scan/delete.js';
import { applyDeletions } from '../src/report/model.js';
import { actor } from '../src/routes/scans.js';
import { createApp } from '../src/app.js';
import { Store } from '../src/store.js';
import { ScanManager } from '../src/scan/manager.js';
import { startMockApis } from './helpers/mock-apis.js';
import { startFakeImap } from './helpers/fake-imap.js';

const TERMS = [
  { id: 'l:conf', type: 'text', value: 'confidencial', listName: 'RH' },
  { id: 'l:sal', type: 'text', value: 'salário', listName: 'RH' },
];
const LIST_TERMS = TERMS.map(({ id, listName, ...t }) => t);
const TENANT = 'contoso.onmicrosoft.com';
const CLIENT = '11111111-2222-3333-4444-555555555555';
const SECRET = 'segredo';

let root;
const cleanups = [];

before(() => {
  root = fs.mkdtempSync(path.join(os.tmpdir(), 'clean-delete-safety-'));
});
after(async () => {
  for (const fn of cleanups.reverse()) await fn();
  fs.rmSync(root, { recursive: true, force: true });
});

function makeRepo(name, files) {
  const dir = path.join(root, name);
  for (const [rel, content] of Object.entries(files)) {
    const file = path.join(dir, rel);
    fs.mkdirSync(path.dirname(file), { recursive: true });
    fs.writeFileSync(file, content, 'utf8');
  }
  return dir;
}

function mail({ subject, body, id = `${crypto.randomUUID()}@x` }) {
  return Buffer.from(
    [`From: Ana <ana@contoso.com>`, 'To: rh@contoso.com', `Subject: ${subject}`, `Message-ID: <${id}>`, 'Content-Type: text/plain; charset=utf-8', '', body, ''].join('\r\n'),
    'utf8',
  );
}

async function startApp(mailEndpoints = {}) {
  const store = await new Store(path.join(root, `data-${crypto.randomUUID()}`)).init();
  const manager = new ScanManager(store, { mailEndpoints });
  const app = createApp({ store, manager, config: { authUser: '', authPassword: '', mailEndpoints } });
  const srv = await new Promise((resolve) => {
    const x = app.listen(0, '127.0.0.1', () => resolve(x));
  });
  const base = `http://127.0.0.1:${srv.address().port}`;
  const api = async (method, url, body, headers = {}) => {
    const res = await fetch(base + url, {
      method,
      headers: { 'Content-Type': 'application/json', 'X-CLEAN': '1', ...headers },
      body: body === undefined ? undefined : JSON.stringify(body),
    });
    const type = res.headers.get('content-type') || '';
    return { status: res.status, data: type.includes('json') ? await res.json() : Buffer.from(await res.arrayBuffer()) };
  };
  const wait = async (id) => {
    let scan;
    for (let i = 0; i < 300; i++) {
      scan = (await api('GET', `/api/scans/${id}`)).data;
      if (!['queued', 'running'].includes(scan.status)) return scan;
      await new Promise((r) => setTimeout(r, 100));
    }
    return scan;
  };
  cleanups.push(async () => {
    srv.close();
    await manager.shutdown();
  });
  return { store, manager, api, wait };
}

const readLines = (file) =>
  fs
    .readFileSync(file, 'utf8')
    .split('\n')
    .filter(Boolean)
    .map((l) => JSON.parse(l));

// -- arquivos -------------------------------------------------------------------------------------

test('arquivo: nomes que começam com "..", pasta trocada por um link para fora do repositório', async () => {
  const dir = makeRepo('real', { 'sub/relatorio.txt': 'x', '..backup.txt': 'y' });
  assert.equal(isInside(dir, path.join(dir, '..backup.txt')), true);
  assert.equal(isInside(dir, path.join(dir, '..', 'fora.txt')), false);
  assert.equal(isWithin(dir, dir), true);
  const outside = makeRepo('fora', { 'relatorio.txt': 'x' });
  const st = fs.statSync(path.join(dir, 'sub', 'relatorio.txt'));
  // Depois da análise, a pasta "sub" vira um link (junção, no Windows) para outra pasta com um
  // arquivo de mesmo nome, tamanho e data: a exclusão não pode sair do repositório.
  fs.rmSync(path.join(dir, 'sub'), { recursive: true });
  fs.symlinkSync(outside, path.join(dir, 'sub'), 'junction');
  const result = await deleteFile(path.join(dir, 'sub', 'relatorio.txt'), { root: dir, expected: { size: st.size, modified: st.mtime.toISOString() } });
  assert.equal(result.status, 'failed');
  assert.match(result.error, /aponta para fora do repositório/);
  assert.equal(fs.existsSync(path.join(outside, 'relatorio.txt')), true);
  assert.equal((await deleteFile(path.join(dir, '..backup.txt'), { root: dir })).status, 'deleted');
});

test('arquivo: pastas do próprio CLEAN nunca têm arquivos excluídos (menos a demonstração)', async () => {
  const app = makeRepo('app', { 'data/db.json': 'confidencial', 'demo/Compartilhamento/a.txt': 'confidencial', 'src/x.js': 'confidencial' });
  const guards = cleanPaths({ dataDir: path.join(app, 'data'), appDir: app });
  const data = await deleteFile(path.join(app, 'data', 'db.json'), { root: app, protect: guards });
  assert.equal(data.status, 'failed');
  assert.match(data.error, /dados do CLEAN/);
  const code = await deleteFile(path.join(app, 'src', 'x.js'), { root: app, protect: guards });
  assert.match(code.error, /instalação do CLEAN/);

  // A análise não entra na pasta de dados; a instalação é analisada, mas não excluída.
  const messages = [];
  const stats = await new Scanner(
    { repositories: [{ id: 'r', name: 'App', path: app, allowDelete: true }], terms: TERMS, options: { deleteMatches: true, resolveOwner: false }, protect: guards },
    (m) => messages.push(m),
  ).run();
  const found = messages.filter((m) => m.type === 'results').flatMap((m) => m.records.map((r) => path.relative(app, r.path).replaceAll('\\', '/')));
  assert.deepEqual(found.sort(), ['demo/Compartilhamento/a.txt', 'src/x.js']);
  assert.equal(stats.deleted, 1);
  assert.equal(stats.deleteErrors, 1);
  assert.equal(fs.existsSync(path.join(app, 'src', 'x.js')), true);
  assert.equal(fs.existsSync(path.join(app, 'data', 'db.json')), true);
  assert.equal(fs.existsSync(path.join(app, 'demo', 'Compartilhamento', 'a.txt')), false);
  assert.ok(messages.some((m) => m.type === 'log' && /Pasta de dados do CLEAN ignorada/.test(m.message)));
});

test('exclusão automática: arquivo alterado depois de analisado fica; cada exclusão é registrada na hora', async () => {
  const dir = makeRepo('mudou', { 'a.txt': 'confidencial', 'b.txt': 'confidencial também' });
  const messages = [];
  // O proprietário é identificado depois da leitura: o arquivo muda nesse meio-tempo.
  const ownerResolver = {
    resolve: async () => {
      fs.writeFileSync(path.join(dir, 'a.txt'), 'texto novo, sem o termo, e maior que o anterior');
      return new Map();
    },
  };
  const stats = await new Scanner(
    { repositories: [{ id: 'r', name: 'R', path: dir, allowDelete: true }], terms: TERMS, options: { deleteMatches: true }, startedBy: 'ana (acesso local)' },
    (m) => messages.push(m),
    { ownerResolver },
  ).run();
  assert.equal(stats.deleted, 1);
  assert.equal(stats.deleteChanged, 1);
  assert.equal(stats.deleteErrors, 0);
  assert.equal(fs.existsSync(path.join(dir, 'a.txt')), true);
  assert.equal(fs.existsSync(path.join(dir, 'b.txt')), false);
  const batches = messages.filter((m) => m.type === 'deletions');
  assert.ok(batches.every((b) => b.items.length === 1), 'um evento por arquivo, logo depois da exclusão');
  assert.ok(messages.findIndex((m) => m.type === 'results') < messages.findIndex((m) => m.type === 'deletions'));
  const events = batches.flatMap((b) => b.items);
  assert.deepEqual(events.map((e) => e.status).sort(), ['changed', 'deleted']);
  assert.ok(events.every((e) => e.item.startsWith(dir) && e.by === 'ana (acesso local)' && e.mode === 'auto'));
});

test('exclusão automática: permissão retirada durante a análise interrompe as exclusões', async () => {
  const dir = makeRepo('revoga', { 'a.txt': 'confidencial', 'b.txt': 'salário' });
  const messages = [];
  let scanner;
  const ownerResolver = {
    resolve: async () => {
      scanner.revokeDeletion({ kind: 'repository', id: 'r', reason: 'a opção "Permitir exclusão" foi desligada' });
      return new Map();
    },
  };
  scanner = new Scanner({ repositories: [{ id: 'r', name: 'R', path: dir, allowDelete: true }], terms: TERMS, options: { deleteMatches: true } }, (m) => messages.push(m), { ownerResolver });
  const stats = await scanner.run();
  assert.equal(stats.deleted, 0);
  assert.equal(stats.deleteErrors, 2);
  assert.equal(fs.readdirSync(dir).length, 2);
  assert.ok(messages.some((m) => m.type === 'log' && /Exclusão automática desativada para "R"/.test(m.message)));
});

test('na fila: permissão retirada, repositório aninhado sem permissão e forma de exclusão confirmada', async () => {
  const store = await new Store(path.join(root, `data-${crypto.randomUUID()}`)).init();
  const manager = new ScanManager(store);
  manager.maxConcurrent = 0; // mantém as análises na fila
  const dir = makeRepo('fila', { 'Juridico/c.txt': 'confidencial', 'x.txt': 'confidencial' });
  const dados = store.createRepository({ name: 'Dados', path: dir, exclude: [], allowDelete: true, audit: {} });
  const juridico = store.createRepository({ name: 'Jurídico', path: path.join(dir, 'Juridico'), exclude: [], allowDelete: false, audit: {} });
  const list = store.createList({ name: 'L', terms: LIST_TERMS.map((t, i) => ({ ...t, id: `t${i}` })) });
  const scan = await manager.start({ repositoryIds: [dados.id], listIds: [list.id], options: { deleteMatches: true }, confirmDelete: 'EXCLUIR' }, { by: 'ana (acesso local)' });
  assert.equal(scan.startedBy, 'ana (acesso local)');
  let config = await manager.workerConfig(scan.id);
  assert.equal(config.repositories[0].allowDelete, true);
  assert.deepEqual(config.repositories[0].keep.map((k) => k.path), [juridico.path]);
  assert.equal(config.startedBy, 'ana (acesso local)');
  assert.ok(config.protect.some((p) => p.path === store.dataDir));

  // O repositório aninhado sem permissão protege os arquivos dele.
  const kept = await deleteFile(path.join(dir, 'Juridico', 'c.txt'), { root: dir, protect: config.repositories[0].keep });
  assert.equal(kept.status, 'failed');
  assert.match(kept.error, /repositório "Jurídico", que não permite exclusão/);

  // A permissão foi retirada enquanto a análise esperava na fila.
  store.updateRepository(dados.id, { allowDelete: false });
  config = await manager.workerConfig(scan.id);
  assert.equal(config.repositories[0].allowDelete, false);
  assert.match(store.getScan(scan.id).log.map((l) => l.message).join('\n'), /Exclusão automática desativada para "Dados"/);

  // "deleteMatches" só com o valor exato.
  const typo = await manager.start({ repositoryIds: [juridico.id], listIds: [list.id], options: { deleteMatches: 'false' } });
  assert.equal(typo.options.deleteMatches, false);

  // E-mail: vale a forma confirmada ao criar a análise (ou a lixeira, se o cadastro mudou para ela).
  const source = store.createMailSource({ name: 'M', type: 'graph', scope: 'list', mailboxes: [{ address: 'a@x.com' }], graph: { tenantId: 't', clientId: 'c' }, secrets: {}, allowDelete: true, deleteMode: 'trash' });
  const mailScan = await manager.start({ kind: 'mail', sourceIds: [source.id], listIds: [list.id], options: { deleteMatches: true }, confirmDelete: 'EXCLUIR' });
  store.updateMailSource(source.id, { deleteMode: 'permanent' });
  const mailConfig = await manager.workerConfig(mailScan.id);
  assert.equal(mailConfig.sources[0].deleteMode, 'trash');
  assert.equal(mailConfig.sources[0].allowDelete, true);
  await manager.shutdown();
});

test('relatório: "excluído" não é trocado por uma tentativa posterior; quem fez a ação', () => {
  const records = [{ id: 1 }, { id: 2 }];
  applyDeletions(records, [
    { recordId: 1, status: 'deleted' },
    { recordId: 1, status: 'missing' },
    { recordId: 2, status: 'failed' },
    { recordId: 2, status: 'deleted' },
  ]);
  assert.deepEqual(
    records.map((r) => r.deletion.status),
    ['deleted', 'deleted'],
  );
  assert.equal(actor({ cleanUser: 'admin', ip: '10.0.0.5' }), 'admin (acesso de 10.0.0.5)');
  assert.equal(actor({ ip: '::ffff:127.0.0.1' }), 'acesso local');
});

test('API: pedidos simultâneos, forma de exclusão mostrada, proxy, registro geral e falha ao registrar', async () => {
  const dir = makeRepo('api-seguranca', { 'a.txt': 'confidencial', 'b.txt': 'confidencial', 'c.txt': 'confidencial' });
  const { store, api, wait } = await startApp();
  const repo = (await api('POST', '/api/repositories', { name: 'R', path: dir, allowDelete: true })).data;
  const list = (await api('POST', '/api/lists', { name: 'L', terms: LIST_TERMS })).data;
  const scan = (await api('POST', '/api/scans', { name: 'Relatório A', repositoryIds: [repo.id], listIds: [list.id] })).data;
  assert.equal((await wait(scan.id)).status, 'completed');
  const [a, b, c] = (await api('GET', `/api/scans/${scan.id}/results?sort=name`)).data.items;
  const url = (item, id = scan.id) => `/api/scans/${id}/results/${item.id}/delete`;

  // Dois pedidos ao mesmo tempo (duas abas, dois usuários): um exclui, o outro é recusado.
  const both = await Promise.all([api('POST', url(a), { confirm: true }), api('POST', url(a), { confirm: true })]);
  assert.deepEqual(both.map((r) => r.status).sort(), [200, 409]);

  // A forma de exclusão mostrada na confirmação precisa ser a do cadastro.
  const other = await api('POST', url(b), { confirm: true, method: 'trash' });
  assert.equal(other.status, 409);
  assert.equal(other.data.code, 'method-changed');
  assert.equal(fs.existsSync(path.join(dir, 'b.txt')), true);

  // Atrás de um proxy na mesma máquina, fica registrado o endereço do navegador.
  const proxied = await api('POST', url(b), { confirm: true, method: 'file' }, { 'X-Forwarded-For': '10.1.2.3' });
  assert.equal(proxied.status, 200);
  assert.equal(proxied.data.deletion.by, 'acesso de 10.1.2.3');
  const summary = (await api('GET', `/api/scans/${scan.id}/summary`)).data;
  assert.deepEqual(summary.deletions, { deleted: 2, missing: 0, changed: 0, failed: 0 });

  // Registro geral: continua existindo depois que o relatório é excluído.
  assert.equal((await api('DELETE', `/api/scans/${scan.id}`)).status, 204);
  const general = readLines(path.join(store.dataDir, 'exclusoes.ndjson'));
  assert.equal(general.length, 2);
  assert.ok(general.every((e) => e.scanId === scan.id && e.scanName === 'Relatório A' && e.status === 'deleted' && e.item.startsWith(dir)));

  // Falha ao gravar o registro: a resposta diz o que aconteceu e a ação fica no Registro da análise.
  const scan2 = (await api('POST', '/api/scans', { repositoryIds: [repo.id], listIds: [list.id] })).data;
  assert.equal((await wait(scan2.id)).status, 'completed');
  const [c2] = (await api('GET', `/api/scans/${scan2.id}/results`)).data.items;
  assert.equal(c2.name, 'c.txt');
  fs.renameSync(path.join(store.dataDir, 'exclusoes.ndjson'), path.join(store.dataDir, 'exclusoes-antigo.ndjson'));
  fs.mkdirSync(path.join(store.dataDir, 'exclusoes.ndjson')); // o registro geral não pode ser gravado
  const unrecorded = await api('POST', url(c2, scan2.id), { confirm: true });
  assert.equal(unrecorded.status, 500);
  assert.match(unrecorded.data.error, /Resultado: excluído\. Mas houve uma falha ao gravar o registro da exclusão/);
  assert.equal(fs.existsSync(path.join(dir, 'c.txt')), false);
  const log = (await api('GET', `/api/scans/${scan2.id}`)).data.log.map((l) => l.message).join('\n');
  assert.match(log, /c\.txt por acesso local: excluído\. Falha ao gravar o registro da exclusão/);
  assert.ok(c);
});

// -- e-mail: IMAP ---------------------------------------------------------------------------------

const imapAccount = ({ trashMessage = false } = {}) => ({
  password: 'p',
  folders: {
    INBOX: [
      { raw: mail({ subject: 'A', body: 'confidencial' }), date: new Date('2026-09-10') },
      { raw: mail({ subject: 'B', body: 'nada' }), date: new Date('2026-09-10'), deleted: true }, // marcada pelo usuário
      { raw: mail({ subject: 'C', body: 'salário' }), date: new Date('2026-09-10') },
    ],
    Lixeira: trashMessage ? [{ raw: mail({ subject: 'D', body: 'confidencial' }), date: new Date('2026-09-10') }] : [],
  },
  special: { Lixeira: '\\Trash' },
});

const imapSource = (port, mode, extra = {}) => ({
  id: 'i',
  name: 'IMAP',
  type: 'imap',
  scope: 'list',
  imap: { host: '127.0.0.1', port, security: 'none' },
  mailboxes: [{ address: 'ana@empresa.com' }],
  excludeMailboxes: [],
  excludeFolders: [],
  secrets: { defaultPassword: 'p' },
  allowDelete: true,
  deleteMode: mode,
  ...extra,
});

async function runMail(sources, endpoints, options = {}, deps = {}) {
  const messages = [];
  const scanner = new MailScanner({ sources, terms: TERMS, options: { deleteMatches: true, ...options }, endpoints }, (m) => messages.push(m), deps);
  const stats = await scanner.run();
  return {
    stats,
    messages,
    records: messages.filter((m) => m.type === 'results').flatMap((m) => m.records),
    events: messages.filter((m) => m.type === 'deletions').flatMap((m) => m.items),
    errors: messages.filter((m) => m.type === 'errors').flatMap((m) => m.items),
  };
}

async function imapCase(serverOptions, mode, { account = imapAccount(), includeTrash = false } = {}) {
  const accounts = { 'ana@empresa.com': account };
  const imap = await startFakeImap(accounts, serverOptions);
  try {
    const run = await runMail([imapSource(imap.port, mode, { excludeFolders: includeTrash ? [] : ['Lixeira'] })], {}, { includeTrash });
    const folders = accounts['ana@empresa.com'].folders;
    return { run, inbox: folders.INBOX.map((m) => ({ subject: /Subject: (.*)/.exec(m.raw.toString())[1], deleted: Boolean(m.deleted) })), trash: folders.Lixeira };
  } finally {
    await imap.close();
  }
}

test('IMAP sem UIDPLUS: exclusão definitiva recusada; mover para a Lixeira só com MOVE', async () => {
  const refused = await imapCase({ capabilities: 'IMAP4rev1 MOVE' }, 'permanent');
  assert.equal(refused.run.stats.deleteErrors, 2);
  assert.match(refused.run.errors[0].message, /UIDPLUS/);
  assert.deepEqual(refused.inbox, [
    { subject: 'A', deleted: false },
    { subject: 'B', deleted: true },
    { subject: 'C', deleted: false },
  ]);

  const moved = await imapCase({ capabilities: 'IMAP4rev1 MOVE' }, 'trash');
  assert.equal(moved.run.stats.deleted, 2);
  assert.deepEqual(moved.inbox, [{ subject: 'B', deleted: true }]);
  assert.equal(moved.trash.length, 2);

  const neither = await imapCase({ capabilities: 'IMAP4rev1' }, 'trash');
  assert.equal(neither.run.stats.deleteErrors, 2);
  assert.equal(neither.inbox.length, 3);
});

test('IMAP sem MOVE: cópia conferida antes de expurgar; cópia recusada não apaga nada', async () => {
  const copied = await imapCase({ capabilities: 'IMAP4rev1 UIDPLUS' }, 'trash');
  assert.equal(copied.run.stats.deleted, 2);
  assert.deepEqual(copied.inbox, [{ subject: 'B', deleted: true }], 'a mensagem marcada pelo usuário continua lá');
  assert.equal(copied.trash.length, 2);

  const quota = await imapCase({ capabilities: 'IMAP4rev1 UIDPLUS', refuse: { copy: true } }, 'trash');
  assert.equal(quota.run.stats.deleted, 0);
  assert.equal(quota.run.stats.deleteErrors, 2);
  assert.match(quota.run.errors[0].message, /recusou copiar/);
  assert.deepEqual(quota.inbox.map((m) => m.subject), ['A', 'B', 'C']);
  assert.equal(quota.trash.length, 0);
});

test('IMAP: marcação recusada ou não permitida não conta como exclusão (e é desfeita)', async () => {
  const refused = await imapCase({ refuse: { store: true } }, 'permanent');
  assert.equal(refused.run.stats.deleted, 0);
  assert.equal(refused.run.stats.deleteErrors, 2);
  assert.equal(refused.inbox.length, 3);

  const notPermanent = await imapCase({ permanentFlags: '\\Seen' }, 'permanent');
  assert.equal(notPermanent.run.stats.deleted, 0);
  assert.match(notPermanent.run.errors[0].message, /recusou marcar/);
  assert.equal(notPermanent.inbox.length, 3);

  const noExpunge = await imapCase({ refuse: { expunge: true } }, 'permanent');
  assert.equal(noExpunge.run.stats.deleted, 0);
  assert.deepEqual(noExpunge.inbox, [
    { subject: 'A', deleted: false },
    { subject: 'B', deleted: true },
    { subject: 'C', deleted: false },
  ]);
});

test('IMAP "mover para a Lixeira": mensagem que já está na Lixeira fica lá', async () => {
  const run = await imapCase({}, 'trash', { account: imapAccount({ trashMessage: true }), includeTrash: true });
  assert.equal(run.run.stats.deleted, 3);
  const inTrash = run.run.events.find((e) => /› D$/.test(e.item));
  assert.equal(inTrash.status, 'deleted');
  assert.equal(inTrash.note, 'já estava na lixeira');
  assert.equal(run.trash.length, 3);
});

test('IMAP: confere o Message-ID e a UIDVALIDITY antes de excluir', async () => {
  const account = imapAccount();
  const accounts = { 'ana@empresa.com': account };
  const imap = await startFakeImap(accounts);
  try {
    const connector = new ImapConnector(imapSource(imap.port, 'permanent'));
    const box = { address: 'ana@empresa.com' };
    const other = await connector.deleteMessages(box, [{ id: 'INBOX:7:1', messageId: '<outra@x>' }], 'permanent');
    assert.equal(other.get('INBOX:7:1').ok, false);
    assert.match(other.get('INBOX:7:1').error, /Message-ID diferente/);
    assert.equal(account.folders.INBOX.length, 3);
    const same = /Message-ID: (.*)/.exec(account.folders.INBOX[0].raw.toString())[1];
    const ok = await connector.deleteMessages(box, [{ id: 'INBOX:7:1', messageId: same }], 'permanent');
    assert.equal(ok.get('INBOX:7:1').ok, true);
    assert.equal(account.folders.INBOX.length, 2);
    account.validity = { INBOX: 8 };
    const recreated = await connector.deleteMessages(box, ['INBOX:7:3'], 'permanent');
    assert.equal(recreated.get('INBOX:7:3').ok, false);
    assert.equal(recreated.get('INBOX:7:3').missing, undefined, 'não é "não encontrada": a mensagem pode continuar lá');
    assert.match(recreated.get('INBOX:7:3').error, /UIDVALIDITY/);
    assert.equal(account.folders.INBOX.length, 2);
  } finally {
    await imap.close();
  }
});

test('Gmail por IMAP: a exclusão definitiva passa pela Lixeira (expurgar fora dela só arquivaria)', async () => {
  const account = () => ({
    password: 'p',
    folders: {
      '[Gmail]/Todos os e-mails': [
        { raw: mail({ subject: 'A', body: 'confidencial' }), date: new Date('2026-09-10') },
        { raw: mail({ subject: 'B', body: 'nada' }), date: new Date('2026-09-10') },
      ],
      '[Gmail]/Lixeira': [],
    },
    special: { '[Gmail]/Todos os e-mails': '\\All', '[Gmail]/Lixeira': '\\Trash' },
  });
  const accounts = { 'ana@empresa.com': account() };
  const imap = await startFakeImap(accounts, { capabilities: 'IMAP4rev1 UIDPLUS MOVE X-GM-EXT-1', gmail: true });
  try {
    const run = await runMail([imapSource(imap.port, 'permanent', { excludeFolders: [] })], {}, { includeTrash: false });
    assert.equal(run.stats.deleted, 1);
    assert.equal(run.events[0].note, null);
    const folders = accounts['ana@empresa.com'].folders;
    assert.equal(folders['[Gmail]/Todos os e-mails'].length, 1);
    assert.equal(folders['[Gmail]/Lixeira'].length, 0, 'excluída de vez, e não só arquivada ou na lixeira');

    accounts['ana@empresa.com'] = account();
    for (const list of Object.values(accounts['ana@empresa.com'].folders)) list.forEach((m, i) => (m.uid = i + 1));
    const trash = await runMail([imapSource(imap.port, 'trash', { excludeFolders: [] })], {}, { includeTrash: false });
    assert.equal(trash.stats.deleted, 1);
    assert.equal(accounts['ana@empresa.com'].folders['[Gmail]/Lixeira'].length, 1);
  } finally {
    await imap.close();
  }
});

test('IMAP pela API: a exclusão manual usa o login configurado para a caixa', async () => {
  const accounts = { 'EMPRESA\\svc\\ana': imapAccount() };
  const imap = await startFakeImap(accounts);
  cleanups.push(() => imap.close());
  const { api, wait } = await startApp();
  const source = (
    await api('POST', '/api/mail-sources', {
      name: 'Exchange',
      type: 'imap',
      imap: { host: '127.0.0.1', port: imap.port, security: 'none', defaultPassword: 'p' },
      mailboxes: [{ address: 'ana@empresa.com', login: 'EMPRESA\\svc\\ana' }],
      excludeFolders: 'Lixeira',
      allowDelete: true,
      deleteMode: 'permanent',
    })
  ).data;
  assert.ok(source.id, JSON.stringify(source));
  const list = (await api('POST', '/api/lists', { name: 'L', terms: LIST_TERMS })).data;
  const scan = (await api('POST', '/api/scans', { kind: 'mail', sourceIds: [source.id], listIds: [list.id] })).data;
  assert.equal((await wait(scan.id)).status, 'completed');
  const items = (await api('GET', `/api/scans/${scan.id}/results`)).data.items;
  assert.equal(items.length, 2);
  const res = await api('POST', `/api/scans/${scan.id}/results/${items[0].id}/delete`, { confirm: true, method: 'permanent' });
  assert.equal(res.status, 200, JSON.stringify(res.data));
  assert.equal(res.data.deletion.status, 'deleted', res.data.deletion.error);
  assert.equal(accounts['EMPRESA\\svc\\ana'].folders.INBOX.length, 2);
  const log = (await api('GET', `/api/scans/${scan.id}`)).data.log.map((l) => l.message).join('\n');
  assert.match(log, /Exclusão manual da mensagem ".*" da caixa ana@empresa\.com por acesso local: excluída\./);
});

// -- e-mail: Microsoft 365 e Google ----------------------------------------------------------------

function graphUsers() {
  const folders = [
    { id: 'inbox', displayName: 'Caixa de Entrada', wellKnown: 'inbox' },
    { id: 'trash', displayName: 'Itens Excluídos', wellKnown: 'deleteditems' },
  ];
  return [
    {
      id: 'u1',
      mail: 'ana@contoso.com',
      displayName: 'Ana',
      folders,
      messages: {
        inbox: [
          { id: 'm1', received: '2026-09-20T10:00:00Z', raw: mail({ subject: 'Folha', body: 'salário' }) },
          { id: 'm2', received: '2026-09-20T10:00:00Z', raw: mail({ subject: 'Oi', body: 'nada' }) },
        ],
      },
    },
  ];
}

const graphSource = (extra = {}) => ({
  id: 'g',
  name: 'M365',
  type: 'graph',
  scope: 'list',
  mailboxes: [{ address: 'ana@contoso.com' }],
  excludeMailboxes: [],
  excludeFolders: ['Itens Excluídos'],
  graph: { tenantId: TENANT, clientId: CLIENT },
  secrets: { clientSecret: SECRET },
  allowDelete: true,
  ...extra,
});

test('Microsoft 365 e Google: resposta perdida não vira "não encontrada"; 404 da caixa é falha', async () => {
  const data = { tenant: TENANT, clientId: CLIENT, secret: SECRET, users: graphUsers(), flakyDelete: true };
  const { privateKey, publicKey } = crypto.generateKeyPairSync('rsa', { modulusLength: 2048 });
  const google = {
    publicKey,
    admin: 'admin@empresa.com',
    flakyDelete: true,
    users: [
      {
        mail: 'caio@empresa.com',
        name: 'Caio',
        labels: [{ id: 'INBOX', name: 'INBOX', type: 'system' }],
        messages: [{ id: 'g1', labelIds: ['INBOX'], internalDate: Date.parse('2026-09-01'), raw: mail({ subject: 'A', body: 'salário' }) }],
      },
    ],
  };
  const mock = await startMockApis({ graph: data, google });
  try {
    // A 1ª exclusão é feita, mas a resposta se perde (503): a repetição encontra 404.
    const run = await runMail([graphSource()], mock.endpoints);
    assert.deepEqual(
      run.events.map((e) => e.status),
      ['deleted'],
    );
    assert.equal(run.stats.deleteMissing, 0);

    const gmailSource = {
      id: 'gg',
      name: 'Google',
      type: 'gmail',
      scope: 'list',
      mailboxes: [{ address: 'caio@empresa.com' }],
      excludeMailboxes: [],
      excludeFolders: [],
      gmail: { clientEmail: 'svc@p.iam.gserviceaccount.com' },
      secrets: { privateKey: privateKey.export({ type: 'pkcs8', format: 'pem' }) },
      allowDelete: true,
    };
    const gmail = await runMail([gmailSource], mock.endpoints);
    assert.deepEqual(
      gmail.events.map((e) => e.status),
      ['deleted'],
    );
    assert.equal(google.users[0].messages.length, 0);

    // Um 404 que não é da mensagem (caixa desativada) é falha, e não "não encontrada".
    data.users = graphUsers();
    data.flakyDelete = false;
    data.deleteError = { status: 404, code: 'MailboxNotEnabledForRESTAPI', message: 'The mailbox is either inactive, soft-deleted, or is hosted on-premise.' };
    const gone = await runMail([graphSource()], mock.endpoints);
    assert.deepEqual(
      gone.events.map((e) => e.status),
      ['failed'],
    );
    assert.equal(gone.stats.deleted + gone.stats.deleteMissing, 0);
  } finally {
    await mock.close();
  }
});

// -- e-mail: registro no cancelamento, mensagens repetidas e permissão retirada --------------------

function stubConnector({ ids, onDelete }) {
  return () => ({
    mailboxes: async () => [{ address: 'a@x.com', name: '' }],
    async *messages() {
      for (const [i, id] of ids.entries()) yield { folder: 'Entrada', id, raw: mail({ subject: `S${i}`, body: 'confidencial' }), truncated: false, size: 100, receivedAt: null };
    },
    deleteMessages: onDelete,
    close: async () => {},
  });
}

const stubSource = { id: 's', name: 'Stub', type: 'graph', scope: 'list', allowDelete: true, deleteMode: 'permanent' };

test('e-mail: cancelar durante a exclusão registra o que já foi excluído; mensagens repetidas são excluídas uma vez', async () => {
  let scanner;
  let received = null;
  const onDelete = async (mailbox, items, mode, { signal, onResult, shouldStop }) => {
    received = { items, signal };
    for (const [i, item] of items.entries()) {
      if (shouldStop()) break;
      await new Promise((r) => setTimeout(r, 5));
      onResult(item.id, { ok: true });
      if (i === 1) scanner.cancel(); // o usuário cancela no meio das exclusões
    }
  };
  const messages = [];
  scanner = new MailScanner(
    { sources: [stubSource], terms: TERMS, options: { deleteMatches: true } },
    (m) => messages.push(m),
    { connectorFactory: stubConnector({ ids: ['m0', 'm1', 'm1', 'm2', 'm3'], onDelete }) },
  );
  const stats = await scanner.run();
  assert.deepEqual(
    received.items.map((i) => i.id),
    ['m0', 'm1', 'm2', 'm3'],
    'a mensagem repetida é enviada uma vez',
  );
  assert.equal(received.signal, null, 'o cancelamento não interrompe as exclusões em andamento');
  assert.ok(received.items.every((i) => /@x>$/.test(i.messageId)));
  const events = messages.filter((m) => m.type === 'deletions').flatMap((m) => m.items);
  assert.equal(events.length, 3, 'm0 e as duas ocorrências de m1');
  assert.ok(events.every((e) => e.status === 'deleted'));
  assert.equal(stats.deleted, 3);
  assert.ok(messages.some((m) => m.type === 'log' && /Cancelado: 2 mensagem\(ns\) da caixa a@x\.com não foram excluídas/.test(m.message)));
});

test('e-mail: permissão retirada durante a exclusão interrompe as seguintes', async () => {
  let scanner;
  const onDelete = async (mailbox, items, mode, { onResult, shouldStop }) => {
    for (const item of items) {
      if (shouldStop()) break;
      onResult(item.id, { ok: true });
      scanner.revokeDeletion({ kind: 'mail', id: 's', reason: 'a opção "Permitir exclusão" foi desligada' });
    }
  };
  const messages = [];
  scanner = new MailScanner(
    { sources: [{ ...stubSource }], terms: TERMS, options: { deleteMatches: true } },
    (m) => messages.push(m),
    { connectorFactory: stubConnector({ ids: ['m0', 'm1', 'm2'], onDelete }) },
  );
  const stats = await scanner.run();
  assert.equal(stats.deleted, 1);
  assert.equal(stats.deleteErrors, 2);
  const failed = messages.filter((m) => m.type === 'deletions').flatMap((m) => m.items).filter((e) => e.status === 'failed');
  assert.ok(failed.every((e) => /desativada no cadastro/.test(e.error)));
});
