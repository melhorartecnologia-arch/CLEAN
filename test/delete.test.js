import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import crypto from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { unzipSync, strFromU8 } from 'fflate';
import { Scanner } from '../src/scan/scanner.js';
import { MailScanner } from '../src/mail/scanner.js';
import { deleteFile, isInside } from '../src/scan/delete.js';
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
  root = fs.mkdtempSync(path.join(os.tmpdir(), 'clean-delete-'));
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

function mail({ subject, body }) {
  return Buffer.from(
    [`From: Ana <ana@contoso.com>`, 'To: rh@contoso.com', `Subject: ${subject}`, `Message-ID: <${crypto.randomUUID()}@x>`, 'Content-Type: text/plain; charset=utf-8', '', body, ''].join('\r\n'),
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
  const api = async (method, url, body) => {
    const res = await fetch(base + url, { method, headers: { 'Content-Type': 'application/json', 'X-CLEAN': '1' }, body: body === undefined ? undefined : JSON.stringify(body) });
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
  return { store, api, wait };
}

test('exclusão de arquivo: dentro do repositório, arquivo alterado e já excluído', async () => {
  const dir = makeRepo('unit', { 'a.txt': 'x', 'b.txt': 'y' });
  assert.equal(isInside(dir, path.join(dir, 'sub', 'a.txt')), true);
  assert.equal(isInside(dir, path.join(dir, '..', 'fora.txt')), false);
  assert.equal(isInside(dir, dir), false);
  assert.equal((await deleteFile(path.join(root, 'fora.txt'), { root: dir })).status, 'failed');
  const st = fs.statSync(path.join(dir, 'a.txt'));
  const expected = { size: st.size, modified: st.mtime.toISOString() };
  fs.writeFileSync(path.join(dir, 'a.txt'), 'conteúdo novo e maior');
  assert.equal((await deleteFile(path.join(dir, 'a.txt'), { root: dir, expected })).status, 'changed');
  assert.equal(fs.existsSync(path.join(dir, 'a.txt')), true, 'arquivo alterado não é excluído sem confirmação');
  assert.equal((await deleteFile(path.join(dir, 'a.txt'), { root: dir, expected, force: true })).status, 'deleted');
  assert.equal((await deleteFile(path.join(dir, 'a.txt'), { root: dir })).status, 'missing');
  assert.equal((await deleteFile(path.join(dir, 'b.txt'), { root: dir })).status, 'deleted');
});

test('arquivos: analisar e excluir automaticamente o que foi encontrado', async () => {
  const dir = makeRepo('auto', { 'RH/folha.txt': 'salário de todos', 'RH/limpo.txt': 'nada aqui', 'confidencial.txt': 'ok', 'Outros/memo.txt': 'material CONFIDENCIAL' });
  const messages = [];
  const scanner = new Scanner(
    { repositories: [{ id: 'r', name: 'Auto', path: dir, allowDelete: true }], terms: TERMS, options: { deleteMatches: true, resolveOwner: false } },
    (m) => messages.push(m),
  );
  const stats = await scanner.run();
  assert.equal(stats.filesMatched, 3);
  assert.equal(stats.deleted, 3);
  assert.equal(stats.deleteErrors, 0);
  assert.deepEqual(fs.readdirSync(dir, { recursive: true }).filter((f) => f.endsWith('.txt')).map((f) => f.replaceAll('\\', '/')), ['RH/limpo.txt']);
  const events = messages.filter((m) => m.type === 'deletions').flatMap((m) => m.items);
  assert.equal(events.length, 3);
  assert.ok(events.every((e) => e.status === 'deleted' && e.mode === 'auto' && e.method === 'file'));
});

test('arquivos pela API: exigências, exclusão manual item a item, filtros e exportações', async () => {
  const dir = makeRepo('manual', { 'a.txt': 'confidencial A', 'b.txt': 'confidencial B', 'c.txt': 'salário C', 'd.txt': 'nada' });
  const { store, api, wait } = await startApp();
  const locked = (await api('POST', '/api/repositories', { name: 'Sem exclusão', path: dir })).data;
  assert.equal(locked.allowDelete, false, 'por padrão a exclusão não é permitida');
  const list = (await api('POST', '/api/lists', { name: 'L', terms: LIST_TERMS })).data;

  const refused = await api('POST', '/api/scans', { repositoryIds: [locked.id], listIds: [list.id], options: { deleteMatches: true }, confirmDelete: 'EXCLUIR' });
  assert.equal(refused.status, 400);
  assert.match(refused.data.error, /não está permitida/);

  const scan1 = (await api('POST', '/api/scans', { repositoryIds: [locked.id], listIds: [list.id] })).data;
  assert.equal((await wait(scan1.id)).status, 'completed');
  const items1 = (await api('GET', `/api/scans/${scan1.id}/results?sort=name`)).data.items;
  assert.ok(items1.every((r) => r.canDelete === false));
  const denied = await api('POST', `/api/scans/${scan1.id}/results/${items1[0].id}/delete`, { confirm: true });
  assert.equal(denied.status, 403);
  assert.equal(fs.existsSync(path.join(dir, 'a.txt')), true);

  // Permite a exclusão no repositório (o relatório anterior passa a permitir a exclusão manual).
  await api('PUT', `/api/repositories/${locked.id}`, { ...locked, allowDelete: true });
  const noConfirm = await api('POST', '/api/scans', { repositoryIds: [locked.id], listIds: [list.id], options: { deleteMatches: true } });
  assert.equal(noConfirm.status, 400);
  assert.match(noConfirm.data.error, /digite EXCLUIR/);

  const items = (await api('GET', `/api/scans/${scan1.id}/results?sort=name`)).data.items;
  const byName = Object.fromEntries(items.map((r) => [r.name, r]));
  assert.equal(byName['a.txt'].canDelete, true);
  assert.equal(byName['a.txt'].deleteMethod, 'file');
  const needConfirm = await api('POST', `/api/scans/${scan1.id}/results/${byName['a.txt'].id}/delete`, {});
  assert.equal(needConfirm.status, 400);
  const ok = await api('POST', `/api/scans/${scan1.id}/results/${byName['a.txt'].id}/delete`, { confirm: true });
  assert.equal(ok.status, 200);
  assert.equal(ok.data.deletion.status, 'deleted');
  assert.equal(ok.data.deletion.by, 'acesso local');
  assert.equal(fs.existsSync(path.join(dir, 'a.txt')), false);
  const again = await api('POST', `/api/scans/${scan1.id}/results/${byName['a.txt'].id}/delete`, { confirm: true });
  assert.equal(again.status, 409);

  // Arquivo alterado depois da análise: pede confirmação (force) antes de excluir.
  fs.writeFileSync(path.join(dir, 'b.txt'), 'confidencial B, agora com mais texto');
  const changed = await api('POST', `/api/scans/${scan1.id}/results/${byName['b.txt'].id}/delete`, { confirm: true });
  assert.equal(changed.status, 409);
  assert.equal(changed.data.code, 'changed');
  assert.equal(fs.existsSync(path.join(dir, 'b.txt')), true);
  const forced = await api('POST', `/api/scans/${scan1.id}/results/${byName['b.txt'].id}/delete`, { confirm: true, force: true });
  assert.equal(forced.data.deletion.status, 'deleted');

  const deleted = (await api('GET', `/api/scans/${scan1.id}/results?deletion=deleted&sort=name`)).data;
  assert.deepEqual(deleted.items.map((r) => r.name), ['a.txt', 'b.txt']);
  assert.ok(deleted.items.every((r) => r.canDelete === false && r.deletion.mode === 'manual'));
  const kept = (await api('GET', `/api/scans/${scan1.id}/results?deletion=kept`)).data;
  assert.deepEqual(kept.items.map((r) => r.name), ['c.txt']);
  const summary = (await api('GET', `/api/scans/${scan1.id}/summary`)).data;
  assert.deepEqual(summary.deletions, { deleted: 2, missing: 0, changed: 0, failed: 0 });
  const scanLog = (await api('GET', `/api/scans/${scan1.id}`)).data.log.map((l) => l.message).join('\n');
  assert.match(scanLog, /Exclusão manual do arquivo .*a\.txt por acesso local: excluído/);

  const xlsx = unzipSync(new Uint8Array((await api('GET', `/api/scans/${scan1.id}/export.xlsx`)).data));
  const sheets = [...strFromU8(xlsx['xl/workbook.xml']).matchAll(/<sheet name="([^"]*)"/g)].map((m) => m[1]);
  assert.ok(sheets.includes('Exclusões'), sheets.join(', '));
  const csv = (await api('GET', `/api/scans/${scan1.id}/export.csv`)).data.toString('utf8');
  assert.match(csv, /Excluído em .* — exclusão manual por acesso local/);

  // Analisar e excluir: com a confirmação digitada, exclui o que restou com ocorrências.
  const scan2 = (await api('POST', '/api/scans', { repositoryIds: [locked.id], listIds: [list.id], options: { deleteMatches: true }, confirmDelete: 'excluir' })).data;
  const done = await wait(scan2.id);
  assert.equal(done.status, 'completed');
  assert.equal(done.stats.deleted, 1);
  assert.deepEqual(fs.readdirSync(dir).sort(), ['d.txt']);
  const auto = (await api('GET', `/api/scans/${scan2.id}/results`)).data.items;
  assert.equal(auto[0].deletion.mode, 'auto');
  const busy = store.getScan(scan2.id);
  assert.equal(busy.options.deleteMatches, true);
});

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
          { id: 'm3', received: '2026-09-20T10:00:00Z', raw: mail({ subject: 'Memo', body: 'confidencial' }) },
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

async function runMail(sources, endpoints, options = {}) {
  const messages = [];
  const stats = await new MailScanner({ sources, terms: TERMS, options: { deleteMatches: true, ...options }, endpoints }, (m) => messages.push(m)).run();
  return {
    stats,
    records: messages.filter((m) => m.type === 'results').flatMap((m) => m.records),
    events: messages.filter((m) => m.type === 'deletions').flatMap((m) => m.items),
    errors: messages.filter((m) => m.type === 'errors').flatMap((m) => m.items),
  };
}

test('Microsoft 365: exclusão automática definitiva, para a lixeira e sem permissão', async () => {
  const data = { tenant: TENANT, clientId: CLIENT, secret: SECRET, users: graphUsers() };
  const mock = await startMockApis({ graph: data });
  try {
    const run = await runMail([graphSource()], mock.endpoints);
    assert.equal(run.records.length, 2);
    assert.equal(run.stats.deleted, 2);
    assert.deepEqual(data.deleted.map((d) => [d.id, d.how]).sort(), [['m1', 'permanentDelete'], ['m3', 'permanentDelete']]);
    assert.ok(data.deleted.every((d) => d.prefer.includes('ImmutableId')), 'identificadores imutáveis');
    assert.deepEqual(data.users[0].messages.inbox.map((m) => m.id), ['m2']);
    const byRecord = new Map(run.records.map((r) => [r.id, r]));
    assert.ok(run.events.every((e) => e.status === 'deleted' && e.method === 'permanent' && byRecord.has(e.recordId)));

    data.users = graphUsers();
    data.deleted = [];
    const trash = await runMail([graphSource({ deleteMode: 'trash' })], mock.endpoints);
    assert.equal(trash.stats.deleted, 2);
    assert.deepEqual(data.users[0].messages.trash.map((m) => m.id).sort(), ['m1', 'm3']);

    data.users = graphUsers();
    data.users[0].readOnly = true;
    const denied = await runMail([graphSource()], mock.endpoints);
    assert.equal(denied.stats.deleteErrors, 2);
    assert.match(denied.errors[0].message, /Mail\.ReadWrite/);
    assert.ok(denied.events.every((e) => e.status === 'failed'));
  } finally {
    await mock.close();
  }
});

test('Google Workspace: exclusão definitiva e escopo de exclusão não autorizado', async () => {
  const { privateKey, publicKey } = crypto.generateKeyPairSync('rsa', { modulusLength: 2048 });
  const users = () => [
    {
      mail: 'caio@empresa.com',
      name: 'Caio',
      labels: [{ id: 'INBOX', name: 'INBOX', type: 'system' }],
      messages: [
        { id: 'g1', labelIds: ['INBOX'], internalDate: Date.parse('2026-09-01'), raw: mail({ subject: 'A', body: 'salário' }) },
        { id: 'g2', labelIds: ['INBOX'], internalDate: Date.parse('2026-09-01'), raw: mail({ subject: 'B', body: 'nada' }) },
      ],
    },
  ];
  const data = { publicKey, admin: 'admin@empresa.com', users: users() };
  const mock = await startMockApis({ google: data });
  const source = {
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
  try {
    const run = await runMail([source], mock.endpoints);
    assert.equal(run.stats.deleted, 1);
    assert.deepEqual(data.users[0].messages.map((m) => m.id), ['g2']);
    assert.deepEqual(run.events.map((e) => e.status), ['deleted'], 'resposta 204 sem corpo não é repetida');
    assert.equal(data.deleted.length, 1);
    data.users = users();
    data.deniedScopes = new Set(['https://mail.google.com/']);
    const denied = await runMail([source], mock.endpoints);
    assert.equal(denied.stats.deleteErrors, 1);
    assert.match(denied.errors[0].message, /autorize o escopo https:\/\/mail\.google\.com\//);
    assert.equal(data.users[0].messages.length, 2);
  } finally {
    await mock.close();
  }
});

test('IMAP: exclusão definitiva só das mensagens encontradas e para a Lixeira', async () => {
  const account = () => ({
    password: 'p',
    folders: {
      INBOX: [
        { raw: mail({ subject: 'A', body: 'confidencial' }), date: new Date('2026-09-10') },
        { raw: mail({ subject: 'B', body: 'nada' }), date: new Date('2026-09-10'), deleted: true }, // marcada pelo usuário: não pode ser expurgada
        { raw: mail({ subject: 'C', body: 'salário' }), date: new Date('2026-09-10') },
      ],
      Lixeira: [],
    },
    special: { Lixeira: '\\Trash' },
  });
  const accounts = { 'ana@empresa.com': account() };
  const imap = await startFakeImap(accounts);
  const source = (mode) => ({
    id: 'i',
    name: 'IMAP',
    type: 'imap',
    scope: 'list',
    imap: { host: '127.0.0.1', port: imap.port, security: 'none' },
    mailboxes: [{ address: 'ana@empresa.com' }],
    excludeMailboxes: [],
    excludeFolders: ['Lixeira'],
    secrets: { defaultPassword: 'p' },
    allowDelete: true,
    deleteMode: mode,
  });
  try {
    const run = await runMail([source('permanent')], {});
    assert.equal(run.stats.deleted, 2);
    const inbox = accounts['ana@empresa.com'].folders.INBOX;
    assert.deepEqual(inbox.map((m) => m.uid), [2], 'só as mensagens encontradas foram expurgadas');
    accounts['ana@empresa.com'].folders = account().folders;
    for (const list of Object.values(accounts['ana@empresa.com'].folders)) list.forEach((m, i) => (m.uid = i + 1));
    const trash = await runMail([source('trash')], {});
    assert.equal(trash.stats.deleted, 2);
    assert.equal(accounts['ana@empresa.com'].folders.Lixeira.length, 2);
    assert.equal(accounts['ana@empresa.com'].folders.INBOX.length, 1);
  } finally {
    await imap.close();
  }
});

test('e-mail pela API: exclusão manual de uma mensagem e conexão sem permissão', async () => {
  const data = { tenant: TENANT, clientId: CLIENT, secret: SECRET, users: graphUsers() };
  const mock = await startMockApis({ graph: data });
  cleanups.push(() => mock.close());
  const { api, wait } = await startApp(mock.endpoints);
  const base = { name: 'M365', type: 'graph', scope: 'list', mailboxes: 'ana@contoso.com', excludeFolders: 'Itens Excluídos', graph: { tenantId: TENANT, clientId: CLIENT, clientSecret: SECRET } };
  const source = (await api('POST', '/api/mail-sources', base)).data;
  assert.equal(source.allowDelete, false);
  const list = (await api('POST', '/api/lists', { name: 'L', terms: LIST_TERMS })).data;
  const refused = await api('POST', '/api/scans', { kind: 'mail', sourceIds: [source.id], listIds: [list.id], options: { deleteMatches: true }, confirmDelete: 'EXCLUIR' });
  assert.equal(refused.status, 400);
  const scan = (await api('POST', '/api/scans', { kind: 'mail', sourceIds: [source.id], listIds: [list.id] })).data;
  assert.equal((await wait(scan.id)).status, 'completed');
  const items = (await api('GET', `/api/scans/${scan.id}/results`)).data.items;
  assert.equal(items.length, 2);
  assert.equal((await api('POST', `/api/scans/${scan.id}/results/${items[0].id}/delete`, { confirm: true })).status, 403);

  const updated = (await api('PUT', `/api/mail-sources/${source.id}`, { ...base, allowDelete: true, deleteMode: 'trash', graph: { tenantId: TENANT, clientId: CLIENT } })).data;
  assert.equal(updated.allowDelete, true);
  assert.equal(updated.deleteMode, 'trash');
  const fresh = (await api('GET', `/api/scans/${scan.id}/results`)).data.items;
  assert.ok(fresh.every((r) => r.canDelete && r.deleteMethod === 'trash'));
  const target = fresh.find((r) => r.subject === 'Memo');
  const res = await api('POST', `/api/scans/${scan.id}/results/${target.id}/delete`, { confirm: true });
  assert.equal(res.data.deletion.status, 'deleted');
  assert.equal(res.data.deletion.method, 'trash');
  assert.deepEqual(data.users[0].messages.trash.map((m) => m.id), ['m3']);
  const missing = await api('POST', `/api/scans/${scan.id}/results/${fresh.find((r) => r.subject === 'Folha').id}/delete`, { confirm: true });
  assert.equal(missing.data.deletion.status, 'deleted');
  data.users[0].messages.inbox = []; // a mensagem some por fora do CLEAN
  const summary = (await api('GET', `/api/scans/${scan.id}/summary`)).data;
  assert.deepEqual(summary.deletions, { deleted: 2, missing: 0, changed: 0, failed: 0 });
});
