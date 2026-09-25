import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import crypto from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { unzipSync, strFromU8 } from 'fflate';
import { MailScanner } from '../src/mail/scanner.js';
import { folderMatcher, addressMatcher } from '../src/mail/common.js';
import http from 'node:http';
import { pool, request } from '../src/mail/http.js';
import { packUids } from '../src/mail/imap.js';
import { createApp } from '../src/app.js';
import { Store } from '../src/store.js';
import { ScanManager } from '../src/scan/manager.js';
import { PRESETS } from '../src/scan/presets.js';
import { startMockApis } from './helpers/mock-apis.js';
import { startFakeImap } from './helpers/fake-imap.js';

const FIXTURES = path.join(import.meta.dirname, 'fixtures');
const DOCX = fs.readFileSync(path.join(FIXTURES, 'doc.docx')); // contém o CPF 529.982.247-25
const cpf = PRESETS.find((p) => p.id === 'cpf');
const TERMS = [
  { id: 'l:cpf', type: 'regex', value: cpf.value, validator: 'cpf', label: 'CPF', listName: 'LGPD' },
  { id: 'l:sal', type: 'text', value: 'salário', listName: 'RH' },
  { id: 'l:conf', type: 'text', value: 'confidencial', listName: 'RH' },
];

const GRAPH_TENANT = 'contoso.onmicrosoft.com';
const GRAPH_CLIENT = '11111111-2222-3333-4444-555555555555';
const GRAPH_SECRET = 'segredo-super-secreto';

/** Mensagem MIME simples, com anexos opcionais. */
function mail({ subject, from = 'Ana Souza <ana@contoso.com>', to = 'rh@contoso.com', body = '', attachments = [], date = 'Thu, 25 Sep 2026 10:00:00 -0300' }) {
  const head = [`From: ${from}`, `To: ${to}`, `Subject: ${subject}`, `Date: ${date}`, `Message-ID: <${crypto.randomUUID()}@contoso.com>`, 'MIME-Version: 1.0'];
  if (!attachments.length) return Buffer.from([...head, 'Content-Type: text/plain; charset=utf-8', '', body, ''].join('\r\n'), 'utf8');
  const parts = [
    ...head,
    'Content-Type: multipart/mixed; boundary="XYZ"',
    '',
    '--XYZ',
    'Content-Type: text/plain; charset=utf-8',
    '',
    body,
    ...attachments.flatMap((a) => ['--XYZ', `Content-Type: application/octet-stream; name="${a.name}"`, 'Content-Disposition: attachment', 'Content-Transfer-Encoding: base64', '', a.data.toString('base64')]),
    '--XYZ--',
    '',
  ];
  return Buffer.from(parts.join('\r\n'), 'utf8');
}

function graphData() {
  const folders = [
    { id: 'inbox', displayName: 'Caixa de Entrada', wellKnown: 'inbox' },
    { id: 'proj', displayName: 'Projetos', parent: 'inbox' },
    { id: 'trash', displayName: 'Itens Excluídos', wellKnown: 'deleteditems' },
    { id: 'junk', displayName: 'Lixo Eletrônico', wellKnown: 'junkemail' },
    { id: 'pessoal', displayName: 'Pessoal' },
  ];
  return {
    tenant: GRAPH_TENANT,
    clientId: GRAPH_CLIENT,
    secret: GRAPH_SECRET,
    throttleOnce: new Set(['m2']),
    users: [
      {
        id: 'u-ana',
        mail: 'ana@contoso.com',
        displayName: 'Ana Souza',
        folders,
        messages: {
          inbox: [
            { id: 'm1', received: '2026-09-20T10:00:00Z', raw: mail({ subject: 'Folha de salário de setembro', body: 'Segue a folha.' }) },
            { id: 'm2', received: '2026-09-21T10:00:00Z', raw: mail({ subject: 'Contrato', body: 'Veja o anexo.', attachments: [{ name: 'contrato.docx', data: DOCX }] }) },
            { id: 'm3', received: '2026-09-22T10:00:00Z', raw: mail({ subject: 'Almoço', body: 'Nada de mais.' }) },
          ],
          proj: [{ id: 'm4', received: '2023-01-01T10:00:00Z', raw: mail({ subject: 'Projeto antigo', body: 'Documento CONFIDENCIAL.' }) }],
          trash: [{ id: 'm5', received: '2026-09-23T10:00:00Z', raw: mail({ subject: 'Apagada', body: 'confidencial na lixeira' }) }],
          junk: [{ id: 'm6', received: '2026-09-23T10:00:00Z', raw: mail({ subject: 'Spam', body: 'confidencial no spam' }) }],
          pessoal: [{ id: 'm7', received: '2026-09-23T10:00:00Z', raw: mail({ subject: 'Pessoal', body: 'salário pessoal' }) }],
        },
      },
      { id: 'u-bia', mail: 'bia@contoso.com', displayName: 'Bia', upnIsMail: false, folders: [folders[0]], messages: { inbox: [{ id: 'b1', received: '2026-09-20T10:00:00Z', raw: mail({ subject: 'Oi', body: 'CPF 529.982.247-25' }) }] } },
      { id: 'u-sem', mail: 'sem.licenca@contoso.com', displayName: 'Sem licença', noMailbox: true, folders: [], messages: {} },
    ],
  };
}

const graphSource = (extra = {}) => ({
  id: 'src-graph',
  name: 'Microsoft 365',
  type: 'graph',
  scope: 'all',
  mailboxes: [],
  excludeMailboxes: [],
  excludeFolders: ['Pessoal'],
  graph: { tenantId: GRAPH_TENANT, clientId: GRAPH_CLIENT },
  secrets: { clientSecret: GRAPH_SECRET },
  ...extra,
});

let mocks;
let google;
let root;

before(async () => {
  root = fs.mkdtempSync(path.join(os.tmpdir(), 'clean-mail-'));
  const { privateKey, publicKey } = crypto.generateKeyPairSync('rsa', { modulusLength: 2048 });
  google = {
    privateKeyPem: privateKey.export({ type: 'pkcs8', format: 'pem' }),
    data: {
      publicKey,
      admin: 'admin@empresa.com',
      users: [
        {
          mail: 'caio@empresa.com',
          name: 'Caio',
          labels: [
            { id: 'INBOX', name: 'INBOX', type: 'system' },
            { id: 'SPAM', name: 'SPAM', type: 'system' },
            { id: 'TRASH', name: 'TRASH', type: 'system' },
            { id: 'UNREAD', name: 'UNREAD', type: 'system' },
            { id: 'Label_1', name: 'Clientes/2026', type: 'user' },
          ],
          messages: [
            { id: 'g1', labelIds: ['INBOX', 'UNREAD', 'Label_1'], internalDate: Date.parse('2026-09-01'), raw: mail({ subject: 'Planilha', body: 'salário e CPF 529.982.247-25' }) },
            { id: 'g2', labelIds: ['SPAM'], internalDate: Date.parse('2026-09-01'), raw: mail({ subject: 'Spam', body: 'confidencial' }) },
            { id: 'g3', labelIds: ['TRASH'], internalDate: Date.parse('2026-09-01'), raw: mail({ subject: 'Lixo', body: 'confidencial na lixeira' }) },
            { id: 'g4', labelIds: [], internalDate: Date.parse('2026-09-01'), raw: mail({ subject: 'Arquivada', body: 'nada' }) },
          ],
        },
        { mail: 'dora@empresa.com', name: 'Dora', disabled: true, labels: [], messages: [] },
      ],
    },
  };
  mocks = await startMockApis({ graph: graphData(), google: google.data });
});

after(async () => {
  await mocks?.close();
  fs.rmSync(root, { recursive: true, force: true });
});

async function runMail(sources, options = {}) {
  const messages = [];
  const scanner = new MailScanner({ sources, terms: TERMS, options, endpoints: mocks.endpoints }, (m) => messages.push(m));
  const stats = await scanner.run();
  const records = messages.filter((m) => m.type === 'results').flatMap((m) => m.records);
  const errors = messages.filter((m) => m.type === 'errors').flatMap((m) => m.items);
  const logs = messages.filter((m) => m.type === 'log');
  return { stats, records, errors, logs, bySubject: Object.fromEntries(records.map((r) => [r.subject, r])) };
}

test('pastas e endereços ignorados: curingas, sem acentos, subpastas', () => {
  const excluded = folderMatcher(['Lixo Eletronico', 'Caixa de Entrada/Pessoal*', 'Arquivo ?']);
  assert.equal(excluded('Lixo Eletrônico'), true);
  assert.equal(excluded('Caixa de Entrada/Pessoal/2024'), true, 'subpasta de pasta ignorada');
  assert.equal(excluded('Caixa de Entrada/Pessoal2'), true);
  assert.equal(excluded('Pessoal'), false, 'padrão com barra compara o caminho');
  assert.equal(excluded('Arquivo 1/Sub'), true);
  assert.equal(excluded('Caixa de Entrada'), false);
  const addresses = addressMatcher(['noreply@*', 'Ana@Contoso.com']);
  assert.equal(addresses('noreply@x.com'), true);
  assert.equal(addresses('ana@contoso.com'), true);
  assert.equal(addresses('bia@contoso.com'), false);
});

test('pool: limita as execuções simultâneas e entrega tudo', async () => {
  let running = 0;
  let peak = 0;
  async function* source() {
    for (let i = 0; i < 20; i++) yield i;
  }
  const out = [];
  for await (const v of pool(source(), 3, async (i) => {
    running++;
    peak = Math.max(peak, running);
    await new Promise((r) => setTimeout(r, 5 + (i % 3) * 5));
    running--;
    return i % 5 === 0 ? undefined : i;
  })) {
    out.push(v);
  }
  assert.equal(peak, 3);
  assert.deepEqual(out.sort((a, b) => a - b), [1, 2, 3, 4, 6, 7, 8, 9, 11, 12, 13, 14, 16, 17, 18, 19]);
});

test('UIDs do IMAP compactados em intervalos', () => {
  assert.equal(packUids([9, 1, 2, 3, 3, 5, 10, 12]), '1:3,5,9:10,12');
  assert.equal(packUids(Array.from({ length: 200000 }, (_, i) => i + 1)), '1:200000');
});

test('HTTP: nova tentativa em falha temporária e tempo limite por inatividade no download', async () => {
  let calls = 0;
  const server = http.createServer((req, res) => {
    calls++;
    if (req.url === '/instavel' && calls === 1) {
      res.writeHead(503, { 'Retry-After': '1', 'Content-Type': 'application/json' });
      return res.end(JSON.stringify({ error: { code: 'ServiceUnavailable', message: 'tente depois' } }));
    }
    if (req.url === '/lento') {
      // Download lento, mas sem pausas maiores que o tempo limite: não pode ser interrompido.
      res.writeHead(200, { 'Content-Type': 'application/octet-stream' });
      let n = 0;
      const timer = setInterval(() => {
        res.write(Buffer.alloc(1000, 65));
        if (++n === 6) {
          clearInterval(timer);
          res.end();
        }
      }, 150);
      return undefined;
    }
    if (req.url === '/parado') {
      res.writeHead(200, { 'Content-Type': 'application/octet-stream' });
      res.write('abc'); // e não envia mais nada
      return undefined;
    }
    res.writeHead(200, { 'Content-Type': 'application/json' });
    return res.end('{"ok":true}');
  });
  await new Promise((resolve) => server.listen(0, '127.0.0.1', resolve));
  const base = `http://127.0.0.1:${server.address().port}`;
  try {
    const retries = [];
    assert.deepEqual(await request(`${base}/instavel`, { onRetry: (r) => retries.push(r.error.status) }), { ok: true });
    assert.deepEqual(retries, [503]);
    const slow = await request(`${base}/lento`, { type: 'buffer', timeoutMs: 400 });
    assert.equal(slow.data.length, 6000, 'o total (900 ms) passa do limite, mas cada pausa não');
    const cut = await request(`${base}/lento`, { type: 'buffer', maxBytes: 2500 });
    assert.equal(cut.data.length, 2500);
    assert.equal(cut.truncated, true);
    await assert.rejects(request(`${base}/parado`, { type: 'buffer', timeoutMs: 300, retries: 0 }), /Tempo esgotado/);
  } finally {
    server.closeAllConnections();
    server.close();
  }
});

test('Microsoft 365: todas as caixas, pastas, anexos, lixeira, limite de taxa e usuário sem caixa', async () => {
  const { stats, records, errors, logs, bySubject } = await runMail([graphSource()]);
  assert.deepEqual(errors, []);
  assert.equal(stats.mailboxesTotal, 3);
  assert.equal(stats.mailboxesSkipped, 1);
  assert.ok(logs.some((l) => /sem\.licenca@contoso\.com ignorada/.test(l.message)));
  assert.ok(logs.some((l) => /limite de requisições/.test(l.message)), 'o 429 foi repetido após o Retry-After');
  // Lixeira incluída por padrão; lixo eletrônico e a pasta "Pessoal" (ignorada) não.
  assert.deepEqual(Object.keys(bySubject).sort(), ['Apagada', 'Contrato', 'Folha de salário de setembro', 'Oi', 'Projeto antigo']);
  assert.equal(stats.messagesSeen, 6);
  const folha = bySubject['Folha de salário de setembro'];
  assert.equal(folha.mailbox, 'ana@contoso.com');
  assert.equal(folha.mailboxName, 'Ana Souza');
  assert.equal(folha.folder, 'Caixa de Entrada');
  assert.equal(folha.from, 'Ana Souza <ana@contoso.com>');
  assert.deepEqual(folha.matches.map((m) => [m.term, m.location]), [['salário', 'subject']]);
  assert.match(folha.webLink, /^https:\/\/outlook\.office365\.com\//);
  const contrato = bySubject.Contrato;
  const found = contrato.matches.find((m) => m.term === 'CPF');
  assert.equal(found.location, 'attachment');
  assert.match(found.samples[0].where, /^Anexo "contrato\.docx"/);
  assert.equal(contrato.attachments[0].name, 'contrato.docx');
  assert.equal(contrato.attachments[0].status, 'ok');
  assert.equal(bySubject['Projeto antigo'].folder, 'Caixa de Entrada/Projetos');
  assert.equal(bySubject.Oi.mailbox, 'bia@contoso.com', 'caixa encontrada pelo endereço quando o UPN é diferente');
  assert.ok(records.every((r) => r.kind === 'mail' && r.sourceType === 'graph'));
});

test('Microsoft 365: data inicial, lixeira ignorada e caixas da lista', async () => {
  const { records } = await runMail([graphSource({ scope: 'list', mailboxes: [{ address: 'ana@contoso.com' }], excludeFolders: [] })], {
    receivedAfter: '2026-01-01T00:00:00Z',
    includeTrash: false,
    includeJunk: true,
  });
  assert.deepEqual(records.map((r) => r.subject).sort(), ['Contrato', 'Folha de salário de setembro', 'Pessoal', 'Spam']);
});

test('Microsoft 365: segredo errado vira erro claro da conexão', async () => {
  const { errors, stats } = await runMail([graphSource({ secrets: { clientSecret: 'errado' } })]);
  assert.equal(stats.mailboxesTotal, 0);
  assert.equal(errors.length, 1);
  assert.match(errors[0].message, /Segredo do cliente inválido/);
});

test('Google Workspace: conta de serviço, marcadores como pastas, spam e usuário sem Gmail', async () => {
  const source = {
    id: 'src-google',
    name: 'Google',
    type: 'gmail',
    scope: 'all',
    mailboxes: [],
    excludeMailboxes: [],
    excludeFolders: [],
    gmail: { clientEmail: 'clean@projeto.iam.gserviceaccount.com', adminEmail: 'admin@empresa.com' },
    secrets: { privateKey: google.privateKeyPem },
  };
  const { records, errors, stats, logs, bySubject } = await runMail([source]);
  assert.deepEqual(errors, []);
  assert.equal(stats.mailboxesSkipped, 1);
  assert.ok(logs.some((l) => /dora@empresa\.com ignorada: Gmail não habilitado/.test(l.message)));
  assert.deepEqual(records.map((r) => r.subject).sort(), ['Lixo', 'Planilha'], 'spam fica de fora por padrão');
  assert.equal(bySubject.Planilha.folder, 'Caixa de entrada; Clientes/2026');
  assert.equal(bySubject.Lixo.folder, 'Lixeira');
  assert.deepEqual(bySubject.Planilha.terms.sort(), ['CPF', 'salário']);
  assert.equal(bySubject.Planilha.date, '2026-09-01T00:00:00.000Z');
  assert.equal(stats.messagesSeen, 3);

  const bad = await runMail([{ ...source, secrets: { privateKey: crypto.generateKeyPairSync('rsa', { modulusLength: 2048 }).privateKey.export({ type: 'pkcs8', format: 'pem' }) } }]);
  assert.match(bad.errors[0].message, /Autorização recusada pelo Google/);
});

test('IMAP: senha padrão e individual, falha de login, mensagem grande e pasta de spam', async () => {
  const big = mail({ subject: 'Anexo enorme', body: 'Relatório confidencial.', attachments: [{ name: 'grande.docx', data: Buffer.concat([DOCX, crypto.randomBytes(1_200_000)]) }] });
  const imap = await startFakeImap({
    'carla@empresa.com': {
      password: 'padrao',
      folders: {
        INBOX: [
          { raw: mail({ subject: 'Salário', body: 'O salário foi pago.' }), date: new Date('2026-09-10T12:00:00Z') },
          { raw: big, date: new Date('2026-09-11T12:00:00Z') },
        ],
        Spam: [{ raw: mail({ subject: 'Promoção', body: 'confidencial' }), date: new Date('2026-09-12T12:00:00Z') }],
        'Clientes/Antigos': [{ raw: mail({ subject: 'Velho', body: 'confidencial' }), date: new Date('2020-01-01T12:00:00Z') }],
      },
      special: { Spam: '\\Junk' },
    },
    'DOMINIO\\svc\\davi': { password: 'individual', folders: { INBOX: [{ raw: mail({ subject: 'CPF do Davi', body: 'CPF 529.982.247-25' }), date: new Date('2026-09-10T12:00:00Z') }] } },
  });
  try {
    const source = {
      id: 'src-imap',
      name: 'Servidor interno',
      type: 'imap',
      scope: 'list',
      imap: { host: '127.0.0.1', port: imap.port, security: 'none' },
      mailboxes: [{ address: 'carla@empresa.com' }, { address: 'davi@empresa.com', login: 'DOMINIO\\svc\\davi' }, { address: 'erro@empresa.com' }],
      excludeMailboxes: [],
      excludeFolders: [],
      secrets: { defaultPassword: 'padrao', passwords: { 'davi@empresa.com': 'individual' } },
    };
    const { records, errors, stats, bySubject } = await runMail([source], { maxMessageSizeMB: 1 });
    assert.deepEqual(records.map((r) => r.subject).sort(), ['Anexo enorme', 'CPF do Davi', 'Salário', 'Velho']);
    assert.equal(bySubject['CPF do Davi'].mailbox, 'davi@empresa.com');
    assert.equal(bySubject.Velho.folder, 'Clientes/Antigos');
    const enorme = bySubject['Anexo enorme'];
    assert.equal(enorme.contentStatus, 'partial');
    assert.match(enorme.contentNote, /apenas os primeiros/);
    assert.equal(enorme.attachments[0].status, 'skipped-size', 'o anexo cortado não é lido');
    assert.equal(stats.messagesPartial, 1);
    assert.equal(errors.length, 1);
    assert.equal(errors[0].path, 'erro@empresa.com');
    assert.match(errors[0].message, /Usuário ou senha recusados/);

    const recent = await runMail([{ ...source, mailboxes: [{ address: 'carla@empresa.com' }] }], { receivedAfter: '2026-01-01T00:00:00Z', includeJunk: true });
    assert.deepEqual(recent.records.map((r) => r.subject).sort(), ['Anexo enorme', 'Promoção', 'Salário']);
  } finally {
    await imap.close();
  }
});

// ---------------------------------------------------------------------------------------------
// API: cadastro com segredos cifrados, teste da conexão e análise completa em worker thread

async function startApp() {
  const store = await new Store(path.join(root, `data-${Math.random().toString(36).slice(2)}`)).init();
  const manager = new ScanManager(store, { mailEndpoints: mocks.endpoints });
  const app = createApp({ store, manager, config: { authUser: '', authPassword: '', mailEndpoints: mocks.endpoints } });
  const srv = await new Promise((resolve) => {
    const x = app.listen(0, '127.0.0.1', () => resolve(x));
  });
  const base = `http://127.0.0.1:${srv.address().port}`;
  const api = async (method, url, body) => {
    const res = await fetch(base + url, { method, headers: { 'Content-Type': 'application/json', 'X-CLEAN': '1' }, body: body === undefined ? undefined : JSON.stringify(body) });
    const type = res.headers.get('content-type') || '';
    return { status: res.status, data: type.includes('json') ? await res.json() : Buffer.from(await res.arrayBuffer()), headers: res.headers };
  };
  return { store, manager, srv, api, close: async () => (srv.close(), await manager.shutdown()) };
}

test('API: segredos cifrados, nunca devolvidos, e mantidos quando o campo fica vazio', async () => {
  const app = await startApp();
  try {
    const bad = await app.api('POST', '/api/mail-sources', { name: 'X', type: 'graph', scope: 'all', graph: { tenantId: GRAPH_TENANT, clientId: 'não-é-guid', clientSecret: 's' } });
    assert.equal(bad.status, 400);
    const created = await app.api('POST', '/api/mail-sources', {
      name: 'Microsoft 365',
      type: 'graph',
      scope: 'all',
      excludeFolders: 'Pessoal',
      graph: { tenantId: GRAPH_TENANT, clientId: GRAPH_CLIENT, clientSecret: GRAPH_SECRET },
    });
    assert.equal(created.status, 201);
    assert.equal(created.data.graph.hasClientSecret, true);
    const listed = await app.api('GET', '/api/mail-sources');
    const everything = JSON.stringify([created.data, listed.data]);
    assert.ok(!everything.includes(GRAPH_SECRET) && !everything.includes('enc:v1'), 'o segredo não sai pela API');
    await app.store.saveNow();
    const db = fs.readFileSync(path.join(app.store.dataDir, 'db.json'), 'utf8');
    assert.ok(!db.includes(GRAPH_SECRET) && db.includes('enc:v1:'), 'o segredo é gravado cifrado');

    const tested = await app.api('POST', '/api/mail-sources/test', { id: created.data.id, type: 'graph', scope: 'all', graph: { tenantId: GRAPH_TENANT, clientId: GRAPH_CLIENT } });
    assert.equal(tested.data.ok, true, JSON.stringify(tested.data));
    assert.ok(tested.data.details.some((d) => /ana@contoso\.com: 3 mensagem/.test(d)), JSON.stringify(tested.data));
    const wrong = await app.api('POST', '/api/mail-sources/test', { type: 'graph', scope: 'all', graph: { tenantId: GRAPH_TENANT, clientId: GRAPH_CLIENT, clientSecret: 'x' } });
    assert.equal(wrong.data.ok, false);
    assert.match(wrong.data.message, /Segredo do cliente inválido/);

    const updated = await app.api('PUT', `/api/mail-sources/${created.data.id}`, { ...created.data, graph: { tenantId: GRAPH_TENANT, clientId: GRAPH_CLIENT, clientSecret: '' } });
    assert.equal(updated.status, 200);
    assert.equal(updated.data.graph.hasClientSecret, true, 'campo vazio mantém o segredo salvo');

    // IMAP: trocar o servidor descarta as senhas salvas (não são enviadas a outro endereço).
    const imapSource = await app.api('POST', '/api/mail-sources', {
      name: 'IMAP',
      type: 'imap',
      imap: { host: 'imap.empresa.com', security: 'tls' },
      mailboxes: [{ address: 'a@empresa.com', password: 'p1' }],
    });
    assert.equal(imapSource.status, 201);
    assert.equal(imapSource.data.imap.port, 993);
    assert.equal(imapSource.data.mailboxes[0].hasPassword, true);
    const sameHost = await app.api('PUT', `/api/mail-sources/${imapSource.data.id}`, { ...imapSource.data, mailboxes: [{ address: 'a@empresa.com' }] });
    assert.equal(sameHost.status, 200);
    const otherHost = await app.api('PUT', `/api/mail-sources/${imapSource.data.id}`, { ...imapSource.data, imap: { host: 'outro.exemplo.com' }, mailboxes: [{ address: 'a@empresa.com' }] });
    assert.equal(otherHost.status, 400);
    assert.match(otherHost.data.error, /Informe a senha da caixa a@empresa\.com/);
    // Porta, segurança ou certificado não confiável também exigem a senha de novo (nada de enviar
    // a senha salva sem criptografia ou para quem apresentar qualquer certificado).
    for (const imap of [
      { host: 'imap.empresa.com', security: 'tls', port: 1993 },
      { host: 'imap.empresa.com', security: 'none' },
      { host: 'imap.empresa.com', security: 'tls', allowSelfSigned: true },
    ]) {
      const changed = await app.api('PUT', `/api/mail-sources/${imapSource.data.id}`, { ...imapSource.data, imap, mailboxes: [{ address: 'a@empresa.com' }] });
      assert.equal(changed.status, 400, JSON.stringify(imap));
      const tested = await app.api('POST', '/api/mail-sources/test', { ...imapSource.data, id: imapSource.data.id, imap, mailboxes: [{ address: 'a@empresa.com' }] });
      assert.equal(tested.status, 400, 'o teste também não usa a senha salva');
    }
  } finally {
    await app.close();
  }
});

test('API: análise de e-mail completa em segundo plano, relatório e exportações', async () => {
  const app = await startApp();
  try {
    const source = await app.api('POST', '/api/mail-sources', {
      name: 'Microsoft 365',
      type: 'graph',
      scope: 'list',
      mailboxes: 'ana@contoso.com\nbia@contoso.com',
      excludeFolders: 'Pessoal',
      graph: { tenantId: GRAPH_TENANT, clientId: GRAPH_CLIENT, clientSecret: GRAPH_SECRET },
    });
    const list = await app.api('POST', '/api/lists', { name: 'Sensíveis', terms: TERMS.map(({ id, listName, ...t }) => t) });
    assert.equal(list.status, 201);
    const invalid = await app.api('POST', '/api/scans', { kind: 'mail', sourceIds: [source.data.id], listIds: [list.data.id], options: { checkSubject: false, checkBody: false, checkAttachments: false, checkAttachmentNames: false } });
    assert.equal(invalid.status, 400);
    const scan = await app.api('POST', '/api/scans', { kind: 'mail', name: 'Varredura de e-mail', sourceIds: [source.data.id], listIds: [list.data.id] });
    assert.equal(scan.status, 201);
    assert.equal(scan.data.kind, 'mail');
    let current;
    for (let i = 0; i < 300; i++) {
      current = await app.api('GET', `/api/scans/${scan.data.id}`);
      if (!['queued', 'running'].includes(current.data.status)) break;
      await new Promise((r) => setTimeout(r, 100));
    }
    assert.equal(current.data.status, 'completed', JSON.stringify(current.data.log));
    assert.equal(current.data.stats.messagesMatched, 5);
    const config = fs.readFileSync(path.join(app.store.scanDir(scan.data.id), 'config.json'), 'utf8');
    assert.ok(!config.includes(GRAPH_SECRET) && !config.includes('enc:v1'), 'a configuração da análise não guarda segredos');

    const onlyMail = await app.api('GET', '/api/scans?kind=mail');
    assert.deepEqual(onlyMail.data.map((s) => s.id), [scan.data.id]);
    const results = await app.api('GET', `/api/scans/${scan.data.id}/results`);
    assert.equal(results.data.total, 5);
    assert.equal(results.data.items[0].subject, 'Apagada', 'mais recentes primeiro');
    const bia = await app.api('GET', `/api/scans/${scan.data.id}/results?mailbox=bia@contoso.com`);
    assert.deepEqual(bia.data.items.map((r) => r.subject), ['Oi']);
    const inAttachments = await app.api('GET', `/api/scans/${scan.data.id}/results?location=attachment`);
    assert.deepEqual(inAttachments.data.items.map((r) => r.subject), ['Contrato']);
    const search = await app.api('GET', `/api/scans/${scan.data.id}/results?q=contrato.docx`);
    assert.deepEqual(search.data.items.map((r) => r.subject), ['Contrato']);
    const summary = await app.api('GET', `/api/scans/${scan.data.id}/summary`);
    assert.equal(summary.data.messages, 5);
    assert.deepEqual(summary.data.byMailbox.map((m) => [m.mailbox, m.messages]), [['ana@contoso.com', 4], ['bia@contoso.com', 1]]);
    assert.deepEqual(summary.data.options.mailboxes, ['ana@contoso.com', 'bia@contoso.com']);
    assert.ok(summary.data.options.senders.some((s) => s.value === 'ana@contoso.com'));

    const csv = await app.api('GET', `/api/scans/${scan.data.id}/export.csv`);
    const text = csv.data.toString('utf8');
    assert.ok(text.startsWith('\uFEFFCaixa;Pasta;Data;Remetente;Assunto;Termo'));
    assert.ok(text.includes('Conteúdo do anexo'));
    const xlsx = await app.api('GET', `/api/scans/${scan.data.id}/export.xlsx`);
    const parts = unzipSync(new Uint8Array(xlsx.data));
    const workbook = strFromU8(parts['xl/workbook.xml']);
    assert.deepEqual([...workbook.matchAll(/<sheet name="([^"]*)"/g)].map((m) => m[1]), ['Resumo', 'Mensagens', 'Ocorrências']);
    const html = await app.api('GET', `/api/scans/${scan.data.id}/export.html`);
    assert.ok(html.data.toString('utf8').includes('Mensagens com ocorrências (5)'));
    const json = await app.api('GET', `/api/scans/${scan.data.id}/export.json`);
    assert.equal(json.data.results.length, 5);
  } finally {
    await app.close();
  }
});
