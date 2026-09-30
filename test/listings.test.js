// Listagens de e-mail (somente leitura): catálogo de contas do domínio e listagem de mensagens por
// caixa. Testa os conectores, o motor, a API, o relatório e as exportações com servidores simulados.
import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import crypto from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { unzipSync, strFromU8 } from 'fflate';
import { MailScanner } from '../src/mail/scanner.js';
import { createApp } from '../src/app.js';
import { Store } from '../src/store.js';
import { ScanManager } from '../src/scan/manager.js';
import { startMockApis } from './helpers/mock-apis.js';
import { startFakeImap } from './helpers/fake-imap.js';

const GRAPH_TENANT = 'contoso.onmicrosoft.com';
const GRAPH_CLIENT = '11111111-2222-3333-4444-555555555555';
const GRAPH_SECRET = 'segredo-super-secreto';

function mail({ subject, from = 'Ana Souza <ana@contoso.com>', to = 'rh@contoso.com', cc = '', body = '', attachments = [], date = 'Thu, 25 Sep 2026 10:00:00 -0300' }) {
  const head = [`From: ${from}`, `To: ${to}`, ...(cc ? [`Cc: ${cc}`] : []), `Subject: ${subject}`, `Date: ${date}`, `Message-ID: <${crypto.randomUUID()}@contoso.com>`, 'MIME-Version: 1.0'];
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
  ];
  return {
    tenant: GRAPH_TENANT,
    clientId: GRAPH_CLIENT,
    secret: GRAPH_SECRET,
    users: [
      {
        id: 'u-ana',
        mail: 'ana@contoso.com',
        displayName: 'Ana Souza',
        proxyAddresses: ['SMTP:ana@contoso.com', 'smtp:ana.souza@contoso.com', 'smtp:a.souza@contoso.com'],
        accountEnabled: true,
        userType: 'Member',
        createdDateTime: '2020-03-15T10:00:00Z',
        department: 'RH',
        jobTitle: 'Analista',
        officeLocation: 'Sede',
        mobilePhone: '+55 21 99999-0000',
        assignedLicenses: [{ skuId: 'sku-1' }],
        folders,
        messages: {
          inbox: [
            { id: 'm1', received: '2026-09-20T10:00:00Z', raw: mail({ subject: 'Contrato', to: 'rh@contoso.com, bia@contoso.com', cc: 'chefe@contoso.com', body: 'x', attachments: [{ name: 'c.docx', data: Buffer.from('conteudo') }] }) },
            { id: 'm2', received: '2026-09-21T10:00:00Z', raw: mail({ subject: 'Almoço', body: 'y' }) },
          ],
          proj: [{ id: 'm3', received: '2026-09-22T10:00:00Z', raw: mail({ subject: 'Projeto', body: 'z' }) }],
        },
      },
      {
        id: 'u-bia',
        mail: 'bia@contoso.com',
        displayName: 'Bia Lima',
        upnIsMail: false,
        upn: 'bia.lima@contoso.com',
        proxyAddresses: ['SMTP:bia@contoso.com'],
        accountEnabled: false,
        userType: 'Member',
        createdDateTime: '2019-01-10T10:00:00Z',
        assignedLicenses: [],
        folders: [folders[0]],
        messages: { inbox: [{ id: 'b1', received: '2026-09-23T10:00:00Z', raw: mail({ subject: 'Oi', from: 'Ana <ana@contoso.com>', to: 'bia@contoso.com', body: 'oi' }) }] },
      },
      {
        id: 'u-ext',
        mail: 'parceiro@fornecedor.com',
        displayName: 'Parceiro Externo',
        userType: 'Guest',
        accountEnabled: true,
        createdDateTime: '2025-06-01T10:00:00Z',
        assignedLicenses: [],
        folders: [],
        messages: {},
      },
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
  excludeFolders: [],
  graph: { tenantId: GRAPH_TENANT, clientId: GRAPH_CLIENT },
  secrets: { clientSecret: GRAPH_SECRET },
  ...extra,
});

let mocks;
let google;
let root;

before(async () => {
  root = fs.mkdtempSync(path.join(os.tmpdir(), 'clean-listing-'));
  const { privateKey, publicKey } = crypto.generateKeyPairSync('rsa', { modulusLength: 2048 });
  google = {
    privateKeyPem: privateKey.export({ type: 'pkcs8', format: 'pem' }),
    data: {
      publicKey,
      admin: 'admin@empresa.com',
      users: [
        {
          mail: 'caio@empresa.com',
          name: 'Caio Dias',
          aliases: ['caio.dias@empresa.com'],
          suspended: false,
          isAdmin: true,
          isMailboxSetup: true,
          creationTime: '2021-02-01T00:00:00.000Z',
          lastLoginTime: '2026-09-01T08:00:00.000Z',
          orgUnitPath: '/Diretoria',
          organizations: [{ primary: true, department: 'Diretoria', title: 'CTO' }],
          labels: [{ id: 'INBOX', name: 'INBOX', type: 'system' }],
          messages: [{ id: 'g1', labelIds: ['INBOX'], internalDate: Date.parse('2026-09-01'), raw: mail({ subject: 'Planilha', body: 'x' }) }],
        },
        {
          mail: 'dora@empresa.com',
          name: 'Dora Reis',
          aliases: [],
          suspended: true,
          isMailboxSetup: true,
          creationTime: '2022-05-01T00:00:00.000Z',
          lastLoginTime: '1970-01-01T00:00:00.000Z',
          orgUnitPath: '/',
          labels: [],
          messages: [],
        },
      ],
    },
  };
  mocks = await startMockApis({ graph: graphData(), google: google.data });
});

after(async () => {
  await mocks?.close();
  fs.rmSync(root, { recursive: true, force: true });
});

async function runListing(sources, listing, options = {}, endpoints = mocks.endpoints) {
  const messages = [];
  const scanner = new MailScanner({ sources, terms: [], options, endpoints, listing }, (m) => messages.push(m));
  const stats = await scanner.run();
  const records = messages.filter((m) => m.type === 'results').flatMap((m) => m.records);
  const errors = messages.filter((m) => m.type === 'errors').flatMap((m) => m.items);
  const logs = messages.filter((m) => m.type === 'log');
  return { stats, records, errors, logs };
}

const gmailSource = (extra = {}) => ({
  id: 'src-google',
  name: 'Google',
  type: 'gmail',
  scope: 'all',
  mailboxes: [],
  excludeMailboxes: [],
  excludeFolders: [],
  gmail: { clientEmail: 'clean@projeto.iam.gserviceaccount.com', adminEmail: 'admin@empresa.com' },
  secrets: { privateKey: google.privateKeyPem },
  ...extra,
});

test('Listagem de contas: Microsoft 365 (apelidos, situação, licença, tipo e cadastro)', async () => {
  const { records, errors, stats } = await runListing([graphSource()], { kind: 'directory' });
  assert.deepEqual(errors, []);
  assert.equal(stats.accounts, 3);
  assert.ok(records.every((r) => r.kind === 'mail-account' && r.sourceType === 'graph'));
  const byAddress = Object.fromEntries(records.map((r) => [r.address, r]));
  const ana = byAddress['ana@contoso.com'];
  assert.equal(ana.name, 'Ana Souza');
  assert.deepEqual(ana.aliases.sort(), ['a.souza@contoso.com', 'ana.souza@contoso.com'], 'apelidos (smtp:) sem o principal (SMTP:)');
  assert.equal(ana.enabled, true);
  assert.equal(ana.licensed, true);
  assert.equal(ana.type, 'Membro', 'userType traduzido para pt-BR');
  assert.equal(ana.department, 'RH');
  assert.equal(ana.title, 'Analista');
  assert.ok(ana.created.startsWith('2020-03-15'));
  const bia = byAddress['bia@contoso.com'];
  assert.equal(bia.enabled, false);
  assert.equal(bia.licensed, false);
  assert.equal(bia.login, 'bia.lima@contoso.com', 'UPN diferente do endereço');
  assert.equal(byAddress['parceiro@fornecedor.com'].type, 'Convidado');
});

test('Listagem de contas: Google Workspace (Admin SDK, tipo, unidade e último acesso)', async () => {
  const { records, errors, stats } = await runListing([gmailSource()], { kind: 'directory' });
  assert.deepEqual(errors, []);
  assert.equal(stats.accounts, 2);
  const byAddress = Object.fromEntries(records.map((r) => [r.address, r]));
  const caio = byAddress['caio@empresa.com'];
  assert.equal(caio.name, 'Caio Dias');
  assert.equal(caio.type, 'Administrador');
  assert.equal(caio.admin, true);
  assert.equal(caio.enabled, true);
  assert.equal(caio.department, 'Diretoria');
  assert.equal(caio.title, 'CTO');
  assert.equal(caio.orgUnit, '/Diretoria');
  assert.ok(caio.lastActivity.startsWith('2026-09-01'));
  const dora = byAddress['dora@empresa.com'];
  assert.equal(dora.enabled, false, 'suspensa = inativa');
  assert.equal(dora.suspended, true);
  assert.equal(dora.lastActivity, null, 'nunca entrou (1970) fica sem data');
});

test('Listagem de contas: IMAP traz só as caixas cadastradas, com aviso', async () => {
  const source = {
    id: 'src-imap',
    name: 'Servidor interno',
    type: 'imap',
    scope: 'list',
    imap: { host: '127.0.0.1', port: 1, security: 'tls' },
    mailboxes: [{ address: 'a@empresa.com', login: 'a@empresa.com' }, { address: 'b@empresa.com' }],
    excludeMailboxes: [],
    excludeFolders: [],
    secrets: { defaultPassword: 'p' },
  };
  const { records, errors, stats } = await runListing([source], { kind: 'directory' });
  assert.deepEqual(errors, []);
  assert.equal(stats.accounts, 2);
  assert.deepEqual(records.map((r) => r.address).sort(), ['a@empresa.com', 'b@empresa.com']);
  assert.ok(records.every((r) => r.type === 'Caixa cadastrada (IMAP)' && /não tem um catálogo/.test(r.note)));
});

test('Listagem de mensagens: Microsoft 365 (cabeçalhos, destinatários, anexos e sem baixar conteúdo)', async () => {
  const before = mocks.calls.length;
  const { records, errors, stats } = await runListing([graphSource({ scope: 'list', mailboxes: [{ address: 'ana@contoso.com' }] })], { kind: 'messages' });
  assert.deepEqual(errors, []);
  assert.equal(stats.messagesSeen, 3);
  assert.ok(records.every((r) => r.kind === 'mail-message' && r.sourceType === 'graph'));
  assert.deepEqual(records.map((r) => r.subject).sort(), ['Almoço', 'Contrato', 'Projeto']);
  const contrato = records.find((r) => r.subject === 'Contrato');
  assert.equal(contrato.mailbox, 'ana@contoso.com');
  assert.equal(contrato.folder, 'Caixa de Entrada');
  assert.ok(contrato.to.some((t) => t.includes('bia@contoso.com')) && contrato.to.some((t) => t.includes('rh@contoso.com')));
  assert.ok(contrato.cc.some((c) => c.includes('chefe@contoso.com')));
  assert.equal(contrato.hasAttachments, true);
  assert.ok(contrato.size > 0 && contrato.date && contrato.from);
  assert.equal(records.find((r) => r.subject === 'Almoço').hasAttachments, false);
  const calls = mocks.calls.slice(before);
  assert.ok(!calls.some((c) => c.includes('/$value')), 'a listagem não baixa o conteúdo das mensagens');
});

test('Listagem de mensagens: IMAP (destinatários do envelope, só os cabeçalhos)', async () => {
  const imap = await startFakeImap({
    'carla@empresa.com': {
      password: 'p',
      folders: {
        INBOX: [
          { raw: mail({ subject: 'Reunião', from: 'Chefe <chefe@empresa.com>', to: 'carla@empresa.com, equipe@empresa.com', cc: 'rh@empresa.com', body: 'vamos' }), date: new Date('2026-09-10T12:00:00Z') },
          { raw: mail({ subject: 'Com anexo', from: 'RH <rh@empresa.com>', to: 'carla@empresa.com', body: 'segue', attachments: [{ name: 'doc.pdf', data: Buffer.from('conteudo') }] }), date: new Date('2026-09-11T12:00:00Z') },
        ],
      },
    },
  });
  try {
    const source = {
      id: 'src-imap',
      name: 'IMAP',
      type: 'imap',
      scope: 'list',
      imap: { host: '127.0.0.1', port: imap.port, security: 'none' },
      mailboxes: [{ address: 'carla@empresa.com' }],
      excludeMailboxes: [],
      excludeFolders: [],
      secrets: { defaultPassword: 'p' },
    };
    const { records, errors, stats } = await runListing([source], { kind: 'messages' });
    assert.deepEqual(errors, []);
    assert.equal(stats.messagesSeen, 2);
    const byS = Object.fromEntries(records.map((r) => [r.subject, r]));
    const m = byS['Reunião'];
    assert.equal(m.folder, 'INBOX');
    assert.equal(m.from, 'Chefe <chefe@empresa.com>');
    assert.ok(m.to.some((t) => t.includes('carla@empresa.com')) && m.to.some((t) => t.includes('equipe@empresa.com')));
    assert.ok(m.cc.some((c) => c.includes('rh@empresa.com')));
    assert.ok(m.sent, 'a data de envio vem do envelope');
    // hasAttachments vem da estrutura do corpo (BODYSTRUCTURE).
    assert.equal(m.hasAttachments, false);
    assert.equal(byS['Com anexo'].hasAttachments, true);
  } finally {
    await imap.close();
  }
});

test('Listagem de mensagens: Google Workspace (metadados, destinatários e anexos)', async () => {
  const mock = await startMockApis({
    google: {
      publicKey: google.data.publicKey,
      admin: 'admin@empresa.com',
      users: [
        {
          mail: 'caio@empresa.com',
          name: 'Caio',
          labels: [{ id: 'INBOX', name: 'INBOX', type: 'system' }],
          messages: [
            { id: 'g1', labelIds: ['INBOX'], internalDate: Date.parse('2026-09-01'), raw: mail({ subject: 'Sem anexo', to: 'rh@empresa.com', body: 'oi' }) },
            {
              id: 'g2',
              labelIds: ['INBOX'],
              internalDate: Date.parse('2026-09-02'),
              raw: mail({ subject: 'Com anexo', to: 'rh@empresa.com', cc: 'chefe@empresa.com', body: 'veja', attachments: [{ name: 'nota.pdf', data: Buffer.from('x') }] }),
              payload: { parts: [{ mimeType: 'text/plain' }, { mimeType: 'application/pdf', filename: 'nota.pdf' }] },
            },
          ],
        },
      ],
    },
  });
  try {
    const source = {
      id: 'g',
      name: 'Google',
      type: 'gmail',
      scope: 'list',
      mailboxes: [{ address: 'caio@empresa.com' }],
      excludeMailboxes: [],
      excludeFolders: [],
      gmail: { clientEmail: 'clean@projeto.iam.gserviceaccount.com' },
      secrets: { privateKey: google.privateKeyPem },
    };
    const { records, errors, stats } = await runListing([source], { kind: 'messages' }, {}, mock.endpoints);
    assert.deepEqual(errors, []);
    assert.equal(stats.messagesSeen, 2);
    const byS = Object.fromEntries(records.map((r) => [r.subject, r]));
    assert.ok(byS['Com anexo'].to.some((t) => t.includes('rh@empresa.com')));
    assert.ok(byS['Com anexo'].cc.some((c) => c.includes('chefe@empresa.com')));
    assert.ok(byS['Com anexo'].sent, 'a data de envio vem dos cabeçalhos');
    assert.equal(byS['Com anexo'].hasAttachments, true);
    assert.equal(byS['Sem anexo'].hasAttachments, false);
  } finally {
    await mock.close();
  }
});

// ---------------------------------------------------------------------------------------------
// API: listagem em segundo plano, relatório somente leitura e exportações

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
    return { status: res.status, data: type.includes('json') ? await res.json() : Buffer.from(await res.arrayBuffer()) };
  };
  const wait = async (id) => {
    let current;
    for (let i = 0; i < 300; i++) {
      current = await api('GET', `/api/scans/${id}`);
      if (!['queued', 'running'].includes(current.data.status)) break;
      await new Promise((r) => setTimeout(r, 50));
    }
    return current;
  };
  return { store, manager, srv, api, wait, close: async () => (srv.close(), await manager.shutdown()) };
}

test('API: catálogo de contas do domínio (relatório somente leitura e exportações)', async () => {
  const app = await startApp();
  try {
    const source = await app.api('POST', '/api/mail-sources', {
      name: 'Microsoft 365',
      type: 'graph',
      scope: 'all',
      graph: { tenantId: GRAPH_TENANT, clientId: GRAPH_CLIENT, clientSecret: GRAPH_SECRET },
    });
    assert.equal(source.status, 201);
    const scan = await app.api('POST', '/api/scans', { listing: { kind: 'directory' }, name: 'Contas', sourceIds: [source.data.id] });
    assert.equal(scan.status, 201);
    assert.deepEqual(scan.data.listing, { kind: 'directory' });
    const done = await app.wait(scan.data.id);
    assert.equal(done.data.status, 'completed', JSON.stringify(done.data.log));
    assert.equal(done.data.stats.accounts, 3);

    const config = fs.readFileSync(path.join(app.store.scanDir(scan.data.id), 'config.json'), 'utf8');
    assert.ok(!config.includes(GRAPH_SECRET) && !config.includes('enc:v1'), 'a configuração não guarda segredos');

    const results = await app.api('GET', `/api/scans/${scan.data.id}/results`);
    assert.equal(results.data.total, 3);
    assert.ok(results.data.items.every((r) => r.canDelete === false && r.deleteMethod === null), 'listagem é somente leitura');
    assert.equal(results.data.items[0].address, 'ana@contoso.com', 'ordenado por endereço');
    const inactive = await app.api('GET', `/api/scans/${scan.data.id}/results?state=inactive`);
    assert.deepEqual(inactive.data.items.map((r) => r.address), ['bia@contoso.com']);

    const summary = await app.api('GET', `/api/scans/${scan.data.id}/summary`);
    assert.equal(summary.data.accounts, 3);
    assert.equal(summary.data.disabled, 1);
    assert.ok(summary.data.byState.some((g) => g.key === 'active'));

    // Somente leitura: a exclusão item a item é recusada.
    const del = await app.api('POST', `/api/scans/${scan.data.id}/results/1/delete`, { confirm: true });
    assert.equal(del.status, 400);

    // A listagem não aparece na lista de análises, só na de listagens.
    const analyses = await app.api('GET', '/api/scans?kind=mail');
    assert.ok(!analyses.data.some((s) => s.id === scan.data.id));
    const listings = await app.api('GET', '/api/scans?kind=mail&listing=only');
    assert.deepEqual(listings.data.map((s) => s.id), [scan.data.id]);

    const xlsx = await app.api('GET', `/api/scans/${scan.data.id}/export.xlsx`);
    const parts = unzipSync(new Uint8Array(xlsx.data));
    const workbook = strFromU8(parts['xl/workbook.xml']);
    assert.deepEqual([...workbook.matchAll(/<sheet name="([^"]*)"/g)].map((m) => m[1]), ['Resumo', 'Contas']);
    const csv = await app.api('GET', `/api/scans/${scan.data.id}/export.csv`);
    assert.ok(csv.data.toString('utf8').startsWith('﻿Endereço principal;Nome;Login (UPN)'));
    const html = await app.api('GET', `/api/scans/${scan.data.id}/export.html`);
    assert.ok(html.data.toString('utf8').includes('Contas do domínio'));
  } finally {
    await app.close();
  }
});

test('API: listagem de mensagens por caixa (metadados, filtros e exportações)', async () => {
  const app = await startApp();
  try {
    const source = await app.api('POST', '/api/mail-sources', {
      name: 'Microsoft 365',
      type: 'graph',
      scope: 'list',
      mailboxes: 'ana@contoso.com',
      graph: { tenantId: GRAPH_TENANT, clientId: GRAPH_CLIENT, clientSecret: GRAPH_SECRET },
    });
    const scan = await app.api('POST', '/api/scans', { listing: { kind: 'messages' }, name: 'Mensagens', sourceIds: [source.data.id], options: { includeTrash: true } });
    assert.equal(scan.status, 201);
    const done = await app.wait(scan.data.id);
    assert.equal(done.data.status, 'completed', JSON.stringify(done.data.log));
    assert.equal(done.data.stats.messagesSeen, 3);

    const results = await app.api('GET', `/api/scans/${scan.data.id}/results`);
    assert.equal(results.data.total, 3);
    assert.ok(results.data.items.every((r) => r.canDelete === false));
    const withAttach = await app.api('GET', `/api/scans/${scan.data.id}/results?attachments=yes`);
    assert.deepEqual(withAttach.data.items.map((r) => r.subject), ['Contrato']);
    const inProjects = await app.api('GET', `/api/scans/${scan.data.id}/results?folder=${encodeURIComponent('Caixa de Entrada/Projetos')}`);
    assert.deepEqual(inProjects.data.items.map((r) => r.subject), ['Projeto']);

    const summary = await app.api('GET', `/api/scans/${scan.data.id}/summary`);
    assert.equal(summary.data.messages, 3);
    assert.equal(summary.data.withAttachments, 1);
    assert.ok(summary.data.byFolder.length >= 2);

    const xlsx = await app.api('GET', `/api/scans/${scan.data.id}/export.xlsx`);
    const parts = unzipSync(new Uint8Array(xlsx.data));
    const workbook = strFromU8(parts['xl/workbook.xml']);
    assert.deepEqual([...workbook.matchAll(/<sheet name="([^"]*)"/g)].map((m) => m[1]), ['Resumo', 'Mensagens']);
    const json = await app.api('GET', `/api/scans/${scan.data.id}/export.json`);
    assert.equal(json.data.results.length, 3);
    assert.equal(json.data.scan.listing.kind, 'messages');
  } finally {
    await app.close();
  }
});

test('API: listagem com tipo inválido ou sem conexões é recusada', async () => {
  const app = await startApp();
  try {
    const source = await app.api('POST', '/api/mail-sources', { name: 'M365', type: 'graph', scope: 'all', graph: { tenantId: GRAPH_TENANT, clientId: GRAPH_CLIENT, clientSecret: GRAPH_SECRET } });
    const badKind = await app.api('POST', '/api/scans', { listing: { kind: 'xpto' }, sourceIds: [source.data.id] });
    assert.equal(badKind.status, 400);
    const noSources = await app.api('POST', '/api/scans', { listing: { kind: 'directory' }, sourceIds: [] });
    assert.equal(noSources.status, 400);
  } finally {
    await app.close();
  }
});
