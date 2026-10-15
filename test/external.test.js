// Destinatários externos ao domínio (somente leitura): separa os destinatários (Para/Cc/Cco) com
// domínio fora da empresa. Testa o motor, o escopo de pastas (Enviados), o Cco, a API, o relatório e
// as exportações com servidores simulados (Microsoft Graph).
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

const TENANT = 'empresa.onmicrosoft.com';
const CLIENT = '11111111-2222-3333-4444-555555555555';
const SECRET = 'segredo-super-secreto';

/** Monta uma mensagem MIME simples (aceita Cco e anexos). */
function mail({ subject, from, to = '', cc = '', bcc = '', body = 'x', attachments = [], date = 'Thu, 25 Sep 2026 10:00:00 -0300' }) {
  const head = [
    `From: ${from}`,
    `To: ${to}`,
    ...(cc ? [`Cc: ${cc}`] : []),
    ...(bcc ? [`Bcc: ${bcc}`] : []),
    `Subject: ${subject}`,
    `Date: ${date}`,
    `Message-ID: <${crypto.randomUUID()}@empresa.com>`,
    'MIME-Version: 1.0',
  ];
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
  const inbox = { id: 'inbox', displayName: 'Caixa de Entrada', wellKnown: 'inbox' };
  const sent = { id: 'sent', displayName: 'Itens Enviados', wellKnown: 'sentitems' };
  return {
    tenant: TENANT,
    clientId: CLIENT,
    secret: SECRET,
    users: [
      {
        id: 'u-ana',
        mail: 'ana@empresa.com',
        displayName: 'Ana Souza',
        accountEnabled: true,
        userType: 'Member',
        folders: [inbox, sent],
        messages: {
          inbox: [
            // Recebida de dentro, com um externo em cópia.
            { id: 'a1', received: '2026-09-20T10:00:00Z', raw: mail({ subject: 'Proposta', from: 'chefe@empresa.com', to: 'ana@empresa.com', cc: 'parceiro@fornecedor.com' }) },
            // Só internos: fica de fora.
            { id: 'a2', received: '2026-09-21T10:00:00Z', raw: mail({ subject: 'Interno', from: 'ana@empresa.com', to: 'colega@empresa.com' }) },
            // Remetente externo, sem destinatário externo: fica de fora (mas conta como remetente externo).
            { id: 'a3', received: '2026-09-22T10:00:00Z', raw: mail({ subject: 'Newsletter', from: 'news@parceiro.com', to: 'ana@empresa.com' }) },
          ],
          sent: [
            // Enviada para fora, com Cc interno, Cco externo e um anexo.
            {
              id: 's1',
              received: '2026-09-23T10:00:00Z',
              raw: mail({ subject: 'Para o cliente', from: 'ana@empresa.com', to: 'cliente@cliente.com', cc: 'chefe@empresa.com', bcc: 'oculto@secreto.com', body: 'corpo-visivel', attachments: [{ name: 'contrato.pdf', data: Buffer.from('CONTEUDO-DO-ANEXO') }] }),
            },
          ],
        },
      },
      {
        id: 'u-bia',
        mail: 'bia@empresa.com',
        displayName: 'Bia Lima',
        accountEnabled: true,
        userType: 'Member',
        folders: [{ ...inbox }],
        messages: {
          inbox: [{ id: 'b1', received: '2026-09-24T10:00:00Z', raw: mail({ subject: 'Outro cliente', from: 'bia@empresa.com', to: 'cliente@cliente.com' }) }],
        },
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
  graph: { tenantId: TENANT, clientId: CLIENT },
  secrets: { clientSecret: SECRET },
  ...extra,
});

let mocks;
let root;

before(async () => {
  root = fs.mkdtempSync(path.join(os.tmpdir(), 'clean-external-'));
  mocks = await startMockApis({ graph: graphData() });
});

after(async () => {
  await mocks?.close();
  fs.rmSync(root, { recursive: true, force: true });
});

async function runExternal(sources, listing, options = {}, endpoints = mocks.endpoints) {
  const messages = [];
  const scanner = new MailScanner({ sources, terms: [], options, endpoints, listing: { kind: 'external', ownDomains: ['empresa.com'], folderScope: 'all', ...listing } }, (m) => messages.push(m));
  const stats = await scanner.run();
  const records = messages.filter((m) => m.type === 'results').flatMap((m) => m.records);
  const errors = messages.filter((m) => m.type === 'errors').flatMap((m) => m.items);
  return { stats, records, errors };
}

test('Motor: separa destinatários externos (Para/Cc/Cco), ignora só-internos e conta remetente externo', async () => {
  const { records, errors, stats } = await runExternal([graphSource()], {});
  assert.deepEqual(errors, []);
  assert.equal(stats.messagesSeen, 5, 'todas as mensagens são verificadas');
  assert.equal(stats.externalMessages, 3, 'a1 (cc externo), s1 (to/bcc externos) e b1 (to externo)');
  assert.equal(stats.externalSenders, 1, 'a3 veio de remetente externo');
  assert.ok(records.every((r) => r.kind === 'mail-external' && r.sourceType === 'graph'));
  assert.deepEqual(records.map((r) => r.subject).sort(), ['Outro cliente', 'Para o cliente', 'Proposta']);

  const proposta = records.find((r) => r.subject === 'Proposta');
  assert.equal(proposta.fromExternal, false, 'chefe@empresa.com é interno');
  assert.deepEqual(
    proposta.external.map((e) => e.address),
    ['parceiro@fornecedor.com'],
  );
  assert.equal(proposta.external[0].domain, 'fornecedor.com');
  assert.deepEqual(proposta.external[0].fields, ['cc']);

  const cliente = records.find((r) => r.subject === 'Para o cliente');
  const byAddr = Object.fromEntries(cliente.external.map((e) => [e.address, e]));
  assert.deepEqual(Object.keys(byAddr).sort(), ['cliente@cliente.com', 'oculto@secreto.com'], 'o Cco externo entra; o Cc interno não');
  assert.deepEqual(byAddr['cliente@cliente.com'].fields, ['to']);
  assert.deepEqual(byAddr['oculto@secreto.com'].fields, ['bcc'], 'Cco lido do cabeçalho Bcc');
});

test('Motor: escopo "Apenas Enviados" varre só a pasta de enviados', async () => {
  const { records, stats } = await runExternal([graphSource({ scope: 'list', mailboxes: [{ address: 'ana@empresa.com' }] })], { folderScope: 'sent' });
  assert.equal(stats.messagesSeen, 1, 'só a mensagem da pasta Enviados');
  assert.deepEqual(records.map((r) => r.subject), ['Para o cliente']);
});

test('Motor: sem domínios internos informados, usa o domínio da própria caixa', async () => {
  const messages = [];
  const scanner = new MailScanner(
    { sources: [graphSource({ scope: 'list', mailboxes: [{ address: 'ana@empresa.com' }] })], terms: [], options: {}, endpoints: mocks.endpoints, listing: { kind: 'external', ownDomains: [], folderScope: 'all' } },
    (m) => messages.push(m),
  );
  await scanner.run();
  const records = messages.filter((m) => m.type === 'results').flatMap((m) => m.records);
  // empresa.com é o domínio da caixa: só fornecedor.com, cliente.com e secreto.com são externos.
  const domains = new Set(records.flatMap((r) => r.external.map((e) => e.domain)));
  assert.deepEqual([...domains].sort(), ['cliente.com', 'fornecedor.com', 'secreto.com']);
  assert.ok(!records.some((r) => r.external.some((e) => e.domain === 'empresa.com')));
});

// ---------------------------------------------------------------------------------------------
// API: varredura em segundo plano, relatório somente leitura e exportações

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

test('API: destinatários externos (relatório somente leitura, apurações e exportações)', async () => {
  const app = await startApp();
  try {
    const source = await app.api('POST', '/api/mail-sources', {
      name: 'Microsoft 365',
      type: 'graph',
      scope: 'all',
      graph: { tenantId: TENANT, clientId: CLIENT, clientSecret: SECRET },
    });
    assert.equal(source.status, 201);
    const scan = await app.api('POST', '/api/scans', {
      listing: { kind: 'external', ownDomains: 'empresa.com', folderScope: 'all' },
      name: 'Externos',
      sourceIds: [source.data.id],
    });
    assert.equal(scan.status, 201);
    assert.equal(scan.data.listing.kind, 'external');
    assert.deepEqual(scan.data.listing.ownDomains, ['empresa.com']);
    const done = await app.wait(scan.data.id);
    assert.equal(done.data.status, 'completed', JSON.stringify(done.data.log));
    assert.equal(done.data.stats.externalMessages, 3);

    // A configuração não guarda segredos.
    const config = fs.readFileSync(path.join(app.store.scanDir(scan.data.id), 'config.json'), 'utf8');
    assert.ok(!config.includes(SECRET) && !config.includes('enc:v1'), 'a configuração não guarda segredos');

    // Somente leitura.
    const results = await app.api('GET', `/api/scans/${scan.data.id}/results`);
    assert.equal(results.data.total, 3);
    assert.ok(results.data.items.every((r) => r.canDelete === false && r.deleteMethod === null));
    const del = await app.api('POST', `/api/scans/${scan.data.id}/results/1/delete`, { confirm: true });
    assert.equal(del.status, 400);

    // Resumo: por domínio, por destinatário e distintos.
    const summary = await app.api('GET', `/api/scans/${scan.data.id}/summary`);
    assert.equal(summary.data.messages, 3);
    assert.equal(summary.data.recipientsTotal, 3, 'parceiro@, cliente@ e oculto@');
    assert.equal(summary.data.domainsTotal, 3, 'fornecedor.com, cliente.com, secreto.com');
    const cliente = summary.data.byRecipient.find((r) => r.address === 'cliente@cliente.com');
    assert.equal(cliente.messages, 2, 'cliente@ aparece em duas mensagens (duas caixas)');
    assert.equal(cliente.mailboxes, 2);
    const fornecedor = summary.data.byDomain.find((d) => d.domain === 'fornecedor.com');
    assert.equal(fornecedor.messages, 1);

    // Filtro por domínio externo (via chave de filtro do gráfico).
    const porDominio = await app.api('GET', `/api/scans/${scan.data.id}/results?domain=cliente.com`);
    assert.deepEqual(porDominio.data.items.map((r) => r.subject).sort(), ['Outro cliente', 'Para o cliente']);

    // Segregação: só aparece na lista de listagens.
    const analyses = await app.api('GET', '/api/scans?kind=mail');
    assert.ok(!analyses.data.some((s) => s.id === scan.data.id));
    const listings = await app.api('GET', '/api/scans?kind=mail&listing=only');
    assert.ok(listings.data.some((s) => s.id === scan.data.id));

    // Exportações.
    const xlsx = await app.api('GET', `/api/scans/${scan.data.id}/export.xlsx`);
    const parts = unzipSync(new Uint8Array(xlsx.data));
    const workbook = strFromU8(parts['xl/workbook.xml']);
    assert.deepEqual([...workbook.matchAll(/<sheet name="([^"]*)"/g)].map((m) => m[1]), ['Resumo', 'Destinatários', 'Domínios', 'Mensagens']);
    const csv = await app.api('GET', `/api/scans/${scan.data.id}/export.csv`);
    assert.ok(csv.data.toString('utf8').startsWith('﻿Destinatário externo;Nome;Domínio'));
    const html = await app.api('GET', `/api/scans/${scan.data.id}/export.html`);
    assert.ok(html.data.toString('utf8').includes('Destinatários externos ao domínio'));
    const json = await app.api('GET', `/api/scans/${scan.data.id}/export.json`);
    assert.equal(json.data.scan.listing.kind, 'external');
    assert.equal(json.data.results.length, 3);

    // Abrir a mensagem sob demanda: conteúdo, anexo e .eml.
    const all = await app.api('GET', `/api/scans/${scan.data.id}/results?pageSize=500`);
    const alvo = all.data.items.find((r) => r.subject === 'Para o cliente');
    assert.ok(alvo, 'a mensagem com anexo está no relatório');
    const msg = await app.api('GET', `/api/scans/${scan.data.id}/results/${alvo.id}/message`);
    assert.equal(msg.status, 200);
    assert.ok(msg.data.to.some((t) => t.includes('cliente@cliente.com')));
    assert.ok(msg.data.bcc.some((t) => t.includes('oculto@secreto.com')), 'o Cco aparece no conteúdo');
    assert.ok(msg.data.body.includes('corpo-visivel'));
    assert.equal(msg.data.attachments.length, 1);
    assert.equal(msg.data.attachments[0].name, 'contrato.pdf');
    const att = await app.api('GET', `/api/scans/${scan.data.id}/results/${alvo.id}/attachment/0`);
    assert.equal(att.status, 200);
    assert.equal(att.data.toString('utf8'), 'CONTEUDO-DO-ANEXO');
    const eml = await app.api('GET', `/api/scans/${scan.data.id}/results/${alvo.id}/message.eml`);
    assert.equal(eml.status, 200);
    assert.ok(eml.data.toString('utf8').includes('Subject: Para o cliente'));

    // Uma linha sem mensagem para abrir é recusada (contas do domínio não têm messageId) — aqui, id inexistente.
    const semMsg = await app.api('GET', `/api/scans/${scan.data.id}/results/999999/message`);
    assert.equal(semMsg.status, 404);
  } finally {
    await app.close();
  }
});
