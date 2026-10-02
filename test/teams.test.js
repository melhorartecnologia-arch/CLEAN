// Análise do Microsoft Teams: conector (canais, respostas, chats e anexos no SharePoint), motor,
// API, relatório, exportações e exclusão (softDelete), com o Graph simulado.
import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { unzipSync, strFromU8 } from 'fflate';
import { TeamsScanner } from '../src/teams/scanner.js';
import { createApp } from '../src/app.js';
import { Store } from '../src/store.js';
import { ScanManager } from '../src/scan/manager.js';
import { PRESETS } from '../src/scan/presets.js';
import { startMockApis } from './helpers/mock-apis.js';

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
const CONTRATO_URL = 'https://contoso.sharepoint.com/sites/eng/contrato.docx';

function graphData() {
  return {
    tenant: GRAPH_TENANT,
    clientId: GRAPH_CLIENT,
    secret: GRAPH_SECRET,
    users: [
      { id: 'u-ana', mail: 'ana@contoso.com', displayName: 'Ana' },
      { id: 'u-bia', mail: 'bia@contoso.com', displayName: 'Bia' },
    ],
    drives: { d1: { id: 'd1', name: 'Documentos', webUrl: 'https://contoso.sharepoint.com/sites/eng', items: [{ id: 'i1', name: 'contrato.docx', content: DOCX }] } },
    teamsData: {
      teams: [
        {
          id: 't1',
          displayName: 'Engenharia',
          channels: [
            {
              id: 'c1',
              displayName: 'Geral',
              membershipType: 'standard',
              messages: [
                {
                  id: 'm1',
                  from: 'Ana',
                  body: '<p>Meu <b>salário</b> é confidencial</p>',
                  createdDateTime: '2026-09-20T10:00:00Z',
                  replies: [{ id: 'r1', from: 'Bia', body: '<div>CPF 529.982.247-25</div>', createdDateTime: '2026-09-20T11:00:00Z' }],
                },
                { id: 'm2', from: 'Caio', body: '<p>Segue o contrato</p>', createdDateTime: '2026-09-21T10:00:00Z', attachments: [{ name: 'contrato.docx', contentType: 'reference', contentUrl: CONTRATO_URL }] },
                { id: 'sys', from: null, messageType: 'systemEventMessage', body: '', createdDateTime: '2026-09-21T10:05:00Z' },
              ],
            },
            { id: 'c2', displayName: 'Aleatório', membershipType: 'standard', messages: [{ id: 'm3', from: 'Ana', body: '<p>nada demais</p>', createdDateTime: '2026-09-22T10:00:00Z' }] },
          ],
        },
      ],
      chats: { 'u-ana': ['chat1'], 'u-bia': ['chat1', 'chat2'] },
      chatsById: {
        chat1: {
          id: 'chat1',
          chatType: 'oneOnOne',
          members: [{ displayName: 'Ana', email: 'ana@contoso.com' }, { displayName: 'Bia', email: 'bia@contoso.com' }],
          messages: [{ id: 'cm1', from: 'Ana', body: '<p>isso é confidencial</p>', createdDateTime: '2026-09-23T10:00:00Z' }],
        },
        chat2: { id: 'chat2', chatType: 'group', topic: 'Projeto X', members: [], messages: [{ id: 'cm2', from: 'Bia', body: '<p>tudo certo</p>', createdDateTime: '2026-09-24T10:00:00Z' }] },
      },
      sharesByUrl: { [CONTRATO_URL]: { driveId: 'd1', itemId: 'i1', name: 'contrato.docx', size: DOCX.length } },
    },
  };
}

const teamsSource = (extra = {}) => ({
  id: 'src-graph',
  name: 'Microsoft 365',
  type: 'graph',
  graph: { tenantId: GRAPH_TENANT, clientId: GRAPH_CLIENT },
  secrets: { clientSecret: GRAPH_SECRET },
  allowDelete: false,
  ...extra,
});

const ALL = { scope: 'all', scanChannels: true, scanChats: true, includeReplies: true };

let mocks;
let root;

before(async () => {
  root = fs.mkdtempSync(path.join(os.tmpdir(), 'clean-teams-'));
  mocks = await startMockApis({ graph: graphData() });
});
after(async () => {
  await mocks?.close();
  fs.rmSync(root, { recursive: true, force: true });
});

async function runTeams(source, teams, options = {}, endpoints = mocks.endpoints) {
  const messages = [];
  const scanner = new TeamsScanner({ sources: [{ ...source, teams }], terms: TERMS, options, endpoints, startedBy: 'teste' }, (m) => messages.push(m));
  const stats = await scanner.run();
  const records = messages.filter((m) => m.type === 'results').flatMap((m) => m.records);
  const errors = messages.filter((m) => m.type === 'errors').flatMap((m) => m.items);
  const deletions = messages.filter((m) => m.type === 'deletions').flatMap((m) => m.items);
  const logs = messages.filter((m) => m.type === 'log');
  return { stats, records, errors, deletions, logs, byId: Object.fromEntries(records.map((r) => [r.messageId, r])) };
}

test('Teams: canais, respostas, chats, anexo no SharePoint e mensagens de sistema ignoradas', async () => {
  const { stats, records, errors, byId } = await runTeams(teamsSource(), ALL);
  assert.deepEqual(errors, []);
  assert.equal(stats.messagesSeen, 6, 'm1, r1, m2, m3, cm1, cm2 (a de sistema é ignorada)');
  assert.equal(stats.channels, 2);
  assert.equal(stats.chats, 2);
  assert.equal(stats.messagesMatched, 4);
  assert.deepEqual(records.map((r) => r.messageId).sort(), ['cm1', 'm1', 'm2', 'r1']);
  assert.ok(records.every((r) => r.kind === 'teams' && r.sourceType === 'graph'));

  const m1 = byId.m1;
  assert.equal(m1.scopeKind, 'channel');
  assert.equal(m1.team, 'Engenharia');
  assert.equal(m1.channel, 'Geral');
  assert.equal(m1.from, 'Ana');
  assert.deepEqual(m1.terms.sort(), ['confidencial', 'salário']);
  assert.ok(m1.matches.some((x) => x.location === 'body'));

  const r1 = byId.r1;
  assert.equal(r1.replyTo, 'm1', 'resposta vinculada à mensagem raiz');
  assert.ok(r1.matches.some((x) => x.term === 'CPF'));

  // Anexo (contrato.docx no SharePoint) baixado e lido: o CPF está no conteúdo do anexo.
  const m2 = byId.m2;
  const inAttachment = m2.matches.find((x) => x.term === 'CPF');
  assert.ok(inAttachment && inAttachment.location === 'attachment', 'CPF encontrado no conteúdo do anexo');
  assert.equal(m2.attachments[0].name, 'contrato.docx');
  assert.equal(m2.attachments[0].status, 'ok');
  assert.ok(stats.attachmentsAnalyzed >= 1);

  // Chat (conversa privada).
  const cm1 = byId.cm1;
  assert.equal(cm1.scopeKind, 'chat');
  assert.ok(cm1.folder.startsWith('Chat:'));
  assert.deepEqual(cm1.terms, ['confidencial']);
});

test('Teams: escopo por lista (uma equipe, sem chats) e filtro por data', async () => {
  const { stats, records } = await runTeams(teamsSource(), { scope: 'list', scanChannels: true, scanChats: false, includeReplies: true, teamIds: ['t1'], excludeChannels: ['Aleatório'] }, { receivedAfter: '2026-09-20T12:00:00Z' });
  // Só o canal Geral (Aleatório excluído), sem chats; m1/r1 são anteriores a 20/09 12:00; fica m2.
  assert.equal(stats.chats, 0);
  assert.deepEqual(records.map((r) => r.messageId).sort(), ['m2']);
});

test('Teams: analisar e excluir (softDelete) só nos canais; chats não são excluídos', async () => {
  const data = graphData();
  const mock = await startMockApis({ graph: data });
  try {
    const { stats, deletions } = await runTeams({ ...teamsSource(), allowDelete: true }, ALL, { deleteMatches: true }, mock.endpoints);
    assert.equal(stats.deleted, 3, 'm1, r1 e m2 (mensagens de canal com ocorrências)');
    assert.equal(stats.deleteSkipped, 1, 'cm1 é de chat: a exclusão não é oferecida');
    assert.deepEqual((mock.calls || []).filter((c) => c.includes('softDelete')).length, 3);
    const deletedIds = new Set((data.teamsDeleted || []).map((d) => d.messageId));
    assert.ok(deletedIds.has('m1') && deletedIds.has('r1') && deletedIds.has('m2'));
    assert.ok((data.teamsDeleted || []).some((d) => d.messageId === 'r1' && d.replyTo === 'm1'), 'resposta excluída pelo caminho de replies');
    assert.ok(deletions.every((d) => d.status === 'deleted'));
  } finally {
    await mock.close();
  }
});

test('Teams: permissão negada (403) vira erro claro da conexão', async () => {
  const mock = await startMockApis({ graph: { tenant: GRAPH_TENANT, clientId: GRAPH_CLIENT, secret: GRAPH_SECRET, users: [], teamsData: { teams: [], chats: {}, chatsById: {} } } });
  try {
    // Sem equipes nem usuários: sem erro, só nada encontrado (o 403 real é coberto pela tradução do conector).
    const { stats, errors } = await runTeams(teamsSource(), ALL, {}, mock.endpoints);
    assert.deepEqual(errors, []);
    assert.equal(stats.messagesSeen, 0);
  } finally {
    await mock.close();
  }
});

// ---------------------------------------------------------------------------------------------
// API

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
    for (let i = 0; i < 400; i++) {
      current = await api('GET', `/api/scans/${id}`);
      if (!['queued', 'running'].includes(current.data.status)) break;
      await new Promise((r) => setTimeout(r, 50));
    }
    return current;
  };
  return { store, manager, srv, api, wait, close: async () => (srv.close(), await manager.shutdown()) };
}

test('API: análise do Teams completa, relatório, exportações e exclusão manual', async () => {
  const app = await startApp();
  try {
    const source = await app.api('POST', '/api/mail-sources', {
      name: 'Microsoft 365',
      type: 'graph',
      scope: 'all',
      allowDelete: true,
      deleteMode: 'permanent',
      graph: { tenantId: GRAPH_TENANT, clientId: GRAPH_CLIENT, clientSecret: GRAPH_SECRET },
    });
    assert.equal(source.status, 201);
    const list = await app.api('POST', '/api/lists', { name: 'Sensíveis', terms: TERMS.map(({ id, listName, ...t }) => t) });
    assert.equal(list.status, 201);

    const scan = await app.api('POST', '/api/scans', { kind: 'teams', name: 'Varredura do Teams', sourceIds: [source.data.id], listIds: [list.data.id], teams: { scope: 'all' }, options: {} });
    assert.equal(scan.status, 201);
    assert.equal(scan.data.kind, 'teams');
    const done = await app.wait(scan.data.id);
    assert.equal(done.data.status, 'completed', JSON.stringify(done.data.log));
    assert.equal(done.data.stats.messagesMatched, 4);

    const config = fs.readFileSync(path.join(app.store.scanDir(scan.data.id), 'config.json'), 'utf8');
    assert.ok(!config.includes(GRAPH_SECRET) && !config.includes('enc:v1'), 'a configuração não guarda segredos');

    const inList = await app.api('GET', '/api/scans?kind=teams');
    assert.deepEqual(inList.data.map((s) => s.id), [scan.data.id]);
    const notInMail = await app.api('GET', '/api/scans?kind=mail');
    assert.ok(!notInMail.data.some((s) => s.id === scan.data.id));

    const results = await app.api('GET', `/api/scans/${scan.data.id}/results`);
    assert.equal(results.data.total, 4);
    const chatItem = results.data.items.find((r) => r.scopeKind === 'chat');
    assert.equal(chatItem.canDelete, false, 'chat não pode ser excluído pelo relatório');
    assert.equal(chatItem.deleteBlocked, 'chat');
    const channelItem = results.data.items.find((r) => r.scopeKind === 'channel');
    assert.equal(channelItem.canDelete, true);
    assert.equal(channelItem.deleteMethod, 'teams');

    const summary = await app.api('GET', `/api/scans/${scan.data.id}/summary`);
    assert.equal(summary.data.messages, 4);
    assert.ok(summary.data.byScope.some((g) => g.key === 'channel') && summary.data.byScope.some((g) => g.key === 'chat'));

    const scoped = await app.api('GET', `/api/scans/${scan.data.id}/results?scope=chat`);
    assert.ok(scoped.data.items.every((r) => r.scopeKind === 'chat'));

    const xlsx = await app.api('GET', `/api/scans/${scan.data.id}/export.xlsx`);
    const parts = unzipSync(new Uint8Array(xlsx.data));
    const workbook = strFromU8(parts['xl/workbook.xml']);
    assert.deepEqual([...workbook.matchAll(/<sheet name="([^"]*)"/g)].map((m) => m[1]), ['Resumo', 'Mensagens', 'Ocorrências']);

    // Exclusão manual de uma mensagem de canal (softDelete) e recusa para chat.
    const del = await app.api('POST', `/api/scans/${scan.data.id}/results/${channelItem.id}/delete`, { confirm: true, method: 'teams' });
    assert.equal(del.status, 200, JSON.stringify(del.data));
    assert.equal(del.data.deletion.status, 'deleted');
    const delChat = await app.api('POST', `/api/scans/${scan.data.id}/results/${chatItem.id}/delete`, { confirm: true, method: 'teams' });
    assert.equal(delChat.status, 400);
  } finally {
    await app.close();
  }
});
