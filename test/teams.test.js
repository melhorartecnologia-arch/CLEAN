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

async function startApp(endpoints = mocks.endpoints) {
  const store = await new Store(path.join(root, `data-${Math.random().toString(36).slice(2)}`)).init();
  const manager = new ScanManager(store, { mailEndpoints: endpoints });
  const app = createApp({ store, manager, config: { authUser: '', authPassword: '', mailEndpoints: endpoints } });
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

test('API: visualizador ao vivo — resolver usuário, conversas, mensagens e respostas', async () => {
  const app = await startApp();
  try {
    const source = await app.api('POST', '/api/mail-sources', { name: 'Microsoft 365', type: 'graph', scope: 'all', graph: { tenantId: GRAPH_TENANT, clientId: GRAPH_CLIENT, clientSecret: GRAPH_SECRET } });
    const sid = source.data.id;

    const user = await app.api('GET', `/api/teams-live/${sid}/user?address=ana@contoso.com`);
    assert.equal(user.status, 200);
    assert.equal(user.data.user.id, 'u-ana');
    assert.equal(user.data.user.address, 'ana@contoso.com');

    const notFound = await app.api('GET', `/api/teams-live/${sid}/user?address=ninguem@contoso.com`);
    assert.equal(notFound.status, 404);

    const conv = await app.api('GET', `/api/teams-live/${sid}/conversations?userId=u-ana`);
    assert.equal(conv.status, 200);
    assert.deepEqual(conv.data.chats.map((c) => c.id), ['chat1']);
    assert.ok(conv.data.chats[0].label.includes('Ana') || conv.data.chats[0].label.includes('Bia'));
    assert.deepEqual(conv.data.teams.map((t) => t.name), ['Engenharia']);
    assert.deepEqual(conv.data.teams[0].channels.map((c) => c.name).sort(), ['Aleatório', 'Geral']);

    const chatMsgs = await app.api('GET', `/api/teams-live/${sid}/messages?kind=chat&chatId=chat1`);
    assert.equal(chatMsgs.status, 200);
    assert.deepEqual(chatMsgs.data.items.map((m) => m.id), ['cm1']);
    assert.match(chatMsgs.data.items[0].text, /confidencial/);

    const t1 = conv.data.teams[0].id;
    const c1 = conv.data.teams[0].channels.find((c) => c.name === 'Geral').id;
    const chanMsgs = await app.api('GET', `/api/teams-live/${sid}/messages?kind=channel&teamId=${t1}&channelId=${c1}`);
    assert.deepEqual(chanMsgs.data.items.map((m) => m.id).sort(), ['m1', 'm2'], 'mensagens raiz; a de sistema é ignorada');

    const replies = await app.api('GET', `/api/teams-live/${sid}/replies?teamId=${t1}&channelId=${c1}&messageId=m1`);
    assert.deepEqual(replies.data.items.map((m) => m.id), ['r1']);
    assert.match(replies.data.items[0].text, /529\.982\.247-25/);

    // Somente leitura: nenhuma resposta vaza segredo e nada é gravado (sem análise criada).
    assert.ok(!JSON.stringify([user.data, conv.data, chatMsgs.data]).includes(GRAPH_SECRET));
    const scans = await app.api('GET', '/api/scans');
    assert.equal(scans.data.length, 0, 'o visualizador ao vivo não cria análises');
  } finally {
    await app.close();
  }
});

test('API: escopo por lista sem equipes/usuários é recusado (não varre nada em silêncio)', async () => {
  const app = await startApp();
  try {
    const source = await app.api('POST', '/api/mail-sources', { name: 'M365', type: 'graph', scope: 'all', graph: { tenantId: GRAPH_TENANT, clientId: GRAPH_CLIENT, clientSecret: GRAPH_SECRET } });
    const list = await app.api('POST', '/api/lists', { name: 'L', terms: [{ type: 'text', value: 'x' }] });
    const common = { kind: 'teams', sourceIds: [source.data.id], listIds: [list.data.id] };
    const noTeams = await app.api('POST', '/api/scans', { ...common, teams: { scope: 'list', teamIds: '', userEmails: '' }, options: { scanChannels: true, scanChats: true } });
    assert.equal(noTeams.status, 400);
    assert.match(noTeams.data.error, /informe ao menos uma equipe/i);
    const noUsers = await app.api('POST', '/api/scans', { ...common, teams: { scope: 'list', teamIds: 't1', userEmails: '' }, options: { scanChannels: true, scanChats: true } });
    assert.equal(noUsers.status, 400);
    assert.match(noUsers.data.error, /informe ao menos um usuário/i);
    // Só canais, com a equipe informada: aceito.
    const ok = await app.api('POST', '/api/scans', { ...common, teams: { scope: 'list', teamIds: 't1', userEmails: '' }, options: { scanChannels: true, scanChats: false } });
    assert.equal(ok.status, 201);
  } finally {
    await app.close();
  }
});

test('API: visualizador ao vivo — validações e proteção de paginação (SSRF)', async () => {
  const app = await startApp();
  try {
    const source = await app.api('POST', '/api/mail-sources', { name: 'Microsoft 365', type: 'graph', scope: 'all', graph: { tenantId: GRAPH_TENANT, clientId: GRAPH_CLIENT, clientSecret: GRAPH_SECRET } });
    const sid = source.data.id;

    // Conexão inexistente: 404.
    const missing = await app.api('GET', '/api/teams-live/nao-existe/user?address=ana@contoso.com');
    assert.equal(missing.status, 404);

    // Parâmetros obrigatórios.
    assert.equal((await app.api('GET', `/api/teams-live/${sid}/user`)).status, 400);
    assert.equal((await app.api('GET', `/api/teams-live/${sid}/conversations`)).status, 400);
    assert.equal((await app.api('GET', `/api/teams-live/${sid}/messages?kind=disco`)).status, 400);
    assert.equal((await app.api('GET', `/api/teams-live/${sid}/messages?kind=channel&teamId=t1`)).status, 400);

    // Proteção SSRF: um token de continuação que não é um nextLink do Graph é recusado (400) e
    // nunca é usado para buscar outro endereço.
    const evil = encodeURIComponent('http://169.254.169.254/latest/meta-data/');
    const badChats = await app.api('GET', `/api/teams-live/${sid}/chats?userId=u-ana&next=${evil}`);
    assert.equal(badChats.status, 400);
    assert.match(badChats.data.error, /pagina[çc][ãa]o inv[áa]lida/i);
    const badMsgs = await app.api('GET', `/api/teams-live/${sid}/messages?kind=chat&chatId=chat1&next=${evil}`);
    assert.equal(badMsgs.status, 400);
    assert.match(badMsgs.data.error, /pagina[çc][ãa]o inv[áa]lida/i);
    const badHist = await app.api('GET', `/api/teams-live/${sid}/history?kind=chat&chatId=chat1&next=${evil}`);
    assert.equal(badHist.status, 400);
    assert.match(badHist.data.error, /pagina[çc][ãa]o inv[áa]lida/i);
  } finally {
    await app.close();
  }
});

test('API: visualizador ao vivo — uma equipe sem acesso aos canais não derruba o restante', async () => {
  const graph = {
    tenant: GRAPH_TENANT,
    clientId: GRAPH_CLIENT,
    secret: GRAPH_SECRET,
    users: [{ id: 'u-ana', mail: 'ana@contoso.com', displayName: 'Ana' }],
    teamsData: {
      teams: [
        { id: 't-ok', displayName: 'Boa', channels: [{ id: 'c-ok', displayName: 'Geral', membershipType: 'standard', messages: [] }] },
        { id: 't-bad', displayName: 'Sem acesso', channelsError: 403, channels: [] },
      ],
      chats: { 'u-ana': ['chatX'] },
      chatsById: { chatX: { id: 'chatX', chatType: 'oneOnOne', members: [{ displayName: 'Ana', email: 'ana@contoso.com' }], messages: [] } },
    },
  };
  const mock = await startMockApis({ graph });
  const app = await startApp(mock.endpoints);
  try {
    const source = await app.api('POST', '/api/mail-sources', { name: 'Microsoft 365', type: 'graph', scope: 'all', graph: { tenantId: GRAPH_TENANT, clientId: GRAPH_CLIENT, clientSecret: GRAPH_SECRET } });
    const conv = await app.api('GET', `/api/teams-live/${source.data.id}/conversations?userId=u-ana`);
    assert.equal(conv.status, 200, 'a conversa abre mesmo com uma equipe sem acesso');
    assert.deepEqual(conv.data.chats.map((c) => c.id), ['chatX'], 'os chats continuam');
    const ok = conv.data.teams.find((t) => t.id === 't-ok');
    const bad = conv.data.teams.find((t) => t.id === 't-bad');
    assert.deepEqual(ok.channels.map((c) => c.name), ['Geral'], 'a equipe boa mantém os canais');
    assert.equal(bad.error, true, 'a equipe sem acesso é marcada com erro');
    assert.deepEqual(bad.channels, [], 'e sem canais');
  } finally {
    await app.close();
    await mock.close();
  }
});

test('API: visualizador ao vivo — imagens embutidas (proxy) e busca geral', async () => {
  const IMG = Buffer.from('bytes-da-imagem-png', 'utf8');
  const imgTag = (ids) => `<img src="https://graph.microsoft.com/v1.0/teams/t1/channels/c1/messages/m1/hostedContents/${ids}/$value">`;
  const graph = {
    tenant: GRAPH_TENANT,
    clientId: GRAPH_CLIENT,
    secret: GRAPH_SECRET,
    users: [{ id: 'u-ana', mail: 'ana@contoso.com', displayName: 'Ana' }],
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
                { id: 'm1', from: 'Ana', contentType: 'html', body: `<div>Veja o relatório anual ${imgTag('HC1')}</div>`, createdDateTime: '2026-09-20T10:00:00Z', hostedContents: { HC1: { contentType: 'image/png', content: IMG } } },
              ],
            },
          ],
        },
        { id: 't2', displayName: 'Sem acesso', channelsError: 403, channels: [] },
      ],
      chats: { 'u-ana': ['chat1'] },
      chatsById: {
        chat1: {
          id: 'chat1',
          chatType: 'oneOnOne',
          members: [{ displayName: 'Ana', email: 'ana@contoso.com' }, { displayName: 'Bia', email: 'bia@contoso.com' }],
          messages: [{ id: 'cm1', from: 'Bia', contentType: 'html', body: '<p>Bom dia! Tudo certo com o contrato?</p>', createdDateTime: '2026-09-23T10:00:00Z' }],
        },
      },
    },
  };
  const mock = await startMockApis({ graph });
  const app = await startApp(mock.endpoints);
  try {
    const source = await app.api('POST', '/api/mail-sources', { name: 'Microsoft 365', type: 'graph', scope: 'all', graph: { tenantId: GRAPH_TENANT, clientId: GRAPH_CLIENT, clientSecret: GRAPH_SECRET } });
    const sid = source.data.id;

    // Imagens embutidas: a mensagem do canal expõe o id da imagem (sem a URL do Graph).
    const chan = await app.api('GET', `/api/teams-live/${sid}/messages?kind=channel&teamId=t1&channelId=c1`);
    const m1 = chan.data.items.find((x) => x.id === 'm1');
    assert.deepEqual(m1.images.map((i) => i.hostedId), ['HC1']);
    assert.ok(!JSON.stringify(chan.data).includes('graph.microsoft.com'), 'a URL do Graph não vaza para o navegador');

    // O proxy serve os bytes com o content-type de imagem (e nada de token).
    const port = app.srv.address().port;
    const raw = await fetch(`http://127.0.0.1:${port}/api/teams-live/${sid}/image?kind=channel&teamId=t1&channelId=c1&messageId=m1&hostedId=HC1`, { headers: { 'X-CLEAN': '1' } });
    assert.equal(raw.status, 200);
    assert.match(raw.headers.get('content-type') || '', /^image\/png/);
    assert.ok(Buffer.from(await raw.arrayBuffer()).equals(IMG), 'os bytes da imagem chegam íntegros');

    // Busca geral: acha no chat (contrato) e no canal (relatório, sem acento), ignora o que não existe.
    const byChat = await app.api('GET', `/api/teams-live/${sid}/search?userId=u-ana&q=${encodeURIComponent('contrato')}`);
    assert.equal(byChat.status, 200);
    assert.ok(byChat.data.matches.some((x) => x.id === 'cm1' && x.conv.kind === 'chat'));
    const byChannel = await app.api('GET', `/api/teams-live/${sid}/search?userId=u-ana&q=relatorio`);
    assert.ok(byChannel.data.matches.some((x) => x.id === 'm1' && x.conv.kind === 'channel'), 'busca sem acento acha "relatório"');
    assert.equal(byChannel.data.truncated, true, 'a equipe sem canais listados torna a busca limitada (truncated)');
    const none = await app.api('GET', `/api/teams-live/${sid}/search?userId=u-ana&q=inexistentexyz`);
    assert.deepEqual(none.data.matches, []);
    const short = await app.api('GET', `/api/teams-live/${sid}/search?userId=u-ana&q=a`);
    assert.equal(short.status, 400, 'termo curto é recusado');

    // Somente leitura: nada foi gravado.
    const scans = await app.api('GET', '/api/scans');
    assert.equal(scans.data.length, 0);
  } finally {
    await app.close();
    await mock.close();
  }
});

test('API: visualizador ao vivo — histórico (/history) e citação (resposta a uma mensagem)', async () => {
  const quoteRef = JSON.stringify({ messageId: 'cm1', messagePreview: 'podemos revisar o contrato?', messageSender: { user: { displayName: 'Bia' } } });
  const graph = {
    tenant: GRAPH_TENANT,
    clientId: GRAPH_CLIENT,
    secret: GRAPH_SECRET,
    users: [{ id: 'u-ana', mail: 'ana@contoso.com', displayName: 'Ana' }],
    teamsData: {
      teams: [],
      chats: { 'u-ana': ['chat1'] },
      chatsById: {
        chat1: {
          id: 'chat1',
          chatType: 'oneOnOne',
          members: [{ displayName: 'Ana', email: 'ana@contoso.com' }, { displayName: 'Bia', email: 'bia@contoso.com' }],
          messages: [
            { id: 'cm1', from: 'Bia', contentType: 'html', body: '<p>podemos revisar o contrato?</p>', createdDateTime: '2026-09-23T09:00:00Z' },
            { id: 'cm2', from: 'Ana', contentType: 'html', body: '<blockquote>podemos revisar o contrato?</blockquote><p>sim, hoje à tarde</p>', createdDateTime: '2026-09-23T10:00:00Z', attachments: [{ id: 'ref1', contentType: 'messageReference', content: quoteRef }] },
          ],
        },
      },
    },
  };
  const mock = await startMockApis({ graph });
  const app = await startApp(mock.endpoints);
  try {
    const source = await app.api('POST', '/api/mail-sources', { name: 'Microsoft 365', type: 'graph', scope: 'all', graph: { tenantId: GRAPH_TENANT, clientId: GRAPH_CLIENT, clientSecret: GRAPH_SECRET } });
    const sid = source.data.id;

    // /history devolve as mensagens da conversa numa só resposta, sem continuação.
    const hist = await app.api('GET', `/api/teams-live/${sid}/history?kind=chat&chatId=chat1`);
    assert.equal(hist.status, 200);
    assert.deepEqual(hist.data.items.map((m) => m.id).sort(), ['cm1', 'cm2']);
    assert.equal(hist.data.next, null);

    // Citação: a resposta a uma mensagem é reconhecida (remetente + prévia), separada dos anexos, e
    // o trecho citado não se repete no corpo.
    const cm2 = hist.data.items.find((m) => m.id === 'cm2');
    assert.ok(cm2.quote, 'a mensagem tem uma citação');
    assert.equal(cm2.quote.sender, 'Bia');
    assert.match(cm2.quote.preview, /revisar o contrato/);
    assert.deepEqual(cm2.attachments, [], 'a referência não vira anexo comum');
    assert.match(cm2.text, /sim, hoje/);
    assert.ok(!/revisar o contrato/.test(cm2.text), 'o trecho citado foi retirado do corpo');

    // history exige a conversa.
    assert.equal((await app.api('GET', `/api/teams-live/${sid}/history?kind=chat`)).status, 400);
  } finally {
    await app.close();
    await mock.close();
  }
});
