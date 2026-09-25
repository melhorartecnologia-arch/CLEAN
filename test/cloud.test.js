// OneDrive e SharePoint: conector (contas, sites, subsites, bibliotecas, pastas, download e
// exclusão), análise (nome, conteúdo, último usuário, exclusão automática) e API (cadastro sem
// segredos, teste da conexão, análise e exclusão pelo relatório), com o simulador do Microsoft Graph.
import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import crypto from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { Scanner } from '../src/scan/scanner.js';
import { DrivesConnector } from '../src/cloud/drives.js';
import { createApp } from '../src/app.js';
import { Store } from '../src/store.js';
import { ScanManager } from '../src/scan/manager.js';
import { startMockApis } from './helpers/mock-apis.js';

const TENANT = 'contoso.onmicrosoft.com';
const CLIENT = '11111111-2222-3333-4444-555555555555';
const SECRET = 'segredo';
const TERMS = [
  { id: 'l:conf', type: 'text', value: 'confidencial', listName: 'RH' },
  { id: 'l:sal', type: 'text', value: 'salário', listName: 'RH' },
];
const LIST_TERMS = TERMS.map(({ id, listName, ...t }) => t);
const DOCX = fs.readFileSync(new URL('./fixtures/doc.docx', import.meta.url));

let root;
const cleanups = [];
before(() => {
  root = fs.mkdtempSync(path.join(os.tmpdir(), 'clean-cloud-'));
});
after(async () => {
  for (const fn of cleanups.reverse()) await fn();
  fs.rmSync(root, { recursive: true, force: true });
});

const ANA = { name: 'Ana Souza', email: 'ana@contoso.com' };
const BRUNO = { name: 'Bruno Lima', email: 'bruno@contoso.com' };
const file = (id, name, text, extra = {}) => ({ id, name, content: Buffer.isBuffer(text) ? text : Buffer.from(text, 'utf8'), lastModifiedBy: ANA, createdBy: BRUNO, ...extra });
const folder = (id, name, children) => ({ id, name, children });

/** Dados do simulador: contas de OneDrive e sites do SharePoint com bibliotecas e arquivos. */
function world() {
  const drives = {
    'd-ana': {
      id: 'd-ana',
      name: 'OneDrive',
      webUrl: 'https://contoso-my.sharepoint.com/personal/ana_contoso_com/Documents',
      items: [
        folder('f-docs', 'Documentos', [
          file('i-folha', 'folha.txt', 'salário de todos'),
          file('i-limpo', 'limpo.txt', 'nada aqui'),
          file('i-docx', 'Contrato João.docx', DOCX, { lastModifiedBy: BRUNO }),
          file('i-temp', '~$Contrato.docx', 'confidencial'),
          folder('f-antigo', 'Arquivo morto', [file('i-velho', 'velho.txt', 'confidencial antigo')]),
        ]),
        { ...file('i-atalho', 'Atalho', 'confidencial'), remote: true },
      ],
    },
    'd-carla': {
      id: 'd-carla',
      name: 'OneDrive',
      webUrl: 'https://contoso-my.sharepoint.com/personal/carla_contoso_com/Documents',
      items: [file('i-conf', 'confidencial.txt', 'sem termos no texto', { lastModifiedBy: { name: 'Carla Dias', email: 'carla@contoso.com' } }), file('i-grande', 'grande.txt', `confidencial ${'x'.repeat(1.5 * 1048576)}`)],
    },
    'd-root': { id: 'd-root', name: 'Documentos', webUrl: 'https://contoso.sharepoint.com/Shared Documents', items: [file('i-raiz', 'politica.txt', 'material confidencial')] },
    'd-fin': {
      id: 'd-fin',
      name: 'Documentos Compartilhados',
      webUrl: 'https://contoso.sharepoint.com/sites/Financeiro/Shared Documents',
      items: [folder('f-2026', '2026', [file('i-fin', 'salarios.txt', 'salário da diretoria'), file('i-fin2', 'orcamento.txt', 'nada')])],
    },
    'd-fin-assets': { id: 'd-fin-assets', name: 'Site Assets', webUrl: 'https://contoso.sharepoint.com/sites/Financeiro/SiteAssets', items: [file('i-asset', 'confidencial.css', 'x')] },
    'd-contratos': { id: 'd-contratos', name: 'Documentos', webUrl: 'https://contoso.sharepoint.com/sites/Financeiro/Contratos/Shared Documents', items: [file('i-ctr', 'contrato.txt', 'cláusula confidencial')] },
    'd-jur': { id: 'd-jur', name: 'Documentos', webUrl: 'https://contoso.sharepoint.com/sites/Juridico/Shared Documents', items: [file('i-jur', 'processo.txt', 'confidencial')] },
  };
  const sites = [
    { id: 'contoso.sharepoint.com,r1,r2', name: 'contoso', displayName: 'Comunicação', webUrl: 'https://contoso.sharepoint.com', driveIds: ['d-root'] },
    { id: 'contoso.sharepoint.com,f1,f2', name: 'Financeiro', displayName: 'Financeiro', webUrl: 'https://contoso.sharepoint.com/sites/Financeiro', driveIds: ['d-fin', 'd-fin-assets'] },
    { id: 'contoso.sharepoint.com,c1,c2', name: 'Contratos', displayName: 'Contratos', webUrl: 'https://contoso.sharepoint.com/sites/Financeiro/Contratos', driveIds: ['d-contratos'], parent: 'contoso.sharepoint.com,f1,f2' },
    { id: 'contoso.sharepoint.com,j1,j2', name: 'Juridico', displayName: 'Jurídico', webUrl: 'https://contoso.sharepoint.com/sites/Juridico', driveIds: ['d-jur'] },
    { id: 'contoso-my.sharepoint.com,p1,p2', name: 'ana', displayName: 'Ana Souza', webUrl: 'https://contoso-my.sharepoint.com/personal/ana_contoso_com', driveIds: ['d-ana'], personal: true },
  ];
  const users = [
    { id: 'u-ana', mail: 'ana@contoso.com', displayName: 'Ana Souza', driveId: 'd-ana', folders: [], messages: {} },
    { id: 'u-bruno', mail: 'bruno@contoso.com', displayName: 'Bruno Lima', folders: [], messages: {} },
    { id: 'u-carla', mail: 'carla@contoso.com', displayName: 'Carla Dias', driveId: 'd-carla', folders: [], messages: {} },
  ];
  return { tenant: TENANT, clientId: CLIENT, secret: SECRET, users, drives, sites };
}

async function withGraph(fn) {
  const data = world();
  const mock = await startMockApis({ graph: data });
  try {
    return await fn(data, mock.endpoints);
  } finally {
    await mock.close();
  }
}

const repo = (type, extra = {}) => ({
  id: `r-${type}`,
  type,
  name: type === 'onedrive' ? 'OneDrive da empresa' : 'SharePoint da empresa',
  path: '',
  exclude: ['Arquivo morto', 'Site Assets'],
  allowDelete: false,
  deleteMode: 'trash',
  graph: { tenantId: TENANT, clientId: CLIENT },
  secrets: { clientSecret: SECRET },
  cloud: { scope: 'all', accounts: [], sites: [], exclude: [] },
  ...extra,
});

async function scan(repositories, endpoints, options = {}, deps = {}) {
  const messages = [];
  const stats = await new Scanner({ repositories, terms: TERMS, options: { maxFileSizeMB: 1, ...options }, endpoints }, (m) => messages.push(m), deps).run();
  return {
    stats,
    messages,
    records: messages.filter((m) => m.type === 'results').flatMap((m) => m.records),
    errors: messages.filter((m) => m.type === 'errors').flatMap((m) => m.items),
    events: messages.filter((m) => m.type === 'deletions').flatMap((m) => m.items),
    logs: messages.filter((m) => m.type === 'log').map((m) => m.message),
  };
}

test('OneDrive: todas as contas, pastas ignoradas, atalhos, conteúdo e último usuário do Microsoft 365', async () => {
  await withGraph(async (data, endpoints) => {
    const run = await scan([repo('onedrive')], endpoints);
    const byName = Object.fromEntries(run.records.map((r) => [r.name, r]));
    assert.deepEqual(Object.keys(byName).sort(), ['Contrato João.docx', 'confidencial.txt', 'folha.txt', 'grande.txt']);
    assert.equal(run.stats.accountsSkipped, 1, 'bruno não tem OneDrive');
    assert.equal(run.stats.libraries, 2);
    assert.ok(run.logs.some((l) => /bruno@contoso\.com: ignorado, sem OneDrive/.test(l)));

    const folha = byName['folha.txt'];
    assert.equal(folha.lastUser, 'ana@contoso.com');
    assert.equal(folha.lastUserSource, 'cloud');
    assert.equal(folha.owner, 'ana@contoso.com', 'dono do OneDrive');
    assert.equal(folha.relativePath, 'Documentos/folha.txt');
    assert.equal(folha.path, 'https://contoso-my.sharepoint.com/personal/ana_contoso_com/Documents/Documentos/folha.txt');
    assert.deepEqual(
      { kind: folha.cloud.kind, account: folha.cloud.account, library: folha.cloud.library, driveId: folha.cloud.driveId, itemId: folha.cloud.itemId },
      { kind: 'onedrive', account: 'ana@contoso.com', library: 'OneDrive', driveId: 'd-ana', itemId: 'i-folha' },
    );
    assert.match(folha.cloud.cTag, /i-folha/);

    // O conteúdo do .docx é lido; o último usuário vem do Microsoft 365, não dos metadados.
    const docx = byName['Contrato João.docx'];
    assert.equal(docx.contentStatus, 'ok');
    assert.equal(docx.lastUser, 'bruno@contoso.com');
    assert.equal(docx.metadata.lastModifiedBy, 'Marcos Revisor');
    assert.equal(byName['confidencial.txt'].matches[0].location, 'name');
    // Texto maior que o limite: só o início é baixado.
    assert.equal(byName['grande.txt'].contentStatus, 'partial');
    assert.ok(!run.records.some((r) => /velho|Atalho|~\$/.test(r.name)));
  });
});

test('OneDrive: contas da lista, conta inexistente e conta ignorada', async () => {
  await withGraph(async (data, endpoints) => {
    const cloud = { scope: 'list', accounts: ['ana@contoso.com', 'carla@contoso.com', 'naoexiste@contoso.com'], sites: [], exclude: ['carla@*'] };
    const run = await scan([repo('onedrive', { cloud })], endpoints);
    assert.deepEqual(run.records.map((r) => r.name).sort(), ['Contrato João.docx', 'folha.txt']);
    assert.ok(run.records.every((r) => r.cloud.account === 'ana@contoso.com'));
    assert.ok(run.errors.some((e) => e.path === 'naoexiste@contoso.com' && /não foi encontrada/.test(e.message)));
  });
});

test('SharePoint: todos os sites (páginas), subsites, sem os pessoais, bibliotecas e sites ignorados', async () => {
  await withGraph(async (data, endpoints) => {
    const cloud = { scope: 'all', accounts: [], sites: [], exclude: ['https://contoso.sharepoint.com/sites/Juridico'] };
    const run = await scan([repo('sharepoint', { cloud })], endpoints);
    const names = run.records.map((r) => `${r.cloud.accountName} › ${r.cloud.library} › ${r.relativePath}`).sort();
    assert.deepEqual(names, [
      'Comunicação › Documentos › politica.txt',
      'Contratos › Documentos › contrato.txt',
      'Financeiro › Documentos Compartilhados › 2026/salarios.txt',
    ]);
    const salarios = run.records.find((r) => r.name === 'salarios.txt');
    assert.equal(salarios.owner, 'bruno@contoso.com', 'no SharePoint, quem criou');
    assert.equal(salarios.path, 'https://contoso.sharepoint.com/sites/Financeiro/Shared Documents/2026/salarios.txt');

    // Sem getAllSites (locatários mais antigos), usa a busca de sites.
    data.noGetAllSites = true;
    const fallback = await scan([repo('sharepoint', { cloud: { ...cloud, exclude: [] } })], endpoints);
    assert.equal(fallback.records.filter((r) => r.cloud.accountName === 'Jurídico').length, 1);
  });
});

test('SharePoint: sites da lista (também um link de biblioteca copiado do navegador) e site inexistente', async () => {
  await withGraph(async (data, endpoints) => {
    const cloud = {
      scope: 'list',
      accounts: [],
      sites: ['https://contoso.sharepoint.com/sites/Financeiro/Shared Documents/Forms/AllItems.aspx', 'https://contoso.sharepoint.com/sites/NaoExiste'],
      exclude: [],
    };
    const run = await scan([repo('sharepoint', { cloud })], endpoints);
    assert.deepEqual(run.records.map((r) => r.name).sort(), ['contrato.txt', 'salarios.txt']);
    assert.ok(run.errors.some((e) => /sites\/NaoExiste/.test(e.path) && /Site não encontrado/.test(e.message)));
  });
});

test('exclusão automática na nuvem: lixeira com If-Match, alterado, bloqueado, definitiva e proteção por cadastro', async () => {
  await withGraph(async (data, endpoints) => {
    // Durante a análise, o arquivo "folha.txt" muda (nova versão) antes da exclusão.
    const factory = (r, options) => {
      const connector = new DrivesConnector(r, options);
      const original = connector.deleteItem.bind(connector);
      connector.deleteItem = (target, mode, opts) => {
        if (target.itemId === 'i-folha') data.drives['d-ana'].items[0].children[0].version = 2;
        return original(target, mode, opts);
      };
      return connector;
    };
    data.lockedItems = new Set(['i-conf']);
    const cloud = { scope: 'list', accounts: ['ana@contoso.com', 'carla@contoso.com'], sites: [], exclude: [] };
    const run = await scan([repo('onedrive', { cloud, allowDelete: true })], endpoints, { deleteMatches: true }, { cloudConnectorFactory: factory });
    const nameOf = new Map(run.records.map((r) => [r.id, r.name]));
    const status = Object.fromEntries(run.events.map((e) => [nameOf.get(e.recordId), e.status]));
    assert.deepEqual(status, { 'folha.txt': 'changed', 'Contrato João.docx': 'deleted', 'confidencial.txt': 'failed', 'grande.txt': 'deleted' });
    assert.equal(run.stats.deleted, 2);
    assert.equal(run.stats.deleteChanged, 1);
    assert.equal(run.stats.deleteErrors, 1);
    assert.ok(run.events.every((e) => e.method === 'trash' && e.mode === 'auto'));
    const trashed = data.driveDeleted.filter((d) => d.how === 'trash');
    assert.ok(trashed.every((d) => /^"c:\{/.test(d.ifMatch)), 'exclusão condicionada à versão analisada (If-Match)');
    assert.ok(data.recycle.some((r) => r.item.id === 'i-grande'));
    assert.match(run.errors.find((e) => /confidencial\.txt/.test(e.path)).message, /bloqueado/);

    // Exclusão definitiva e proteção por outro repositório (lista, sem exclusão) do mesmo locatário.
    const data2 = world();
    Object.assign(data, { drives: data2.drives, driveDeleted: [], recycle: [], lockedItems: new Set() });
    const keep = { all: null, accounts: [{ value: 'carla@contoso.com', error: 'Protegido pelo repositório "Diretoria", que não permite exclusão.' }], sites: [] };
    const permanent = await scan([repo('onedrive', { allowDelete: true, deleteMode: 'permanent', keep })], endpoints, { deleteMatches: true });
    const names = new Map(permanent.records.map((r) => [r.id, r.name]));
    const byName = Object.fromEntries(permanent.events.map((e) => [names.get(e.recordId), e]));
    assert.equal(byName['confidencial.txt'].status, 'failed');
    assert.match(byName['confidencial.txt'].error, /Diretoria/);
    assert.equal(byName['grande.txt'].status, 'failed');
    assert.equal(byName['folha.txt'].status, 'deleted');
    assert.deepEqual(data.driveDeleted.map((d) => [d.id, d.how]).sort(), [['i-docx', 'permanent'], ['i-folha', 'permanent']]);
    assert.ok(permanent.events.every((e) => e.method === 'permanent'));
  });
});

// -- API ------------------------------------------------------------------------------------------

async function startApp(mailEndpoints) {
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
    let s;
    for (let i = 0; i < 300; i++) {
      s = (await api('GET', `/api/scans/${id}`)).data;
      if (!['queued', 'running'].includes(s.status)) return s;
      await new Promise((r) => setTimeout(r, 100));
    }
    return s;
  };
  cleanups.push(async () => {
    srv.close();
    await manager.shutdown();
  });
  return { store, api, wait };
}

test('API: cadastro sem segredos, credenciais de uma conexão de e-mail, teste, análise e exclusão pelo relatório', async () => {
  const data = world();
  const mock = await startMockApis({ graph: data });
  cleanups.push(() => mock.close());
  const { store, api, wait } = await startApp(mock.endpoints);

  // Validações
  const noAccounts = await api('POST', '/api/repositories', { type: 'onedrive', name: 'X', scope: 'list', accounts: '', graph: { tenantId: TENANT, clientId: CLIENT, clientSecret: SECRET } });
  assert.equal(noAccounts.status, 400);
  assert.match(noAccounts.data.error, /ao menos uma conta/);
  const badSite = await api('POST', '/api/repositories', { type: 'sharepoint', name: 'X', scope: 'list', sites: 'http://contoso.sharepoint.com/sites/A', graph: { tenantId: TENANT, clientId: CLIENT, clientSecret: SECRET } });
  assert.equal(badSite.status, 400);
  assert.match(badSite.data.error, /https:\/\//);

  // Credenciais copiadas de uma conexão de e-mail do Microsoft 365
  const mailSource = (await api('POST', '/api/mail-sources', { name: 'M365', type: 'graph', scope: 'list', mailboxes: 'ana@contoso.com', graph: { tenantId: TENANT, clientId: CLIENT, clientSecret: SECRET } })).data;
  const created = await api('POST', '/api/repositories', {
    type: 'onedrive',
    name: 'OneDrive',
    scope: 'list',
    accounts: 'ana@contoso.com\ncarla@contoso.com',
    credentialsFrom: mailSource.id,
    exclude: 'Arquivo morto',
    allowDelete: true,
  });
  assert.equal(created.status, 201, JSON.stringify(created.data));
  const repoData = created.data;
  assert.equal(repoData.type, 'onedrive');
  assert.equal(repoData.deleteMode, 'trash', 'padrão: lixeira do OneDrive');
  assert.equal(repoData.secrets, undefined);
  assert.equal(repoData.graph.hasClientSecret, true);
  assert.equal(repoData.path, 'OneDrive: ana@contoso.com, carla@contoso.com');
  const listed = (await api('GET', '/api/repositories')).data;
  assert.ok(!JSON.stringify(listed).includes(SECRET) && listed.every((r) => r.secrets === undefined));
  const stored = store.getRepository(repoData.id);
  assert.ok(stored.secrets.clientSecret && !stored.secrets.clientSecret.includes(SECRET), 'segredo cifrado no cadastro');

  // Teste da conexão (segredo salvo, com o id do repositório)
  const tested = (await api('POST', '/api/repositories/test', { id: repoData.id, type: 'onedrive', scope: 'list', accounts: 'ana@contoso.com', graph: { tenantId: TENANT, clientId: CLIENT } })).data;
  assert.equal(tested.ok, true, JSON.stringify(tested));
  assert.match(tested.details[0], /OneDrive de ana@contoso\.com: acesso OK/);
  const wrong = (await api('POST', '/api/repositories/test', { type: 'onedrive', scope: 'all', graph: { tenantId: TENANT, clientId: CLIENT, clientSecret: 'errado' } })).data;
  assert.equal(wrong.ok, false);
  assert.match(wrong.message, /Segredo do cliente inválido/);

  // Análise pela API (thread de análise) e exclusão manual
  const list = (await api('POST', '/api/lists', { name: 'L', terms: LIST_TERMS })).data;
  const scanRes = (await api('POST', '/api/scans', { repositoryIds: [repoData.id], listIds: [list.id], options: { maxFileSizeMB: 1 } })).data;
  const done = await wait(scanRes.id);
  assert.equal(done.status, 'completed', JSON.stringify(done.log));
  assert.equal(done.summary.repositories[0].type, 'onedrive');
  const config = fs.readFileSync(path.join(store.scanDir(scanRes.id), 'config.json'), 'utf8');
  assert.ok(!config.includes(SECRET) && !config.includes('clientSecret'), 'o config.json da análise não guarda segredos');
  const items = (await api('GET', `/api/scans/${scanRes.id}/results?sort=name`)).data.items;
  const folha = items.find((r) => r.name === 'folha.txt');
  assert.equal(folha.canDelete, true);
  assert.equal(folha.deleteMethod, 'trash');
  assert.equal(folha.lastUser, 'ana@contoso.com');

  const del = await api('POST', `/api/scans/${scanRes.id}/results/${folha.id}/delete`, { confirm: true, method: 'trash' });
  assert.equal(del.status, 200, JSON.stringify(del.data));
  assert.equal(del.data.deletion.status, 'deleted');
  assert.equal(del.data.deletion.method, 'trash');
  assert.ok(data.recycle.some((r) => r.item.id === 'i-folha'));

  // Arquivo alterado depois da análise: pede confirmação e depois exclui mesmo assim.
  const docx = items.find((r) => r.name === 'Contrato João.docx');
  data.drives['d-ana'].items[0].children.find((i) => i.id === 'i-docx').version = 3;
  const changed = await api('POST', `/api/scans/${scanRes.id}/results/${docx.id}/delete`, { confirm: true, method: 'trash' });
  assert.equal(changed.status, 409);
  assert.equal(changed.data.code, 'changed');
  const forced = await api('POST', `/api/scans/${scanRes.id}/results/${docx.id}/delete`, { confirm: true, method: 'trash', force: true });
  assert.equal(forced.data.deletion.status, 'deleted');

  // Forma de exclusão mudou no cadastro depois de a página ser aberta.
  const put = (extra) => api('PUT', `/api/repositories/${repoData.id}`, { type: 'onedrive', name: 'OneDrive', scope: 'list', accounts: 'ana@contoso.com\ncarla@contoso.com', graph: { tenantId: TENANT, clientId: CLIENT }, exclude: 'Arquivo morto', allowDelete: true, ...extra });
  assert.equal((await put({ deleteMode: 'permanent' })).status, 200);
  const other = items.find((r) => r.name === 'confidencial.txt');
  const mismatch = await api('POST', `/api/scans/${scanRes.id}/results/${other.id}/delete`, { confirm: true, method: 'trash' });
  assert.equal(mismatch.status, 409);
  assert.equal(mismatch.data.code, 'method-changed');
  const permanent = await api('POST', `/api/scans/${scanRes.id}/results/${other.id}/delete`, { confirm: true, method: 'permanent' });
  assert.equal(permanent.data.deletion.status, 'deleted');
  assert.equal(permanent.data.deletion.method, 'permanent');
  assert.ok(data.driveDeleted.some((d) => d.id === 'i-conf' && d.how === 'permanent'));

  // O repositório passou a ser outro tipo: a exclusão pelo relatório fica indisponível.
  assert.equal((await api('PUT', `/api/repositories/${repoData.id}`, { type: 'sharepoint', name: 'OneDrive', scope: 'all', graph: { tenantId: TENANT, clientId: CLIENT }, allowDelete: true })).status, 200);
  const after = (await api('GET', `/api/scans/${scanRes.id}/results?sort=name`)).data.items;
  const grande = after.find((r) => r.name === 'grande.txt');
  assert.equal(grande.canDelete, false);
  assert.equal(grande.deleteBlocked, 'changed');
  const refused = await api('POST', `/api/scans/${scanRes.id}/results/${grande.id}/delete`, { confirm: true });
  assert.equal(refused.status, 409);

  // Voltar a ser pasta local descarta as credenciais salvas.
  const local = await api('PUT', `/api/repositories/${repoData.id}`, { type: 'local', name: 'Pasta', path: root });
  assert.equal(local.status, 200);
  assert.equal(store.getRepository(repoData.id).secrets, null);
  assert.equal(local.data.graph, null);
});
