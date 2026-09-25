// OneDrive e SharePoint: salvaguardas da análise e da exclusão (nome de logon x e-mail, contas
// repetidas, proteção por outros repositórios, arquivo renomeado ou movido, endereço de site com erro,
// downloads interrompidos no tempo esgotado, falha no meio de uma pasta, alcance e credenciais).
import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import crypto from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { DrivesConnector, keptCloud, keptCloudTarget, coveredByRepo, cloudTarget } from '../src/cloud/drives.js';
import { createApp } from '../src/app.js';
import { Store } from '../src/store.js';
import { ScanManager } from '../src/scan/manager.js';
import { startMockApis } from './helpers/mock-apis.js';
import { TENANT, CLIENT, SECRET, LIST_TERMS, world, withGraph, repo, scan, file, folder } from './helpers/cloud-world.js';

let root;
const cleanups = [];
before(() => {
  root = fs.mkdtempSync(path.join(os.tmpdir(), 'clean-cloud-safety-'));
});
after(async () => {
  for (const fn of cleanups.reverse()) await fn();
  fs.rmSync(root, { recursive: true, force: true });
});

const connectorFor = (r, endpoints) => new DrivesConnector(r, { endpoints });

test('OneDrive: conta ignorada pelo nome de logon, conta repetida (e-mail e logon) e proteção pelo logon', async () => {
  await withGraph(async (data, endpoints) => {
    // O nome de logon (UPN) da Carla é diferente do e-mail.
    data.users.find((u) => u.id === 'u-carla').upn = 'c.dias@contoso.onmicrosoft.com';

    // Todas as contas, ignorando a Carla pelo nome de logon.
    const ignored = await scan([repo('onedrive', { cloud: { scope: 'all', accounts: [], sites: [], exclude: ['c.dias@contoso.onmicrosoft.com'] } })], endpoints);
    assert.ok(ignored.records.every((r) => r.cloud.account === 'ana@contoso.com'));

    // A mesma conta pelo e-mail e pelo logon: analisada uma vez.
    const twice = await scan([repo('onedrive', { cloud: { scope: 'list', accounts: ['carla@contoso.com', 'c.dias@contoso.onmicrosoft.com'], sites: [], exclude: [] } })], endpoints);
    assert.deepEqual(twice.records.map((r) => r.name).sort(), ['confidencial.txt', 'grande.txt']);
    assert.ok(twice.records[0].cloud.aliases.includes('c.dias@contoso.onmicrosoft.com'));

    // Outro repositório (sem exclusão, com locatário escrito de outro jeito) protege a Carla pelo logon.
    const protector = { ...repo('onedrive'), id: 'r-dir', name: 'Diretoria', graph: { tenantId: 'CONTOSO.onmicrosoft.com', clientId: CLIENT }, cloud: { scope: 'list', accounts: ['c.dias@contoso.onmicrosoft.com'], sites: [], exclude: [] } };
    const deleter = repo('onedrive', { allowDelete: true });
    const keep = keptCloud(deleter, [deleter, protector]);
    const run = await scan([{ ...deleter, keep }], endpoints, { deleteMatches: true });
    const byName = new Map(run.records.map((r) => [r.id, r.name]));
    const status = Object.fromEntries(run.events.map((e) => [byName.get(e.recordId), e]));
    assert.equal(status['confidencial.txt'].status, 'failed');
    assert.match(status['confidencial.txt'].error, /Diretoria/);
    assert.equal(status['folha.txt'].status, 'deleted');
    assert.ok(!data.driveDeleted.some((d) => d.drive === 'd-carla'));
  });
});

test('proteção: se as contas protegidas não puderem ser conferidas, nada do OneDrive é excluído', async () => {
  await withGraph(async (data, endpoints) => {
    const deleter = repo('onedrive', { allowDelete: true });
    const keep = { all: null, accounts: [{ value: 'diretoria@contoso.com', error: 'x' }], sites: [] };
    const connector = connectorFor(deleter, endpoints);
    connector.resolveUser = async () => {
      throw Object.assign(new Error('Falha de rede'), { status: 0 });
    };
    await connector.resolveKept(keep);
    const blocked = keptCloudTarget({ kind: 'onedrive', driveId: 'd-ana', account: 'ana@contoso.com', aliases: [] }, keep);
    assert.match(blocked.error, /Não foi possível conferir as contas protegidas/);
    // Site raiz protegido não protege os sites em /sites/ (coleções separadas); subsite protegido protege o site acima.
    const sp = keptCloud(repo('sharepoint', { allowDelete: true, cloud: { scope: 'all', accounts: [], sites: [], exclude: [] } }), [
      { ...repo('sharepoint'), id: 'r-raiz', name: 'Raiz', cloud: { scope: 'list', accounts: [], sites: ['https://contoso.sharepoint.com'], exclude: [] } },
      { ...repo('sharepoint'), id: 'r-ctr', name: 'Contratos', cloud: { scope: 'list', accounts: [], sites: ['https://contoso.sharepoint.com/sites/Financeiro/Contratos/Shared Documents/Forms/AllItems.aspx'], exclude: [] } },
    ]);
    assert.equal(keptCloudTarget({ kind: 'sharepoint', account: 'https://contoso.sharepoint.com/sites/Juridico' }, sp), null);
    assert.match(keptCloudTarget({ kind: 'sharepoint', account: 'https://contoso.sharepoint.com' }, sp).error, /Raiz/);
    assert.match(keptCloudTarget({ kind: 'sharepoint', account: 'https://contoso.sharepoint.com/sites/Financeiro/Contratos' }, sp).error, /Contratos/);
    assert.match(keptCloudTarget({ kind: 'sharepoint', account: 'https://contoso.sharepoint.com/sites/Financeiro' }, sp).error, /Contratos/);
  });
});

test('exclusão: arquivo renomeado, movido ou alterado entre a conferência e a exclusão; sem versão registrada', async () => {
  await withGraph(async (data, endpoints) => {
    const connector = connectorFor(repo('onedrive'), endpoints);
    const run = await scan([repo('onedrive', { cloud: { scope: 'list', accounts: ['ana@contoso.com'], sites: [], exclude: [] } })], endpoints);
    const record = (name) => run.records.find((r) => r.name === name);
    const docs = data.drives['d-ana'].items[0];

    // Renomeado depois da análise (o cTag não muda): não é excluído sem confirmação.
    const folha = docs.children.find((i) => i.id === 'i-folha');
    folha.name = 'folha-antiga.txt';
    folha.renamed = true;
    assert.deepEqual(await connector.deleteItem(cloudTarget(record('folha.txt')), 'trash'), { status: 'changed', error: 'O arquivo foi renomeado depois da análise.' });

    // Movido para outra pasta.
    const docx = docs.children.splice(docs.children.findIndex((i) => i.id === 'i-docx'), 1)[0];
    data.drives['d-ana'].items.push(folder('f-outra', 'Outra', [docx]));
    assert.deepEqual(await connector.deleteItem(cloudTarget(record('Contrato João.docx')), 'permanent'), { status: 'changed', error: 'O arquivo foi movido para outra pasta depois da análise.' });

    // Alterado entre a conferência e a exclusão definitiva: o If-Match faz o servidor recusar.
    const grande = await scan([repo('onedrive', { cloud: { scope: 'list', accounts: ['carla@contoso.com'], sites: [], exclude: [] } })], endpoints);
    const target = cloudTarget(grande.records.find((r) => r.name === 'confidencial.txt'));
    const racing = connectorFor(repo('onedrive'), endpoints);
    const originalApi = racing.api.bind(racing);
    racing.api = async (url, options = {}) => {
      const res = await originalApi(url, options);
      if (/\?\$select=id,name,size/.test(url)) data.drives['d-carla'].items[0].version = 2; // nova versão logo após a conferência
      return res;
    };
    assert.equal((await racing.deleteItem(target, 'permanent')).status, 'changed');
    assert.ok(!(data.driveDeleted || []).some((d) => d.id === 'i-conf'));

    // Sem cTag/eTag registrados, confere tamanho e data de modificação.
    const noTag = { ...target, cTag: null, eTag: null, size: 999 };
    assert.equal((await connector.deleteItem(noTag, 'trash')).status, 'changed');
  });
});

test('SharePoint: endereço de subsite com erro não vira o site acima; biblioteca com nome próprio; subsites de site ignorado', async () => {
  await withGraph(async (data, endpoints) => {
    const connector = connectorFor(repo('sharepoint'), endpoints);
    await assert.rejects(connector.resolveSite('https://contoso.sharepoint.com/sites/Financeiro/Contrato'), /Site não encontrado/);
    data.drives['d-fin'].name = 'Contratos Assinados';
    const site = await connector.resolveSite('https://contoso.sharepoint.com/sites/Financeiro/Contratos Assinados/Forms/AllItems.aspx');
    assert.equal(site.webUrl, 'https://contoso.sharepoint.com/sites/Financeiro');

    // Ignorar o site pelo endereço exato ignora também os subsites dele.
    const run = await scan([repo('sharepoint', { cloud: { scope: 'all', accounts: [], sites: [], exclude: ['https://contoso.sharepoint.com/sites/Financeiro'] } })], endpoints);
    assert.ok(run.records.every((r) => !r.cloud.account.includes('/sites/Financeiro')), run.records.map((r) => r.cloud.account).join(', '));
    assert.ok(run.records.some((r) => r.cloud.accountName === 'Jurídico'));
  });
});

test('análise: download interrompido de fato no tempo esgotado; falha no meio de uma pasta mantém as subpastas', async () => {
  await withGraph(async (data, endpoints) => {
    data.slowDownloads = new Set(['i-folha']);
    data.slowMs = 1500;
    const messages = [];
    const { Scanner } = await import('../src/scan/scanner.js');
    const scanner = new Scanner(
      { repositories: [repo('onedrive', { cloud: { scope: 'list', accounts: ['ana@contoso.com'], sites: [], exclude: [] } })], terms: [{ id: 't', type: 'text', value: 'salário' }], options: {}, endpoints, fileTimeoutMs: 300 },
      (m) => messages.push(m),
    );
    await scanner.run();
    await new Promise((r) => setTimeout(r, 100));
    assert.deepEqual(data.abortedDownloads, ['i-folha'], 'o download lento foi interrompido');
    const errors = messages.filter((m) => m.type === 'errors').flatMap((m) => m.items);
    assert.ok(errors.some((e) => /folha\.txt/.test(e.path) && /Tempo esgotado/.test(e.message)));

    // A 2ª página da raiz falha: as pastas da 1ª página continuam sendo analisadas.
    const w = world();
    Object.assign(data, { drives: w.drives, slowDownloads: null, pageSize: 1, failPageOf: 'root' });
    const run = await scan([repo('onedrive', { cloud: { scope: 'list', accounts: ['ana@contoso.com'], sites: [], exclude: [] } })], endpoints);
    assert.deepEqual(run.records.map((r) => r.name).sort(), ['Contrato João.docx', 'folha.txt']);
    assert.ok(run.errors.some((e) => /OneDrive de ana@contoso\.com/.test(e.path)));
  });
});

test('alcance do repositório: conta pelo logon, conta retirada ou ignorada, site da lista com subsites', () => {
  const od = (cloud) => ({ type: 'onedrive', cloud: { scope: 'list', accounts: [], sites: [], exclude: [], ...cloud } });
  const target = { kind: 'onedrive', account: 'carla@contoso.com', aliases: ['carla@contoso.com', 'c.dias@contoso.onmicrosoft.com'] };
  assert.equal(coveredByRepo(od({ accounts: ['C.Dias@contoso.onmicrosoft.com'] }), target), true);
  assert.equal(coveredByRepo(od({ accounts: ['ana@contoso.com'] }), target), false);
  assert.equal(coveredByRepo(od({ scope: 'all', exclude: ['c.dias@*'] }), target), false);
  const sp = { type: 'sharepoint', cloud: { scope: 'list', accounts: [], sites: ['https://contoso.sharepoint.com/sites/Financeiro/Shared Documents/Forms/AllItems.aspx'], exclude: [] } };
  assert.equal(coveredByRepo(sp, { kind: 'sharepoint', account: 'https://contoso.sharepoint.com/sites/Financeiro/Contratos' }), true);
  assert.equal(coveredByRepo(sp, { kind: 'sharepoint', account: 'https://contoso.sharepoint.com/sites/Juridico' }), false);
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

test('API: pasta passou a ser ignorada, conta retirada, credenciais ligadas à conexão de e-mail e OneDrive pessoal na lista de sites', async () => {
  const data = world();
  const mock = await startMockApis({ graph: data });
  cleanups.push(() => mock.close());
  const { store, api, wait } = await startApp(mock.endpoints);
  const graph = { tenantId: TENANT, clientId: CLIENT, clientSecret: SECRET };

  const personal = await api('POST', '/api/repositories', { type: 'sharepoint', name: 'X', scope: 'list', sites: 'https://contoso-my.sharepoint.com/personal/ana_contoso_com', graph });
  assert.equal(personal.status, 400);
  assert.match(personal.data.error, /OneDrive pessoal/);

  const mailSource = (await api('POST', '/api/mail-sources', { name: 'M365', type: 'graph', scope: 'list', mailboxes: 'ana@contoso.com', graph })).data;
  const body = { type: 'onedrive', name: 'OneDrive', scope: 'list', accounts: 'ana@contoso.com\ncarla@contoso.com', credentialsFrom: mailSource.id, allowDelete: true };
  const created = (await api('POST', '/api/repositories', body)).data;
  assert.equal(created.credentialsFrom, mailSource.id);
  const list = (await api('POST', '/api/lists', { name: 'L', terms: LIST_TERMS })).data;
  const scanRes = (await api('POST', '/api/scans', { repositoryIds: [created.id], listIds: [list.id], options: { maxFileSizeMB: 1 } })).data;
  assert.equal((await wait(scanRes.id)).status, 'completed');
  const items = () => api('GET', `/api/scans/${scanRes.id}/results?sort=name`).then((r) => r.data.items);

  // Segredo renovado na conexão de e-mail: vale também para o repositório.
  const before = store.getRepository(created.id).secrets.clientSecret;
  data.secret = 'novo-segredo';
  assert.equal((await api('PUT', `/api/mail-sources/${mailSource.id}`, { name: 'M365', type: 'graph', scope: 'list', mailboxes: 'ana@contoso.com', graph: { ...graph, clientSecret: 'novo-segredo' } })).status, 200);
  assert.notEqual(store.getRepository(created.id).secrets.clientSecret, before);
  const tested = (await api('POST', '/api/repositories/test', { id: created.id, ...body })).data;
  assert.equal(tested.ok, true, JSON.stringify(tested));

  // Trocar o locatário sem informar o segredo de novo é recusado.
  const otherTenant = await api('PUT', `/api/repositories/${created.id}`, { ...body, credentialsFrom: '', graph: { tenantId: '99999999-8888-7777-6666-555555555555', clientId: CLIENT } });
  assert.equal(otherTenant.status, 400);
  assert.match(otherTenant.data.error, /informe o segredo do cliente novamente/);

  // A pasta "Documentos" passou a ser ignorada: os arquivos dela não são excluídos pelo relatório.
  assert.equal((await api('PUT', `/api/repositories/${created.id}`, { ...body, exclude: 'Documentos' })).status, 200);
  const folha = (await items()).find((r) => r.name === 'folha.txt');
  assert.equal(folha.canDelete, false);
  assert.equal(folha.deleteBlocked, 'excluded');
  const refused = await api('POST', `/api/scans/${scanRes.id}/results/${folha.id}/delete`, { confirm: true });
  assert.equal(refused.status, 409);
  assert.match(refused.data.error, /passou a ignorar/);

  // A Carla saiu da lista: os arquivos dela também não.
  assert.equal((await api('PUT', `/api/repositories/${created.id}`, { ...body, accounts: 'ana@contoso.com' })).status, 200);
  const conf = (await items()).find((r) => r.name === 'confidencial.txt');
  assert.equal(conf.deleteBlocked, 'changed');
  assert.equal((await api('POST', `/api/scans/${scanRes.id}/results/${conf.id}/delete`, { confirm: true })).status, 409);
  assert.ok(!(data.driveDeleted || []).length, 'nada foi excluído');

  // A conexão de e-mail foi removida: o repositório continua com as credenciais, sem a ligação.
  assert.equal((await api('DELETE', `/api/mail-sources/${mailSource.id}`)).status, 204);
  assert.equal(store.getRepository(created.id).credentialsFrom, null);
  assert.ok(store.getRepository(created.id).secrets.clientSecret);
});
