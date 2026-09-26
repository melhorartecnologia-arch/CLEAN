import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import crypto from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {
  sanitizeRetention,
  cutoffDate,
  fileDate,
  cloudDate,
  describeRetention,
  patternMatcher,
  ageBucket,
  amountText,
  RetentionError,
} from '../src/retention/policy.js';
import { Scanner } from '../src/scan/scanner.js';
import { MailScanner } from '../src/mail/scanner.js';
import { Store } from '../src/store.js';
import { ScanManager } from '../src/scan/manager.js';
import { Scheduler } from '../src/schedule/scheduler.js';
import { createApp } from '../src/app.js';
import { startMockApis } from './helpers/mock-apis.js';
import { startFakeImap } from './helpers/fake-imap.js';
import { withGraph, repo as cloudRepo, file, folder } from './helpers/cloud-world.js';
import { exportRetentionCsv, exportRetentionHtml } from '../src/report/retention-exports.js';
import { summarizeRetention, filterRecords, filterMailRecords } from '../src/report/model.js';
import { PassThrough } from 'node:stream';

let root;
const stores = [];
const DAY = 86400000;
const OLD = new Date('2015-03-10T12:00:00Z');

before(() => {
  root = fs.mkdtempSync(path.join(os.tmpdir(), 'clean-ret-'));
});

after(async () => {
  await Promise.all(stores.map((s) => s.close()));
  fs.rmSync(root, { recursive: true, force: true });
});

test('política: validação, data de corte, critérios e descrição', () => {
  const fails = (input, kind, pattern, extra) => assert.throws(() => sanitizeRetention(input, kind, extra), (err) => err instanceof RetentionError && pattern.test(err.message));
  assert.deepEqual(sanitizeRetention({ amount: 5 }, 'files'), { criterion: 'used', amount: 5, unit: 'years', maxDeletions: 1000, deleteMode: 'permanent', patterns: [] });
  assert.deepEqual(sanitizeRetention({ amount: '2', unit: 'months', includeJunk: false }, 'mail'), {
    criterion: 'received',
    amount: 2,
    unit: 'months',
    maxDeletions: 1000,
    deleteMode: 'permanent',
    includeTrash: true,
    includeJunk: false,
  });
  fails({ amount: 0 }, 'files', /de 1 a 100 anos/);
  fails({ amount: 5, unit: 'weeks' }, 'files', /unidade/);
  fails({ amount: 5, criterion: 'opened' }, 'files', /critério/);
  fails({ amount: 5, criterion: 'modified' }, 'mail', /critério/);
  fails({ amount: 5, criterion: 'accessed' }, 'files', /último acesso/, { cloud: true });
  fails({ amount: 5, patterns: 'Pasta\\*.tmp' }, 'files', /sem pastas/);
  fails({ amount: 5, maxDeletions: -1 }, 'files', /limite/);
  assert.deepEqual(sanitizeRetention({ amount: 1, patterns: '*.tmp\n\n*.bak\n*.tmp' }, 'files').patterns, ['*.tmp', '*.bak']);

  const now = new Date(2026, 2, 31, 10, 0); // 31/03/2026
  assert.equal(+cutoffDate({ amount: 5, unit: 'years' }, now), +new Date(2021, 2, 31, 10, 0));
  assert.equal(+cutoffDate({ amount: 1, unit: 'months' }, now), +new Date(2026, 1, 28, 10, 0), '31/03 menos 1 mês: último dia de fevereiro');
  assert.equal(+cutoffDate({ amount: 30, unit: 'days' }, now), +new Date(2026, 2, 1, 10, 0));

  const [t1, t3, t5] = [Date.UTC(2010, 0, 1), Date.UTC(2012, 0, 1), Date.UTC(2015, 0, 1)];
  const st = { mtimeMs: t1, atimeMs: t5, birthtimeMs: t3 };
  assert.equal(fileDate(st, 'modified'), t1);
  assert.equal(fileDate(st, 'accessed'), t5);
  assert.equal(fileDate(st, 'created'), t3);
  assert.equal(fileDate(st, 'used'), t5, 'sem uso: a data mais recente');
  assert.equal(fileDate({ mtimeMs: t1, atimeMs: 0, birthtimeMs: 0 }, 'created'), null, 'sem data de criação: desconhecida');
  // Datas zeradas ou padrão de sistemas antigos (01/01/1970, 01/01/1980) não contam como idade real.
  assert.equal(fileDate({ mtimeMs: Date.UTC(1980, 0, 1), atimeMs: 0, birthtimeMs: 0 }, 'modified'), null);
  assert.equal(fileDate({ mtimeMs: Date.UTC(1980, 0, 1), atimeMs: t5, birthtimeMs: 0 }, 'used'), t5);
  // Anos pelo calendário: 29/02 menos 1 ano é 28/02 (como 12 meses), e não 01/03.
  const leap = new Date(2028, 1, 29, 12, 0);
  assert.equal(+cutoffDate({ amount: 1, unit: 'years' }, leap), +new Date(2027, 1, 28, 12, 0));
  assert.equal(+cutoffDate({ amount: 1, unit: 'years' }, leap), +cutoffDate({ amount: 12, unit: 'months' }, leap));
  // Padrões de nomes: comparação linear (um padrão com vários * não fica lento).
  const started = Date.now();
  assert.equal(patternMatcher(['*a*a*a*a*a*.tmp'])(`${'a'.repeat(5000)}.txt`), false);
  assert.ok(Date.now() - started < 1000);
  assert.equal(patternMatcher(['relat*rio??.pdf'])('RELATÓRIO01.PDF'), true);
  const item = { lastModifiedDateTime: '2020-01-01T00:00:00Z', createdDateTime: '2022-01-01T00:00:00Z' };
  assert.equal(cloudDate(item, 'modified'), Date.parse('2020-01-01T00:00:00Z'));
  assert.equal(cloudDate(item, 'used'), Date.parse('2022-01-01T00:00:00Z'));
  assert.equal(cloudDate({}, 'used'), null);

  assert.equal(describeRetention({ criterion: 'used', amount: 5, unit: 'years', patterns: [] }, 'files'), 'Arquivos sem uso há mais de 5 anos');
  assert.equal(describeRetention({ criterion: 'created', amount: 1, unit: 'months', patterns: ['*.tmp'] }, 'files'), 'Arquivos criados há mais de 1 mês (somente *.tmp)');
  assert.equal(describeRetention({ criterion: 'received', amount: 2, unit: 'years', includeTrash: true, includeJunk: false }, 'mail'), 'Mensagens recebidas há mais de 2 anos (inclusive a Lixeira; sem o Lixo Eletrônico)');
  assert.equal(amountText({ amount: 1, unit: 'years' }), '1 ano');
  assert.equal(patternMatcher(['*.TMP', 'rascunho?.docx'])('arquivo.tmp'), true);
  assert.equal(patternMatcher(['*.tmp'])('arquivo.txt'), false);
  assert.equal(patternMatcher([])('qualquer'), true);
  assert.equal(ageBucket(400).label, '1 a 2 anos');
  assert.equal(ageBucket(5000).label, 'Mais de 10 anos');
});

test('relatório: faixas de idade, grupos com o tamanho e os mais antigos primeiro', () => {
  const file = (id, days, extra) => ({ id, size: 100 * id, retention: { criterion: 'modified', date: new Date(Date.now() - days * DAY).toISOString(), ageDays: days }, terms: [], matches: [], ...extra });
  const files = [
    file(1, 400, { extension: '.pdf', lastUser: 'ana', repositoryId: 'r1', repositoryName: 'RH' }),
    file(2, 5000, { extension: '.pdf', lastUser: null, repositoryId: 'r1', repositoryName: 'RH' }),
    file(3, 800, { extension: '', lastUser: 'ana', repositoryId: 'r2', repositoryName: 'Fin' }),
  ];
  const s = summarizeRetention(files);
  assert.deepEqual(s.byAge.map((b) => [b.key, b.count, b.bytes]), [['1-2-anos', 1, 100], ['2-5-anos', 1, 300], ['mais-de-10-anos', 1, 200]]);
  assert.equal(s.count, 3);
  assert.equal(s.bytes, 600);
  // Pelo espaço e, no empate, pela quantidade.
  assert.deepEqual(s.byExtension.map((g) => [g.key, g.count, g.bytes]), [['.pdf', 2, 300], ['', 1, 300]]);
  assert.deepEqual(s.byUser.map((g) => [g.key, g.identified, g.count]), [['ana', true, 2], ['', false, 1]]);
  assert.deepEqual(s.byRepository.map((g) => [g.key, g.name, g.count, g.bytes]), [['r1', 'RH', 2, 300], ['r2', 'Fin', 1, 300]]);
  // Os mais antigos primeiro; filtro pela faixa de idade.
  assert.deepEqual(filterRecords(files, { sort: 'oldest' }).map((r) => r.id), [2, 3, 1]);
  assert.deepEqual(filterRecords(files, { age: '2-5-anos' }).map((r) => r.id), [3]);
  const mails = [
    { id: 1, mailbox: 'a@x.com', mailboxName: 'A', folder: 'Caixa de Entrada', size: 10, date: '2015-01-01T00:00:00Z', retention: { criterion: 'received', date: '2015-01-01T00:00:00Z', ageDays: 4285 }, terms: [], matches: [] },
    { id: 2, mailbox: 'b@x.com', mailboxName: '', folder: 'Itens Enviados', size: 20, date: '2012-01-01T00:00:00Z', retention: { criterion: 'received', date: '2012-01-01T00:00:00Z', ageDays: 5381 }, terms: [], matches: [] },
  ];
  const m = summarizeRetention(mails, 'mail');
  assert.deepEqual(m.byMailbox.map((g) => [g.key, g.name, g.count]), [['a@x.com', 'A', 1], ['b@x.com', '', 1]]);
  assert.deepEqual(m.byFolder.map((g) => g.key).sort(), ['Caixa de Entrada', 'Itens Enviados']);
  assert.equal(m.byExtension, undefined);
  assert.deepEqual(filterMailRecords(mails, { sort: 'oldest' }).map((r) => r.id), [2, 1]);
});

/** Pasta com arquivos antigos e recentes (a data de modificação e de acesso é ajustável). */
function tree() {
  const dir = fs.mkdtempSync(path.join(root, 'repo-'));
  const put = (rel, date = null, atime = date) => {
    const target = path.join(dir, rel);
    fs.mkdirSync(path.dirname(target), { recursive: true });
    fs.writeFileSync(target, `conteúdo de ${rel}`);
    if (date) fs.utimesSync(target, atime, date);
    return target;
  };
  put('RH/antigo.docx', OLD);
  put('RH/antigo2.tmp', OLD);
  put('Temp/velho.tmp', OLD);
  put('Temp/aberto-ontem.tmp', OLD, new Date(Date.now() - DAY)); // modificado em 2015, aberto ontem
  put('recente.txt');
  return dir;
}

async function runFiles(dir, retention, { deleteMatches = false, allowDelete = true, keep = [] } = {}) {
  const messages = [];
  const policy = { ...sanitizeRetention(retention, 'files'), cutoff: cutoffDate(sanitizeRetention(retention, 'files')).toISOString() };
  const scanner = new Scanner(
    { repositories: [{ id: 'r1', name: 'Arquivos', path: dir, exclude: [], allowDelete, keep }], terms: [], options: { deleteMatches, resolveOwner: false, checkContent: false }, retention: policy, startedBy: 'política "Limpeza"' },
    (m) => messages.push(m),
  );
  const stats = await scanner.run();
  const records = messages.filter((m) => m.type === 'results').flatMap((m) => m.records);
  const events = messages.filter((m) => m.type === 'deletions').flatMap((m) => m.items);
  const logs = messages.filter((m) => m.type === 'log').map((m) => m.message);
  return { stats, records, events, logs, names: records.map((r) => r.name).sort() };
}

test('arquivos: expirados pelo critério, padrões de nomes, simulação e exclusão com limite', async () => {
  const dir = tree();
  // Simulação: lista, não exclui. Última modificação: os quatro de 2015.
  let run = await runFiles(dir, { criterion: 'modified', amount: 5 });
  assert.deepEqual(run.names, ['aberto-ontem.tmp', 'antigo.docx', 'antigo2.tmp', 'velho.tmp']);
  assert.equal(run.stats.filesSeen, 5);
  assert.equal(run.stats.filesMatched, 4);
  assert.ok(run.stats.bytesExpired > 0);
  assert.equal(run.events.length, 0);
  const rec = run.records.find((r) => r.name === 'antigo.docx');
  assert.equal(rec.retention.criterion, 'modified');
  assert.equal(rec.retention.date, OLD.toISOString());
  assert.ok(rec.retention.ageDays > 365 * 11);
  assert.deepEqual(rec.terms, []);
  assert.equal(rec.contentStatus, 'not-requested');
  assert.match(run.logs[0], /Retenção iniciada/);

  // Último acesso: o aberto ontem não expira.
  run = await runFiles(dir, { criterion: 'accessed', amount: 5 });
  assert.deepEqual(run.names, ['antigo.docx', 'antigo2.tmp', 'velho.tmp']);
  // Sem uso: a criação (agora) conta, então nada expira nos arquivos criados pelo teste.
  run = await runFiles(dir, { criterion: 'used', amount: 5 });
  assert.equal(run.stats.filesMatched, 0);
  // Somente *.tmp.
  run = await runFiles(dir, { criterion: 'modified', amount: 5, patterns: ['*.tmp'] });
  assert.deepEqual(run.names, ['aberto-ontem.tmp', 'antigo2.tmp', 'velho.tmp']);

  // Exclusão com limite de 2 por execução: os demais ficam só listados.
  run = await runFiles(dir, { criterion: 'accessed', amount: 5, maxDeletions: 2 }, { deleteMatches: true });
  assert.equal(run.stats.deleted, 2);
  assert.equal(run.stats.deleteSkipped, 1);
  assert.equal(run.events.length, 2);
  assert.ok(run.events.every((e) => e.mode === 'retention' && e.status === 'deleted' && e.method === 'file'));
  assert.equal(run.events[0].by, 'política "Limpeza"');
  assert.ok(run.logs.some((l) => /Limite de 2 exclusões desta execução atingido/.test(l)));
  const left = ['RH/antigo.docx', 'RH/antigo2.tmp', 'Temp/velho.tmp'].filter((f) => fs.existsSync(path.join(dir, f)));
  assert.equal(left.length, 1);
  assert.ok(fs.existsSync(path.join(dir, 'Temp/aberto-ontem.tmp')), 'aberto ontem: mantido');
  assert.ok(fs.existsSync(path.join(dir, 'recente.txt')));

  // Repositório protegido dentro do analisado e exclusão não permitida: nada é excluído.
  const dir2 = tree();
  // Os protegidos não são tentados (nem contam no limite: com limite 2, os dois de Temp são excluídos).
  run = await runFiles(dir2, { criterion: 'modified', amount: 5, maxDeletions: 2 }, { deleteMatches: true, keep: [{ path: path.join(dir2, 'RH'), error: 'Protegido.' }] });
  assert.ok(fs.existsSync(path.join(dir2, 'RH/antigo.docx')));
  assert.equal(run.stats.deleteProtected, 2);
  assert.equal(run.events.filter((e) => e.status === 'failed').length, 0);
  assert.equal(run.stats.deleted, 2);
  assert.ok(!fs.existsSync(path.join(dir2, 'Temp/velho.tmp')));
  assert.ok(run.logs.some((l) => /2 arquivo\(s\) expirado\(s\) em locais protegidos/.test(l)));
  run = await runFiles(tree(), { criterion: 'modified', amount: 5 }, { deleteMatches: true, allowDelete: false });
  assert.equal(run.stats.deleted, 0);
});

test('OneDrive e SharePoint: pela data do Microsoft 365, sem baixar o conteúdo, exclusão definitiva', async () => {
  await withGraph(async (data, endpoints) => {
    const drive = data.drives['d-fin'];
    drive.items = [
      folder('f-2015', '2015', [file('i-velho', 'balanco.xlsx', 'x', { modified: '2015-01-10T10:00:00Z', created: '2015-01-10T10:00:00Z' })]),
      file('i-novo', 'novo.txt', 'y', { modified: '2026-09-10T10:00:00Z', created: '2015-01-10T10:00:00Z' }),
    ];
    const retention = sanitizeRetention({ criterion: 'used', amount: 5, maxDeletions: 0 }, 'files', { cloud: true });
    const repository = cloudRepo('sharepoint', {
      allowDelete: true,
      deleteMode: 'permanent',
      exclude: [],
      cloud: { scope: 'list', accounts: [], sites: ['https://contoso.sharepoint.com/sites/Financeiro'], exclude: [] },
      keep: { all: null, accounts: [], sites: [] },
    });
    const messages = [];
    const stats = await new Scanner(
      { repositories: [repository], terms: [], options: { deleteMatches: true, checkContent: false }, retention: { ...retention, cutoff: cutoffDate(retention).toISOString() }, endpoints },
      (m) => messages.push(m),
    ).run();
    const records = messages.filter((m) => m.type === 'results').flatMap((m) => m.records);
    assert.deepEqual(records.map((r) => r.name), ['balanco.xlsx'], 'novo.txt foi modificado agora (sem uso: a data mais recente)');
    assert.equal(records[0].cloud.itemId, 'i-velho');
    assert.equal(records[0].retention.date, '2015-01-10T10:00:00.000Z');
    assert.equal(stats.deleted, 1);
    assert.equal(data.downloads, undefined, 'nenhum conteúdo baixado');
    const events = messages.filter((m) => m.type === 'deletions').flatMap((m) => m.items);
    assert.equal(events[0].method, 'permanent');
  });
});

/** Mensagem MIME simples. */
const mime = (subject, from = 'Ana Souza <ana@contoso.com>') =>
  Buffer.from([`From: ${from}`, 'To: rh@contoso.com', `Subject: ${subject}`, 'Date: Tue, 10 Mar 2015 10:00:00 -0300', `Message-ID: <${crypto.randomUUID()}@contoso.com>`, '', 'corpo', ''].join('\r\n'));

/** Texto gravado por uma exportação. */
async function collect(write) {
  const chunks = [];
  const out = new PassThrough();
  out.on('data', (c) => chunks.push(c));
  await write(out);
  out.end();
  return Buffer.concat(chunks).toString('utf8');
}

async function runMail(sources, retention, endpoints, { deleteMatches = false } = {}) {
  const policy = sanitizeRetention(retention, 'mail');
  const messages = [];
  const stats = await new MailScanner(
    { sources, terms: [], options: { deleteMatches, includeTrash: policy.includeTrash, includeJunk: policy.includeJunk }, retention: { ...policy, cutoff: cutoffDate(policy).toISOString() }, endpoints, startedBy: 'política "E-mails antigos"' },
    (m) => messages.push(m),
  ).run();
  return {
    stats,
    records: messages.filter((m) => m.type === 'results').flatMap((m) => m.records),
    events: messages.filter((m) => m.type === 'deletions').flatMap((m) => m.items),
    logs: messages.filter((m) => m.type === 'log').map((m) => m.message),
  };
}

test('e-mail: Microsoft 365, Gmail e IMAP — só os cabeçalhos das mensagens antigas', async () => {
  const graph = {
    tenant: 'contoso.onmicrosoft.com',
    clientId: '11111111-2222-3333-4444-555555555555',
    secret: 's',
    users: [
      {
        id: 'u-ana',
        mail: 'ana@contoso.com',
        displayName: 'Ana Souza',
        folders: [
          { id: 'inbox', displayName: 'Caixa de Entrada', wellKnown: 'inbox' },
          { id: 'trash', displayName: 'Itens Excluídos', wellKnown: 'deleteditems' },
          { id: 'junk', displayName: 'Lixo Eletrônico', wellKnown: 'junkemail' },
        ],
        messages: {
          inbox: [
            { id: 'm-velha', received: '2015-03-10T13:00:00Z', raw: mime('Balanço 2014') },
            { id: 'm-nova', received: '2026-09-20T10:00:00Z', raw: mime('Reunião') },
          ],
          trash: [{ id: 'm-lixeira', received: '2016-01-01T10:00:00Z', raw: mime('Apagada há tempo') }],
          junk: [{ id: 'm-spam', received: '2016-01-01T10:00:00Z', raw: mime('Promoção') }],
        },
      },
    ],
  };
  const { privateKey, publicKey } = crypto.generateKeyPairSync('rsa', { modulusLength: 2048 });
  const google = {
    publicKey,
    admin: 'admin@empresa.com',
    users: [
      {
        mail: 'caio@empresa.com',
        name: 'Caio',
        labels: [{ id: 'INBOX', name: 'INBOX', type: 'system' }],
        messages: [
          { id: 'g-velha', labelIds: ['INBOX'], internalDate: Date.parse('2015-05-01'), raw: mime('Nota fiscal 2015', 'Loja <vendas@loja.com>') },
          { id: 'g-nova', labelIds: ['INBOX'], internalDate: Date.parse('2026-09-01'), raw: mime('Hoje') },
        ],
      },
    ],
  };
  const mocks = await startMockApis({ graph, google });
  const imap = await startFakeImap({
    'carla@empresa.com': {
      password: 'p',
      folders: {
        INBOX: [
          { raw: mime('Contrato antigo', 'Bruno <bruno@empresa.com>'), date: new Date('2014-06-01T10:00:00Z') },
          { raw: mime('Recente'), date: new Date() },
        ],
      },
    },
  });
  try {
    const m365 = {
      id: 's-m365',
      name: 'Microsoft 365',
      type: 'graph',
      scope: 'list',
      mailboxes: [{ address: 'ana@contoso.com' }],
      excludeMailboxes: [],
      excludeFolders: [],
      graph: { tenantId: graph.tenant, clientId: graph.clientId },
      secrets: { clientSecret: 's' },
      allowDelete: true,
      deleteMode: 'permanent',
    };
    // Simulação, sem o Lixo Eletrônico: a velha e a da lixeira.
    let run = await runMail([m365], { amount: 5, includeJunk: false }, mocks.endpoints);
    assert.deepEqual(run.records.map((r) => r.subject).sort(), ['Apagada há tempo', 'Balanço 2014']);
    const velha = run.records.find((r) => r.subject === 'Balanço 2014');
    assert.equal(velha.from, 'Ana Souza <ana@contoso.com>');
    assert.equal(velha.folder, 'Caixa de Entrada');
    assert.equal(velha.retention.criterion, 'received');
    assert.ok(velha.internetMessageId);
    assert.ok(!mocks.calls.some((c) => c.endsWith('/$value')), 'nenhuma mensagem baixada');
    // Exportações: uma linha por mensagem expirada, com a data de recebimento e a idade.
    const policy = sanitizeRetention({ amount: 5, includeJunk: false }, 'mail');
    const mailScan = {
      kind: 'mail',
      name: 'E-mails antigos',
      status: 'completed',
      retention: { ...policy, cutoff: cutoffDate(policy).toISOString() },
      options: { deleteMatches: false },
      stats: run.stats,
      summary: { sources: [{ name: 'Microsoft 365', type: 'graph', scope: 'list', mailboxCount: 1 }] },
    };
    const csvText = await collect((out) => exportRetentionCsv(run.records, out, mailScan));
    assert.ok(csvText.startsWith('\uFEFF'), 'BOM para o Excel');
    const csv = csvText.trim().split('\r\n');
    assert.equal(csv.length, 3);
    assert.match(csv[0], /^Conexão;Caixa;Pasta;Recebida em;Idade \(dias\);Faixa de idade;Remetente;Assunto/);
    assert.match(csv.find((l) => l.includes('Balanço 2014')), /ana@contoso\.com;Caixa de Entrada;10\/03\/2015 \d\d:00:00;\d{4};Mais de 10 anos;Ana Souza <ana@contoso\.com>;Balanço 2014/);
    const html = await collect((out) => exportRetentionHtml(mailScan, run.records, out));
    assert.match(html, /Mensagens recebidas há mais de 5 anos \(inclusive a Lixeira; sem o Lixo Eletrônico\)/);
    assert.match(html, /Somente listar os itens expirados \(simulação\)/);
    assert.match(html, /Mensagens expiradas \(2\)/);
    // Exclusão definitiva.
    run = await runMail([m365], { amount: 5 }, mocks.endpoints, { deleteMatches: true });
    assert.equal(run.stats.deleted, 3);
    assert.deepEqual(graph.deleted.map((d) => d.how), ['permanentDelete', 'permanentDelete', 'permanentDelete']);
    assert.ok(run.events.every((e) => e.mode === 'retention'));
    assert.deepEqual(Object.values(graph.users[0].messages).flat().map((m) => m.id), ['m-nova']);

    // Gmail: before: na busca e só os metadados.
    const gmail = {
      id: 's-gmail',
      name: 'Google',
      type: 'gmail',
      scope: 'list',
      mailboxes: [{ address: 'caio@empresa.com' }],
      excludeMailboxes: [],
      excludeFolders: [],
      gmail: { clientEmail: 'svc@projeto.iam.gserviceaccount.com', adminEmail: 'admin@empresa.com' },
      secrets: { privateKey: privateKey.export({ type: 'pkcs8', format: 'pem' }) },
      allowDelete: true,
      deleteMode: 'permanent',
    };
    run = await runMail([gmail], { amount: 5 }, mocks.endpoints, { deleteMatches: true });
    assert.deepEqual(run.records.map((r) => [r.subject, r.from]), [['Nota fiscal 2015', 'Loja <vendas@loja.com>']]);
    assert.ok(mocks.calls.some((c) => c.endsWith('/messages/g-velha')));
    assert.ok(!mocks.calls.some((c) => c.endsWith('/messages/g-nova')), 'a recente nem é consultada (before: na busca)');
    assert.deepEqual(google.users[0].messages.map((m) => m.id), ['g-nova']);

    // IMAP: data interna e envelope, exclusão definitiva conferindo o Message-ID.
    const imapSource = {
      id: 's-imap',
      name: 'IMAP',
      type: 'imap',
      scope: 'list',
      imap: { host: '127.0.0.1', port: imap.port, security: 'none' },
      mailboxes: [{ address: 'carla@empresa.com' }],
      excludeMailboxes: [],
      excludeFolders: [],
      secrets: { defaultPassword: 'p' },
      allowDelete: true,
      deleteMode: 'permanent',
    };
    run = await runMail([imapSource], { amount: 5, maxDeletions: 0 }, {}, { deleteMatches: true });
    assert.deepEqual(run.records.map((r) => [r.subject, r.from]), [['Contrato antigo', 'Bruno <bruno@empresa.com>']]);
    assert.equal(run.stats.deleted, 1, JSON.stringify(run.events));
  } finally {
    await imap.close();
    await mocks.close();
  }
});

// ------------------------------------------------------------------------------------------------
// API das políticas (agendamentos com purpose "retention") e execução pelo gerenciador.

test('API das políticas de retenção', async () => {
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
    const dir = tree();
    const repoRes = await api('POST', '/api/repositories', { name: 'Arquivos', path: dir });
    const repoId = repoRes.data.id;
    const body = {
      purpose: 'retention',
      kind: 'files',
      name: 'Limpeza anual',
      repositoryIds: [repoId],
      retention: { criterion: 'accessed', amount: 5, unit: 'years', maxDeletions: 100 },
      options: { resolveOwner: false },
      action: 'analyze',
      rule: null,
    };
    const created = await api('POST', '/api/schedules', body);
    assert.equal(created.status, 201, JSON.stringify(created.data));
    const id = created.data.id;
    assert.equal(created.data.state, 'manual');
    assert.equal(created.data.description, 'Somente quando executada manualmente');
    assert.equal(created.data.retentionText, 'Arquivos sem acesso há mais de 5 anos');
    assert.deepEqual((await api('GET', '/api/schedules')).data, [], 'a lista de agendamentos não mostra as políticas');
    assert.equal((await api('GET', '/api/schedules?purpose=retention')).data.length, 1);

    // Simulação.
    const simulated = await api('POST', `/api/schedules/${id}/run`, {});
    assert.equal(simulated.status, 201);
    let scan = await wait(simulated.data.scan.id);
    assert.equal(scan.stats.filesMatched, 3);
    assert.equal(scan.retention.criterion, 'accessed');
    assert.ok(scan.retention.cutoff);
    assert.equal(scan.options.deleteMatches, false);
    assert.ok(fs.existsSync(path.join(dir, 'RH/antigo.docx')));

    // Exclusão: exige "Permitir exclusão" e EXCLUIR; o critério "último acesso" não vale na nuvem.
    const withDelete = { ...body, action: 'delete', confirmDelete: 'EXCLUIR' };
    assert.match((await api('PUT', `/api/schedules/${id}`, withDelete)).data.error, /A exclusão não está permitida em "Arquivos"/);
    await api('PUT', `/api/repositories/${repoId}`, { name: 'Arquivos', path: dir, allowDelete: true });
    assert.match((await api('PUT', `/api/schedules/${id}`, { ...withDelete, confirmDelete: '' })).data.error, /digite EXCLUIR/);
    const saved = await api('PUT', `/api/schedules/${id}`, withDelete);
    assert.equal(saved.status, 200);
    assert.equal(saved.data.action, 'delete');
    assert.equal((await api('POST', `/api/schedules/${id}/run`, {})).status, 400, 'executar com exclusão pede confirmação');
    const sim2 = await api('POST', `/api/schedules/${id}/run`, { simulate: true });
    scan = await wait(sim2.data.scan.id);
    assert.equal(scan.options.deleteMatches, false, '"Simular agora" não exclui');
    assert.match(scan.name, /\(simulação\)/);
    const real = await api('POST', `/api/schedules/${id}/run`, { confirm: true });
    scan = await wait(real.data.scan.id);
    assert.equal(scan.stats.deleted, 3);
    assert.match(scan.startedBy, /^política de retenção "Limpeza anual", executada agora por acesso local \(exclusão confirmada por acesso local em /);
    assert.ok(!fs.existsSync(path.join(dir, 'RH/antigo.docx')));
    assert.ok(fs.existsSync(path.join(dir, 'Temp/aberto-ontem.tmp')));
    const deletions = fs.readFileSync(path.join(store.dataDir, 'exclusoes.ndjson'), 'utf8').trim().split('\n').map((l) => JSON.parse(l));
    assert.ok(deletions.every((d) => d.mode === 'retention'));

    // Relatório: faixas de idade, filtro por idade e exportações com uma linha por item expirado.
    const summary = (await api('GET', `/api/scans/${scan.id}/summary`)).data;
    assert.deepEqual(summary.retention.byAge.map((b) => [b.key, b.count]), [['mais-de-10-anos', 3]]);
    assert.equal((await api('GET', `/api/scans/${scan.id}/results?age=mais-de-10-anos`)).data.total, 3);
    assert.equal((await api('GET', `/api/scans/${scan.id}/results?age=ate-1-ano`)).data.total, 0);
    const oldest = (await api('GET', `/api/scans/${scan.id}/results?sort=oldest`)).data.items;
    assert.ok(oldest.every((r) => r.retention.criterion === 'accessed' && r.retention.ageDays > 3650));
    // Sem ordenação informada: os mais antigos primeiro.
    const byDefault = (await api('GET', `/api/scans/${scan.id}/results`)).data.items;
    assert.deepEqual(byDefault.map((r) => r.id), oldest.map((r) => r.id));
    const csv = await (await fetch(`${base}/api/scans/${scan.id}/export.csv`)).text();
    const lines = csv.trim().split('\r\n');
    assert.equal(lines.length, 4, 'cabeçalho e três arquivos (sem termos)');
    assert.match(lines[0], /Data considerada;Idade \(dias\);Faixa de idade/);
    assert.match(lines[1], /Mais de 10 anos/);
    assert.match(lines[1], /Excluído em .* — política de retenção ""Limpeza anual"", executada agora por acesso local \(exclusão confirmada por/);
    const html = await (await fetch(`${base}/api/scans/${scan.id}/export.html`)).text();
    assert.match(html, /Relatório CLEAN – retenção/);
    assert.match(html, /Arquivos sem acesso há mais de 5 anos/);
    assert.match(html, /Data de corte \(expiram os anteriores\)/);
    assert.match(html, /Arquivos expirados \(3\)/);
    assert.match(html, /Excluir os itens expirados \(definitivamente\)/);
    const xlsx = await fetch(`${base}/api/scans/${scan.id}/export.xlsx`);
    assert.equal(xlsx.status, 200);
    assert.ok((await xlsx.arrayBuffer()).byteLength > 1000);

    // Agendada (regra de recorrência) e o repositório em uso não pode ser excluído.
    const scheduled = await api('PUT', `/api/schedules/${id}`, { ...withDelete, rule: { frequency: 'weekly', startDate: '2026-09-01', time: '03:00', weekdays: [0] } });
    assert.equal(scheduled.data.state, 'active');
    assert.equal((await api('DELETE', `/api/repositories/${repoId}`)).status, 409);
    // Não vira agendamento de análise por termos.
    assert.match((await api('PUT', `/api/schedules/${id}`, { ...withDelete, purpose: 'terms' })).data.error, /transformar/);
    assert.equal((await api('DELETE', `/api/schedules/${id}`)).status, 204);
  } finally {
    server.close();
    await scheduler.stop();
  }
});
