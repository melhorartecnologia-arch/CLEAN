// Salvaguardas das políticas de retenção (achados da revisão): a data do critério é conferida de
// novo na hora de excluir, o limite da execução conta só o que foi excluído, locais protegidos não
// são tentados, mensagens já na Lixeira não são movidas de novo, datas zeradas não expiram, as
// simulações não tiram o lugar dos relatórios guardados e os textos citam a política.
import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import crypto from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { sanitizeRetention, cutoffDate } from '../src/retention/policy.js';
import { Scanner } from '../src/scan/scanner.js';
import { MailScanner } from '../src/mail/scanner.js';
import { Store } from '../src/store.js';
import { ScanManager } from '../src/scan/manager.js';
import { Scheduler, deletionPins, deletionCriteria, scheduleProblems, checkDeleteSchedules } from '../src/schedule/scheduler.js';
import { assertUnused } from '../src/routes/validate.js';
import { startMockApis } from './helpers/mock-apis.js';
import { startFakeImap } from './helpers/fake-imap.js';

let root;
const OLD = new Date('2015-03-10T12:00:00Z');

before(() => {
  root = fs.mkdtempSync(path.join(os.tmpdir(), 'clean-ret-safety-'));
});

after(() => {
  fs.rmSync(root, { recursive: true, force: true });
});

/** Pasta com os arquivos informados, todos de 2015 (modificação e último acesso). */
function oldFiles(names) {
  const dir = fs.mkdtempSync(path.join(root, 'repo-'));
  for (const name of names) {
    const target = path.join(dir, name);
    fs.mkdirSync(path.dirname(target), { recursive: true });
    fs.writeFileSync(target, name);
    fs.utimesSync(target, OLD, OLD);
  }
  return dir;
}

const policyOf = (input, kind = 'files') => {
  const r = sanitizeRetention(input, kind);
  return { ...r, cutoff: cutoffDate(r).toISOString() };
};

async function runFiles(dir, retention, { ownerResolver = null, onMessage = null } = {}) {
  const messages = [];
  let scanner;
  const emit = (m) => {
    messages.push(m);
    onMessage?.(m, scanner);
  };
  scanner = new Scanner(
    {
      repositories: [{ id: 'r1', name: 'Dados', path: dir, exclude: [], allowDelete: true, keep: [] }],
      terms: [],
      options: { deleteMatches: true, resolveOwner: Boolean(ownerResolver), checkContent: false, concurrency: 1 },
      retention,
    },
    emit,
    ownerResolver ? { ownerResolver } : {},
  );
  const stats = await scanner.run();
  return {
    stats,
    events: messages.filter((m) => m.type === 'deletions').flatMap((m) => m.items),
    logs: messages.filter((m) => m.type === 'log').map((m) => m.message),
  };
}

test('arquivo aberto depois da listagem deixa de estar expirado e não é excluído', async () => {
  const dir = oldFiles(['contrato.docx']);
  const target = path.join(dir, 'contrato.docx');
  // Enquanto os proprietários são identificados, um usuário abre o arquivo (o último acesso muda;
  // o conteúdo e a data de modificação, não).
  const ownerResolver = {
    async resolve(paths) {
      for (const p of paths) fs.utimesSync(p, new Date(), fs.statSync(p).mtime);
      return new Map();
    },
  };
  const run = await runFiles(dir, policyOf({ criterion: 'accessed', amount: 5 }), { ownerResolver });
  assert.ok(fs.existsSync(target));
  assert.equal(run.stats.deleted, 0);
  assert.equal(run.stats.deleteChanged, 1);
  assert.deepEqual(run.events.map((e) => e.status), ['changed']);
  assert.match(run.events[0].error, /deixou de estar expirado/);
});

test('limite da execução: conta só o que foi excluído; as falhas têm o mesmo limite', async () => {
  // O primeiro arquivo listado vira uma pasta antes da exclusão (falha persistente).
  const breakFirst = {
    async resolve(paths) {
      fs.rmSync(paths[0]);
      fs.mkdirSync(paths[0]);
      return new Map();
    },
  };
  let dir = oldFiles(['a.tmp', 'b.tmp', 'c.tmp']);
  let run = await runFiles(dir, policyOf({ criterion: 'modified', amount: 5, maxDeletions: 2 }), { ownerResolver: breakFirst });
  assert.equal(run.stats.deleteErrors, 1);
  assert.equal(run.stats.deleted, 2, 'a vaga da falha volta: os outros dois arquivos são excluídos');
  assert.equal(run.stats.deleteSkipped, 0);

  // Com as duas falhando, o limite das falhas (1) interrompe as tentativas.
  const breakAll = {
    async resolve(paths) {
      for (const p of paths) {
        fs.rmSync(p);
        fs.mkdirSync(p);
      }
      return new Map();
    },
  };
  dir = oldFiles(['a.tmp', 'b.tmp']);
  run = await runFiles(dir, policyOf({ criterion: 'modified', amount: 5, maxDeletions: 1 }), { ownerResolver: breakAll });
  assert.equal(run.stats.deleteErrors, 1);
  assert.equal(run.stats.deleteSkipped, 1);
  assert.ok(run.logs.some((l) => /Uma falha de exclusão nesta execução \(o limite da política\)/.test(l)));

  // Mensagem do limite no singular.
  dir = oldFiles(['a.tmp', 'b.tmp']);
  run = await runFiles(dir, policyOf({ criterion: 'modified', amount: 5, maxDeletions: 1 }));
  assert.equal(run.stats.deleted, 1);
  assert.ok(run.logs.some((l) => /Limite de 1 exclusão desta execução atingido/.test(l)));
});

test('exclusão desligada durante a execução: os demais itens não viram falhas', async () => {
  const dir = oldFiles(Array.from({ length: 20 }, (_, i) => `f${String(i).padStart(2, '0')}.tmp`));
  const run = await runFiles(dir, policyOf({ criterion: 'modified', amount: 5, maxDeletions: 10 }), {
    onMessage: (m, scanner) => {
      if (m.type === 'deletions') scanner.revokeDeletion({ all: true, reason: 'a política de retenção foi pausada' });
    },
  });
  assert.equal(run.stats.deleted, 1);
  assert.equal(run.stats.deleteErrors, 0);
  assert.deepEqual(run.events.map((e) => e.status), ['deleted']);
  assert.ok(run.logs.some((l) => /Exclusão automática desativada: a política de retenção foi pausada/.test(l)));
});

const mime = (subject) => Buffer.from([`From: Ana <ana@contoso.com>`, 'To: rh@contoso.com', `Subject: ${subject}`, `Message-ID: <${crypto.randomUUID()}@contoso.com>`, '', 'corpo', ''].join('\r\n'));

async function runMail(source, retention, endpoints = {}) {
  const records = [];
  const events = [];
  const stats = await new MailScanner(
    { sources: [source], terms: [], options: { deleteMatches: true, includeTrash: retention.includeTrash, includeJunk: retention.includeJunk }, retention, endpoints },
    (m) => {
      if (m.type === 'results') records.push(...m.records);
      if (m.type === 'deletions') events.push(...m.items);
    },
  ).run();
  return { stats, records, events };
}

test('e-mail "para a lixeira": as mensagens que já estão na Lixeira não são movidas de novo', async () => {
  const graph = {
    tenant: 'contoso.onmicrosoft.com',
    clientId: '11111111-2222-3333-4444-555555555555',
    secret: 's',
    users: [
      {
        id: 'u-ana',
        mail: 'ana@contoso.com',
        displayName: 'Ana',
        // "Deleted Items" antes de "Inbox", como no Graph em inglês.
        folders: [
          { id: 'trash', displayName: 'Deleted Items', wellKnown: 'deleteditems' },
          { id: 'inbox', displayName: 'Inbox', wellKnown: 'inbox' },
        ],
        messages: {
          trash: [],
          inbox: [
            { id: 'm-a', received: '2015-03-10T13:00:00Z', raw: mime('A') },
            { id: 'm-b', received: '2015-02-10T13:00:00Z', raw: mime('B') },
          ],
        },
      },
    ],
  };
  const mocks = await startMockApis({ graph });
  try {
    const source = {
      id: 's',
      name: 'M365',
      type: 'graph',
      scope: 'list',
      mailboxes: [{ address: 'ana@contoso.com' }],
      excludeMailboxes: [],
      excludeFolders: [],
      graph: { tenantId: graph.tenant, clientId: graph.clientId },
      secrets: { clientSecret: 's' },
      allowDelete: true,
      deleteMode: 'trash',
    };
    const retention = policyOf({ amount: 5, deleteMode: 'trash', maxDeletions: 1 }, 'mail');
    const runs = [];
    for (let i = 0; i < 3; i++) {
      graph.deleted = [];
      const run = await runMail(source, retention, mocks.endpoints);
      runs.push({ deleted: run.stats.deleted, inTrash: run.stats.alreadyInTrash, calls: graph.deleted.map((d) => d.id) });
    }
    // 1ª: move uma (limite 1); 2ª: a da Lixeira não conta e a outra é movida; 3ª: nada a mover.
    assert.equal(runs[0].deleted, 1);
    assert.deepEqual(runs[1], { deleted: 1, inTrash: 1, calls: [runs[1].calls[0]] });
    assert.notEqual(runs[1].calls[0], runs[0].calls[0]);
    assert.deepEqual(runs[2], { deleted: 0, inTrash: 2, calls: [] });
    const last = await runMail(source, retention, mocks.endpoints);
    assert.ok(last.records.every((r) => r.inTrash), 'o relatório marca as que já estão na Lixeira');
    assert.equal(last.events.length, 0);
  } finally {
    await mocks.close();
  }
});

test('e-mail: IMAP com a Lixeira e data interna zerada', async () => {
  const at = (iso) => new Date(iso);
  const accounts = {
    'carla@empresa.com': {
      password: 'p',
      folders: {
        INBOX: [
          { raw: mime('Data zerada (01/01/1970)'), date: new Date(0) },
          { raw: mime('Antiga'), date: at('2014-06-01T10:00:00Z') },
        ],
        Lixeira: [{ raw: mime('Já na Lixeira'), date: at('2013-01-01T10:00:00Z') }],
      },
      special: { Lixeira: '\\Trash' },
    },
  };
  const imap = await startFakeImap(accounts);
  try {
    const source = {
      id: 's',
      name: 'IMAP',
      type: 'imap',
      scope: 'list',
      imap: { host: '127.0.0.1', port: imap.port, security: 'none' },
      mailboxes: [{ address: 'carla@empresa.com' }],
      excludeMailboxes: [],
      excludeFolders: [],
      secrets: { defaultPassword: 'p' },
      allowDelete: true,
      deleteMode: 'trash',
    };
    const run = await runMail(source, policyOf({ amount: 5, deleteMode: 'trash' }, 'mail'));
    assert.deepEqual(run.records.map((r) => [r.subject, r.inTrash]).sort(), [
      ['Antiga', false],
      ['Já na Lixeira', true],
    ]);
    assert.equal(run.stats.retentionUnknown, 1, 'a data zerada não conta como idade');
    assert.equal(run.stats.alreadyInTrash, 1);
    assert.equal(run.stats.deleted, 1);
    assert.deepEqual(
      accounts['carla@empresa.com'].folders.Lixeira.filter((m) => !m.deleted).map((m) => m.raw.toString().match(/Subject: (.*)/)[1]).sort(),
      ['Antiga', 'Já na Lixeira'],
    );
  } finally {
    await imap.close();
  }
});

test('agendador: relatórios guardados sem contar as simulações, políticas manuais e textos', async () => {
  const dir = oldFiles(['a.tmp', 'b.tmp']);
  const store = await new Store(path.join(root, `data-${crypto.randomUUID()}`)).init();
  const manager = new ScanManager(store);
  const warns = [];
  const original = console.warn;
  console.warn = (...args) => warns.push(args.join(' '));
  const scheduler = new Scheduler({ store, manager, tickMs: 3600000 });
  try {
    const repo = store.createRepository({ name: 'Repo', path: dir, type: 'local', exclude: [], allowDelete: true });
    const data = {
      purpose: 'retention',
      name: 'Temporários',
      kind: 'files',
      targetIds: [repo.id],
      listIds: [],
      options: { resolveOwner: false },
      retention: sanitizeRetention({ criterion: 'modified', amount: 5 }, 'files'),
      action: 'delete',
      rule: null,
      period: { type: 'all' },
      catchUp: true,
      keepLast: 1,
    };
    const schedule = store.createSchedule({
      ...data,
      enabled: true,
      history: [],
      runCount: 0,
      countDone: 0,
      nextRunAt: null,
      deleteConfirmation: { by: 'ana', at: new Date().toISOString(), targets: deletionPins('files', [repo]), criteria: deletionCriteria(store, data) },
    });
    // Política manual: sem aviso de "regra inválida" ao iniciar.
    await scheduler.start();
    assert.deepEqual(warns, []);

    const wait = async (id) => {
      for (let i = 0; i < 400; i++) {
        const s = store.getScan(id);
        if (s && !['queued', 'running'].includes(s.status) && !manager.isActive(id)) return s;
        await new Promise((r) => setTimeout(r, 25));
      }
      throw new Error('não terminou');
    };
    const real = await scheduler.runNow(schedule.id, 'ana');
    assert.equal((await wait(real.id)).stats.deleted, 2);
    scheduler.collectOutcomes();
    await scheduler.pruning;
    const sim = await scheduler.runNow(schedule.id, 'ana', { simulate: true });
    assert.equal(sim.simulated, true);
    await wait(sim.id);
    scheduler.collectOutcomes();
    await scheduler.pruning;
    assert.ok(store.getScan(real.id), 'a simulação não tira o lugar do relatório da execução que excluiu');
    assert.ok(store.getScan(sim.id));

    // Textos: a política é citada como política (e não como agendamento).
    const conflict = (() => {
      try {
        assertUnused(store, 'files', repo.id, 'O repositório');
      } catch (err) {
        return err.message;
      }
      return '';
    })();
    assert.equal(conflict, 'O repositório está em uso na política de retenção "Temporários". Retire-o da política (ou exclua a política) antes de excluir.');
    const { warning } = checkDeleteSchedules(store, () => store.updateRepository(repo.id, { ...repo, allowDelete: false }), scheduler);
    assert.match(warning, /^A exclusão da política de retenção "Temporários" ficou suspensa com esta alteração: ela não é executada até ser salva e confirmada de novo em Retenção\.$/);
    assert.deepEqual(scheduleProblems(store, store.getSchedule(schedule.id)), [
      'A exclusão foi desligada no cadastro do repositório "Repo". Permita a exclusão de novo ou edite a política para "Somente listar (simulação)".',
    ]);
  } finally {
    console.warn = original;
    await scheduler.stop();
    await manager.shutdown?.({ graceMs: 100 });
    await store.close();
  }
});
