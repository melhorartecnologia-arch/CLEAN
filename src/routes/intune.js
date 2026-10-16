// Rotas /api/intune: reinício agendado dos dispositivos (iPhone/iPad) gerenciados no Microsoft
// Intune, pela API Microsoft Graph. Reaproveita uma conexão Microsoft 365 (Graph) já cadastrada como
// credencial (o aplicativo precisa das permissões DeviceManagementManagedDevices.Read.All e
// DeviceManagementManagedDevices.PrivilegedOperations.All).
import { Router } from 'express';
import { HttpError, bad, text } from './validate.js';
import { actor } from './scans.js';
import { validateRule, RuleError, describeRule, upcoming, lastOccurrence, nextOccurrence } from '../schedule/recurrence.js';
import { jobProblems, remainingOf } from '../intune/scheduler.js';
import { IntuneConnector } from '../intune/connector.js';
import { sanitizeScope, listDevices, deviceStale, STALE_SYNC_HOURS } from '../intune/reboot.js';
import { friendlyError } from '../scan/errors.js';

const NEVER = 'Pela regra informada, o agendamento nunca seria executado (a data e a hora já passaram). Confira a data de início, o horário e o término.';

function parseRule(input) {
  try {
    return validateRule(input);
  } catch (err) {
    if (err instanceof RuleError) throw bad(err.message);
    throw err;
  }
}

/** Valida o agendamento de reinício recebido da interface. */
function parseJob(body = {}, { store, existing = null }) {
  const name = text(body.name, 'o nome do agendamento', { required: true, max: 120 });
  const sourceId = text(body.sourceId, 'a conexão Microsoft 365', { required: true, max: 64 });
  const source = store.getMailSource(sourceId);
  if (!source) throw bad('Selecione uma conexão Microsoft 365 válida.');
  if (source.type !== 'graph') throw bad('O reinício pelo Intune precisa de uma conexão do Microsoft 365 (Graph).');
  let scope;
  try {
    scope = sanitizeScope(body.scope);
  } catch (err) {
    if (err.status === 400) throw bad(err.message);
    throw err;
  }
  return {
    name,
    sourceId,
    sourceName: source.name,
    scope,
    rule: parseRule(body.rule),
    catchUp: body.catchUp !== false,
  };
}

export function intuneRouter({ store, scheduler, endpoints = {} }) {
  const router = Router();
  const now = () => scheduler.now();

  const find = (id) => {
    const job = store.getIntuneJob(id);
    if (!job) throw new HttpError(404, 'Agendamento de reinício não encontrado.');
    return job;
  };

  /** Conexões Microsoft 365 (Graph) que podem ser usadas como credencial do Intune. */
  const graphSources = () =>
    store
      .listMailSources()
      .filter((s) => s.type === 'graph')
      .map((s) => ({ id: s.id, name: s.name, auth: s.graph?.auth || 'secret', account: s.graph?.account ? { username: s.graph.account.username, address: s.graph.account.address } : null }))
      .sort((a, b) => a.name.localeCompare(b.name, 'pt-BR'));

  /** Monta o conector do Intune a partir de uma conexão, com os segredos decifrados. */
  const connectorFor = (source, signal) => {
    let secrets;
    try {
      secrets = store.openMailSecrets(source);
    } catch (err) {
      throw new HttpError(409, err.message);
    }
    const grantId = source.graph?.account?.grantId;
    const onRefreshToken = grantId ? (token) => store.saveRefreshToken(source.id, grantId, token) : undefined;
    return new IntuneConnector({ ...source, secrets }, { endpoints, signal, onRefreshToken });
  };

  const sourceFor = (id) => {
    const source = store.getMailSource(id);
    if (!source) throw new HttpError(404, 'Conexão Microsoft 365 não encontrada.');
    if (source.type !== 'graph') throw new HttpError(400, 'A conexão escolhida não é do Microsoft 365 (Graph).');
    return source;
  };

  /** Dados enviados à interface (sem segredos). */
  const view = (job, { history = false } = {}) => {
    const { history: all = [], ...rest } = job;
    const source = store.getMailSource(job.sourceId);
    const remaining = remainingOf(job);
    let description = '';
    let end = null;
    try {
      description = describeRule(job.rule);
      end = job.enabled && job.nextRunAt ? lastOccurrence(job.rule, { after: now(), remaining }) : null;
    } catch {
      // regra inválida (dados corrompidos): aparece em "problems"
    }
    return {
      ...rest,
      source: { id: job.sourceId, name: source?.name || job.sourceName || job.sourceId, missing: !source, type: source?.type || null },
      description,
      remaining: Number.isFinite(remaining) ? remaining : null,
      endsAt: end ? end.toISOString() : null,
      state: !job.enabled ? 'paused' : job.nextRunAt ? 'active' : 'finished',
      problems: jobProblems(store, job),
      lastRun: all[0] || null,
      ...(history ? { history: all } : {}),
    };
  };

  const sameRule = (a, b) => JSON.stringify(a) === JSON.stringify(b);

  // Conexões Microsoft 365 disponíveis como credencial.
  router.get('/sources', (req, res) => {
    res.json(graphSources());
  });

  // Testa o acesso ao Intune com uma conexão (autentica e lê o primeiro aparelho iPhone/iPad).
  router.post('/test', async (req, res) => {
    const source = sourceFor(text(req.body?.sourceId, 'a conexão Microsoft 365', { required: true, max: 64 }));
    const connector = connectorFor(source, AbortSignal.timeout(60000));
    try {
      const result = await connector.test();
      res.json(result);
    } catch (err) {
      res.json({ ok: false, message: friendlyError(err) });
    } finally {
      await connector.close().catch(() => {});
    }
  });

  // Lista os aparelhos (iPhone/iPad) de uma conexão (prévia no formulário e conferência do escopo).
  router.get('/devices', async (req, res) => {
    const source = sourceFor(text(req.query?.sourceId, 'a conexão Microsoft 365', { required: true, max: 64 }));
    const connector = connectorFor(source, AbortSignal.timeout(120000));
    try {
      const devices = await listDevices(connector, { mode: 'all', names: [] });
      const now = Date.now();
      res.json({ devices: devices.map((d) => ({ ...d, stale: deviceStale(d.lastSync, now) })), total: devices.length, staleHours: STALE_SYNC_HOURS });
    } catch (err) {
      throw new HttpError(502, `Não foi possível listar os aparelhos: ${friendlyError(err)}`);
    } finally {
      await connector.close().catch(() => {});
    }
  });

  // Descrição da regra e as próximas execuções (prévia no formulário, antes de salvar).
  router.post('/preview', (req, res) => {
    const rule = parseRule(req.body?.rule);
    const from = now();
    const existing = typeof req.body?.jobId === 'string' ? store.getIntuneJob(req.body.jobId) : null;
    const remaining = remainingOf({ rule, countDone: existing && sameRule(existing.rule, rule) ? existing.countDone : 0 });
    const next = upcoming(rule, from, 5, { remaining }).map((d) => d.toISOString());
    const end = lastOccurrence(rule, { after: from, remaining });
    res.json({ rule, description: describeRule(rule), next, endsAt: end ? end.toISOString() : null, remaining: Number.isFinite(remaining) ? remaining : null });
  });

  router.get('/', (req, res) => {
    const list = store
      .listIntuneJobs()
      .map((j) => view(j))
      .sort((a, b) => a.name.localeCompare(b.name, 'pt-BR'));
    res.json(list);
  });

  router.post('/', (req, res) => {
    const by = actor(req);
    const data = parseJob(req.body || {}, { store });
    if (!nextOccurrence(data.rule, now(), { remaining: remainingOf({ rule: data.rule }) })) throw bad(NEVER);
    const job = store.createIntuneJob({ ...data, enabled: req.body?.enabled !== false, createdBy: by, updatedBy: by, runCount: 0, countDone: 0, history: [], nextRunAt: null });
    store.updateIntuneJobState(job.id, { nextRunAt: scheduler.plan(job) });
    res.status(201).json(view(store.getIntuneJob(job.id)));
  });

  router.get('/:id', (req, res) => {
    res.json(view(find(req.params.id), { history: true }));
  });

  router.put('/:id', (req, res) => {
    const existing = find(req.params.id);
    const by = actor(req);
    const data = parseJob(req.body || {}, { store, existing });
    const same = sameRule(existing.rule, data.rule);
    const countDone = same ? existing.countDone || 0 : 0; // regra nova: a contagem recomeça
    const planned = scheduler.plan({ ...existing, ...data, countDone });
    if (existing.enabled && !same && !planned) throw bad(NEVER);
    // Horário que acabou de chegar e ainda não foi executado: continua valendo.
    const due = same && existing.enabled && existing.nextRunAt && Date.parse(existing.nextRunAt) <= +now();
    const job = store.updateIntuneJob(existing.id, { ...data, countDone, updatedBy: by });
    store.updateIntuneJobState(job.id, { nextRunAt: due ? existing.nextRunAt : planned });
    res.json(view(store.getIntuneJob(job.id)));
  });

  router.post('/:id/pause', (req, res) => {
    const job = find(req.params.id);
    store.updateIntuneJob(job.id, { enabled: false, updatedBy: actor(req) });
    store.updateIntuneJobState(job.id, { nextRunAt: null });
    res.json(view(store.getIntuneJob(job.id)));
  });

  router.post('/:id/resume', (req, res) => {
    const job = find(req.params.id);
    store.updateIntuneJob(job.id, { enabled: true, updatedBy: actor(req) });
    // Retomado: os horários que passaram durante a pausa não são executados.
    store.updateIntuneJobState(job.id, { nextRunAt: scheduler.plan(store.getIntuneJob(job.id)) });
    res.json(view(store.getIntuneJob(job.id)));
  });

  // Executar agora. simulate: só lista os aparelhos do escopo. Reinício real pede confirm: true.
  router.post('/:id/run', async (req, res) => {
    const job = find(req.params.id);
    const simulate = req.body?.simulate === true;
    if (!simulate && req.body?.confirm !== true) throw bad('Confirme a execução: ela reinicia os aparelhos do escopo.');
    try {
      const run = await scheduler.runNow(job.id, actor(req), { simulate });
      res.status(201).json({ run, job: view(store.getIntuneJob(job.id), { history: true }) });
    } catch (err) {
      throw new HttpError(err.status && Number.isInteger(err.status) ? err.status : 500, err.message);
    }
  });

  router.delete('/:id', (req, res) => {
    find(req.params.id);
    store.deleteIntuneJob(req.params.id);
    res.status(204).end();
  });

  return router;
}
