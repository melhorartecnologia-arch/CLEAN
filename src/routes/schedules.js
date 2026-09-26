// Rotas /api/schedules: análises agendadas com regras de recorrência.
import { Router } from 'express';
import { HttpError, bad, text } from './validate.js';
import { actor } from './scans.js';
import { ScanError, sanitizeOptions, sanitizeMailOptions } from '../scan/manager.js';
import { validateRule, RuleError, describeRule, upcoming, lastOccurrence, nextOccurrence } from '../schedule/recurrence.js';
import { targetOf, deletionPins, deletionCriteria, scheduleProblems, remainingOf } from '../schedule/scheduler.js';

const ids = (value) => (Array.isArray(value) ? [...new Set(value.filter((v) => typeof v === 'string'))] : []);
const KEEP_MAX = 500;
const NEVER = 'Pela regra informada, o agendamento nunca seria executado (a data e a hora já passaram). Confira a data de início, o horário e o término.';

function parseRule(input) {
  try {
    return validateRule(input);
  } catch (err) {
    if (err instanceof RuleError) throw bad(err.message);
    throw err;
  }
}

/** Período analisado em cada execução: todos, os últimos N dias ou desde a execução anterior. */
function parsePeriod(input) {
  const p = input && typeof input === 'object' ? input : {};
  if (p.type === 'days') {
    const days = Number(p.days);
    if (!Number.isInteger(days) || days < 1 || days > 3650) throw bad('O período deve ser de 1 a 3650 dias.');
    return { type: 'days', days };
  }
  if (p.type === 'since-last') {
    // A análise completa periódica é obrigatória: ela pega o que as incrementais não veem (itens
    // com erro de leitura, pastas movidas inteiras, mensagens movidas entre pastas).
    const fullEvery = p.fullEvery === undefined || p.fullEvery === null || p.fullEvery === '' ? 7 : Number(p.fullEvery);
    if (!Number.isInteger(fullEvery) || fullEvery < 2 || fullEvery > 50) throw bad('A análise completa periódica deve ser a cada 2 a 50 execuções.');
    return { type: 'since-last', fullEvery };
  }
  return { type: 'all' };
}

/**
 * Valida o agendamento recebido da interface. A exclusão automática exige, a cada vez que o
 * agendamento é salvo, "Permitir exclusão" em todos os locais e a confirmação digitada (EXCLUIR);
 * o alcance e a forma de exclusão de cada local ficam registrados junto com quem confirmou.
 */
function parseSchedule(body = {}, { store, existing = null, by }) {
  const kind = body.kind === 'mail' ? 'mail' : 'files';
  if (existing && existing.kind !== kind) throw bad('Não é possível trocar o tipo do agendamento (arquivos ou e-mail).');
  const name = text(body.name, 'o nome do agendamento', { required: true, max: 120 });
  const targetIds = ids(kind === 'mail' ? body.sourceIds : body.repositoryIds);
  const targets = targetIds.map((id) => targetOf(store, kind, id));
  if (targets.length === 0 || targets.some((t) => !t)) throw bad(kind === 'mail' ? 'Selecione conexões de e-mail válidas.' : 'Selecione repositórios válidos.');
  const listIds = ids(body.listIds);
  const lists = listIds.map((id) => store.getList(id));
  if (lists.length === 0 || lists.some((l) => !l)) throw bad('Selecione listas de referência válidas.');
  if (lists.every((l) => !(l.terms || []).length)) throw bad('As listas selecionadas não possuem termos.');
  let options;
  try {
    // O período de cada execução substitui a data fixa ("a partir de").
    options = kind === 'mail' ? sanitizeMailOptions({ ...body.options, receivedAfter: null }) : sanitizeOptions({ ...body.options, modifiedAfter: null });
  } catch (err) {
    if (err instanceof ScanError) throw bad(err.message);
    throw err;
  }
  const deleting = options.deleteMatches === true;
  options.deleteMatches = false; // decidido em cada execução pela ação do agendamento
  const keepLast = body.keepLast === undefined || body.keepLast === null || body.keepLast === '' ? 0 : Number(body.keepLast);
  if (!Number.isInteger(keepLast) || keepLast < 0 || keepLast > KEEP_MAX) throw bad(`Informe quantos relatórios guardar: de 1 a ${KEEP_MAX} (ou 0, para todos).`);
  const data = {
    name,
    kind,
    targetIds,
    listIds,
    options,
    action: deleting ? 'delete' : 'analyze',
    rule: parseRule(body.rule),
    period: parsePeriod(body.period),
    catchUp: body.catchUp !== false,
    keepLast,
    // Nomes no momento em que foi salvo (para explicar a falha se algo for excluído do cadastro).
    names: Object.fromEntries([...targets, ...lists].map((x) => [x.id, x.name])),
    deleteConfirmation: null,
  };
  if (deleting) {
    const blocked = targets.filter((t) => !t.allowDelete).map((t) => `"${t.name}"`);
    const where = kind === 'mail' ? 'da conexão de e-mail' : 'do repositório';
    if (blocked.length) throw bad(`A exclusão não está permitida em ${blocked.join(', ')}. Ative "Permitir exclusão" no cadastro ${where} ou escolha "Somente analisar".`);
    if (String(body.confirmDelete || '').trim().toUpperCase() !== 'EXCLUIR') throw bad('Para agendar a análise com exclusão automática, digite EXCLUIR na confirmação.');
    // Registra o alcance e a forma de exclusão de cada local e os critérios (termos, exclusões e
    // locais protegidos) no momento da confirmação: se algo disso mudar, a exclusão fica suspensa.
    data.deleteConfirmation = { by, at: new Date().toISOString(), targets: deletionPins(kind, targets), criteria: deletionCriteria(store, data) };
  }
  return data;
}

export function schedulesRouter({ store, manager, scheduler }) {
  const router = Router();
  const now = () => scheduler.now();

  const find = (id) => {
    const schedule = store.getSchedule(id);
    if (!schedule) throw new HttpError(404, 'Agendamento não encontrado.');
    return schedule;
  };

  /** Situação da última execução (com a situação atual da análise, se o relatório ainda existe). */
  const withScan = (entry) => {
    if (!entry) return null;
    const scan = entry.scanId ? store.getScan(entry.scanId) : null;
    return { ...entry, scanStatus: scan?.status || null, reportExists: Boolean(scan) };
  };

  /** Dados enviados à interface (sem o alcance registrado na confirmação da exclusão). */
  const view = (schedule, { history = false } = {}) => {
    const { deleteConfirmation, names, history: all = [], ...rest } = schedule;
    const named = (id, item) => ({ id, name: item?.name || names?.[id] || id, missing: !item });
    const remaining = remainingOf(schedule);
    let description = '';
    let end = null;
    try {
      description = describeRule(schedule.rule);
      end = schedule.enabled && schedule.nextRunAt ? lastOccurrence(schedule.rule, { after: now(), remaining }) : null;
    } catch {
      // regra inválida (dados corrompidos): aparece em "problems"
    }
    return {
      ...rest,
      targets: (schedule.targetIds || []).map((id) => named(id, targetOf(store, schedule.kind, id))),
      lists: (schedule.listIds || []).map((id) => named(id, store.getList(id))),
      deleteConfirmation: deleteConfirmation ? { by: deleteConfirmation.by, at: deleteConfirmation.at } : null,
      description,
      remaining: Number.isFinite(remaining) ? remaining : null,
      endsAt: end ? end.toISOString() : null,
      state: !schedule.enabled ? 'paused' : schedule.nextRunAt ? 'active' : 'finished',
      running: store.listScans().some((s) => s.scheduleId === schedule.id && manager.isActive(s.id)),
      lastRun: withScan(all[0]),
      problems: scheduleProblems(store, schedule),
      ...(history ? { history: all.map(withScan) } : {}),
    };
  };
  const sameRule = (a, b) => JSON.stringify(a) === JSON.stringify(b);

  const translate = (err) => {
    if (err instanceof ScanError) return new HttpError(err.status || 400, err.message);
    return err;
  };

  router.get('/', (req, res) => {
    scheduler.collectOutcomes();
    const kind = req.query.kind === 'mail' || req.query.kind === 'files' ? req.query.kind : '';
    const list = store
      .listSchedules()
      .filter((s) => !kind || s.kind === kind)
      .map((s) => view(s))
      .sort((a, b) => a.name.localeCompare(b.name, 'pt-BR'));
    res.json(list);
  });

  // Descrição da regra e as próximas execuções (prévia no formulário, antes de salvar). Na edição
  // (scheduleId) com a mesma regra, "depois de N execuções" desconta as já feitas.
  router.post('/preview', (req, res) => {
    const rule = parseRule(req.body?.rule);
    const from = now();
    const existing = typeof req.body?.scheduleId === 'string' ? store.getSchedule(req.body.scheduleId) : null;
    const remaining = remainingOf({ rule, countDone: existing && sameRule(existing.rule, rule) ? existing.countDone : 0 });
    const next = upcoming(rule, from, 5, { remaining }).map((d) => d.toISOString());
    const end = lastOccurrence(rule, { after: from, remaining });
    res.json({ rule, description: describeRule(rule), next, endsAt: end ? end.toISOString() : null, remaining: Number.isFinite(remaining) ? remaining : null });
  });

  router.post('/', (req, res) => {
    const by = actor(req);
    const data = parseSchedule(req.body || {}, { store, by });
    if (!nextOccurrence(data.rule, now(), { remaining: remainingOf({ rule: data.rule }) })) throw bad(NEVER);
    const schedule = store.createSchedule({ ...data, enabled: req.body?.enabled !== false, createdBy: by, updatedBy: by, runCount: 0, countDone: 0, history: [], nextRunAt: null });
    store.updateScheduleState(schedule.id, { nextRunAt: scheduler.plan(schedule) });
    res.status(201).json(view(schedule));
  });

  router.get('/:id', (req, res) => {
    scheduler.collectOutcomes();
    res.json(view(find(req.params.id), { history: true }));
  });

  router.put('/:id', (req, res) => {
    const existing = find(req.params.id);
    const by = actor(req);
    const data = parseSchedule(req.body || {}, { store, existing, by });
    const same = sameRule(existing.rule, data.rule);
    const countDone = same ? existing.countDone || 0 : 0; // regra nova: a contagem recomeça
    const planned = scheduler.plan({ ...existing, ...data, countDone });
    // Regra alterada que nunca mais executaria: recusada (com a mesma regra, um agendamento já
    // encerrado pode ser renomeado ou ajustado).
    if (existing.enabled && !same && !planned) throw bad(NEVER);
    // Horário que acabou de chegar e ainda não foi executado: continua valendo (o agendador o
    // executa na próxima verificação).
    const due = same && existing.enabled && existing.nextRunAt && Date.parse(existing.nextRunAt) <= +now();
    const schedule = store.updateSchedule(existing.id, { ...data, countDone, updatedBy: by });
    store.updateScheduleState(schedule.id, { nextRunAt: due ? existing.nextRunAt : planned });
    scheduler.reviewRuns(schedule.id);
    res.json(view(schedule));
  });

  router.post('/:id/pause', (req, res) => {
    const schedule = find(req.params.id);
    store.updateSchedule(schedule.id, { enabled: false, nextRunAt: null, updatedBy: actor(req) });
    // Uma execução com exclusão em andamento deixa de excluir; na fila, começa sem excluir.
    scheduler.reviewRuns(schedule.id);
    res.json(view(schedule));
  });

  router.post('/:id/resume', (req, res) => {
    const schedule = find(req.params.id);
    store.updateSchedule(schedule.id, { enabled: true, updatedBy: actor(req) });
    // Retomado: os horários que passaram durante a pausa não são executados.
    store.updateScheduleState(schedule.id, { nextRunAt: scheduler.plan(schedule) });
    res.json(view(schedule));
  });

  router.post('/:id/run', async (req, res) => {
    const schedule = find(req.params.id);
    if (schedule.action === 'delete' && req.body?.confirm !== true) throw bad('Confirme a execução: ela exclui os itens encontrados.');
    try {
      const scan = await scheduler.runNow(schedule.id, actor(req));
      res.status(201).json({ scan, schedule: view(store.getSchedule(schedule.id)) });
    } catch (err) {
      throw translate(err);
    }
  });

  router.delete('/:id', (req, res) => {
    const schedule = find(req.params.id);
    store.deleteSchedule(schedule.id);
    scheduler.reviewRuns(schedule.id);
    res.status(204).end();
  });

  return router;
}
