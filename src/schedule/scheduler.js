// Agendador: inicia as análises agendadas nos horários das regras de recorrência.
//
// A cada poucos segundos, os agendamentos ativos cujo horário chegou são executados (a análise
// entra na fila do gerenciador como qualquer outra). Regras:
// - Sobreposição: se uma execução do mesmo agendamento ainda estiver em andamento (ou na fila), a
//   nova é pulada e fica registrada no histórico.
// - Horário perdido (CLEAN parado ou computador desligado): na volta, o agendamento é executado
//   uma vez (se "executar assim que possível" estiver marcado) ou o horário fica registrado como
//   perdido; em ambos os casos, os demais horários perdidos não são repetidos.
// - "Depois de N execuções": contam as execuções de fato iniciadas pelo agendamento.
// - Exclusão automática: vale somente para o que foi confirmado ao salvar o agendamento — o
//   alcance e a forma de exclusão de cada local e também os critérios (termos das listas,
//   exclusões e locais protegidos). Se algo disso mudar, a execução não é iniciada (e uma execução
//   que esperava na fila começa sem excluir) até o agendamento ser salvo e confirmado de novo.
import crypto from 'node:crypto';
import { nextOccurrence, countBetween, validateRule } from './recurrence.js';
import { ScanError, sanitizeOptions, sanitizeMailOptions } from '../scan/manager.js';
import { deletionScope, mailDeletionScope, isCloudRepo, keptPaths } from '../scan/delete.js';
import { keptCloud } from '../cloud/drives.js';
import { sanitizeRetention } from '../retention/policy.js';

const TICK_MS = 15000;
// Atraso tolerado (servidor ocupado): acima disso, o horário conta como perdido.
const GRACE_MS = 5 * 60000;
export const MAX_HISTORY = 50;
// Margem da análise incremental: diferenças de relógio entre este servidor e os servidores de
// arquivos, o Microsoft 365 ou o Google.
export const INCREMENTAL_MARGIN_MS = 3600000;
const DAY_MS = 86400000;
const FINISHED = new Set(['completed', 'failed', 'cancelled', 'interrupted']);

const pad = (n) => String(n).padStart(2, '0');
const hash = (value) => crypto.createHash('sha256').update(JSON.stringify(value)).digest('hex').slice(0, 32);

/** Data e hora do servidor (dd/mm/aaaa hh:mm). */
export function fmt(value) {
  const d = new Date(value);
  return `${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${d.getFullYear()} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

const NOUNS = {
  files: { one: 'o repositório', gender: 'o', registry: 'do repositório' },
  mail: { one: 'a conexão de e-mail', gender: 'a', registry: 'da conexão de e-mail' },
};

/** Repositório (arquivos) ou conexão de e-mail (mail) do agendamento, ou null. */
export function targetOf(store, kind, id) {
  return kind === 'mail' ? store.getMailSource(id) : store.getRepository(id);
}

/** Alcance da exclusão de um repositório ou conexão (muda com o caminho, as contas, os sites ou as caixas). */
export function scopeOf(kind, target) {
  return kind === 'mail' ? mailDeletionScope(target) : deletionScope(target);
}

/** Forma de exclusão: 'file' (pastas do Windows), 'trash' ou 'permanent'. */
export function modeOf(kind, target) {
  if (kind === 'mail') return target.deleteMode === 'trash' ? 'trash' : 'permanent';
  if (!isCloudRepo(target)) return 'file';
  return target.deleteMode === 'permanent' ? 'permanent' : 'trash';
}

/** Exclusão confirmada ao salvar: alcance e forma de cada repositório ou conexão. */
export function deletionPins(kind, targets) {
  return Object.fromEntries(targets.map((t) => [t.id, { scope: scopeOf(kind, t), mode: modeOf(kind, t) }]));
}

const termKey = (t) => JSON.stringify([t.type, t.value, Boolean(t.wholeWord), t.validator || null]);

/**
 * O que decide o que é excluído, além do alcance: os termos das listas, as exclusões (pastas,
 * contas, sites, caixas e pastas de e-mail) e os locais protegidos por repositórios que não
 * permitem exclusão. Registrado ao confirmar a exclusão automática.
 */
export function deletionCriteria(store, { kind, targetIds = [], listIds = [] }) {
  const all = store.listRepositories();
  const terms = listIds.flatMap((id) => (store.getList(id)?.terms || []).map(termKey)).sort();
  const places = [...targetIds].sort().map((id) => {
    const t = targetOf(store, kind, id);
    if (!t) return [id, null];
    if (kind === 'mail') return [id, t.excludeMailboxes || [], t.excludeFolders || []];
    if (!isCloudRepo(t)) return [id, t.exclude || [], keptPaths(t, all).map((k) => k.path).sort()];
    const keep = keptCloud(t, all);
    const kept = [Boolean(keep.all), keep.accounts.map((k) => k.value).sort(), keep.sites.map((k) => k.value).sort()];
    return [id, t.exclude || [], t.cloud?.exclude || [], kept];
  });
  return hash([kind, terms, places]);
}

/** Quantas execuções ainda faltam no término "depois de N execuções" (sem esse término: infinitas). */
export function remainingOf(schedule) {
  const rule = schedule.rule || {};
  if (rule.frequency === 'once' || rule.end !== 'count') return Infinity;
  return Math.max(0, rule.count - (schedule.countDone || 0));
}

/**
 * Problemas que impedem a execução do agendamento agora (regra inválida, cadastros excluídos,
 * listas sem termos, exclusão automática que não vale mais). Lista vazia: pode executar.
 */
export function scheduleProblems(store, schedule) {
  const problems = [];
  const retention = schedule.purpose === 'retention';
  // Política de retenção sem regra: executada só manualmente.
  if (schedule.rule || !retention) {
    try {
      validateRule(schedule.rule);
    } catch (err) {
      problems.push(`A regra de recorrência salva é inválida (${err.message}) Edite o agendamento.`);
    }
  }
  const noun = NOUNS[schedule.kind] || NOUNS.files;
  const name = (id) => schedule.names?.[id] || id;
  const targets = [];
  for (const id of schedule.targetIds || []) {
    const t = targetOf(store, schedule.kind, id);
    if (t) targets.push(t);
    else problems.push(`${noun.one[0].toUpperCase()}${noun.one.slice(1)} "${name(id)}" foi excluíd${noun.gender} do cadastro. Edite o agendamento.`);
  }
  const lists = [];
  for (const id of schedule.listIds || []) {
    const l = store.getList(id);
    if (l) lists.push(l);
    else problems.push(`A lista de referência "${name(id)}" foi excluída. Edite o agendamento.`);
  }
  if (!retention && lists.length && lists.every((l) => !(l.terms || []).length)) problems.push('As listas de referência do agendamento não têm termos.');
  if (retention) {
    try {
      sanitizeRetention(schedule.retention, schedule.kind, { cloud: targets.some(isCloudRepo) });
    } catch (err) {
      problems.push(`${err.message} Edite a política.`);
    }
  }
  if (schedule.action === 'delete') {
    const confirmation = schedule.deleteConfirmation || {};
    const pins = confirmation.targets || {};
    const again = 'Edite o agendamento e confirme a exclusão de novo (ou escolha "Somente analisar").';
    for (const t of targets) {
      const pin = pins[t.id];
      if (!t.allowDelete) problems.push(`A exclusão foi desligada no cadastro ${noun.registry} "${t.name}". Permita a exclusão de novo ou edite o agendamento para "Somente analisar".`);
      else if (!pin || pin.scope !== scopeOf(schedule.kind, t)) {
        const what = schedule.kind === 'mail' ? 'a conta, o servidor ou as caixas' : 'o caminho, as contas ou os sites';
        problems.push(`O cadastro ${noun.registry} "${t.name}" mudou (${what}) depois que a exclusão automática foi confirmada neste agendamento. ${again}`);
      } else if (!retention && pin.mode !== 'permanent' && modeOf(schedule.kind, t) === 'permanent') {
        // (na retenção vale a forma de exclusão da própria política, confirmada ao salvá-la)
        problems.push(`A exclusão em "${t.name}" passou a ser definitiva depois que a exclusão automática foi confirmada neste agendamento. ${again}`);
      }
    }
    if (targets.length === (schedule.targetIds || []).length && lists.length === (schedule.listIds || []).length && confirmation.criteria !== deletionCriteria(store, schedule)) {
      const what = schedule.kind === 'mail' ? 'as caixas ou pastas ignoradas' : 'as pastas ignoradas ou os locais protegidos por repositórios sem exclusão';
      problems.push(`Os termos das listas de referência, ${what} mudaram depois que a exclusão automática foi confirmada neste agendamento. ${again}`);
    }
  }
  return problems;
}

/**
 * Aplica uma alteração num cadastro (repositório, conexão ou lista) e devolve { result, warning }:
 * o aviso cita os agendamentos com exclusão automática que ficaram suspensos por causa dela.
 */
export function checkDeleteSchedules(store, apply, scheduler = null) {
  const affected = store.listSchedules().filter((s) => s.action === 'delete');
  const fine = new Set(affected.filter((s) => scheduleProblems(store, s).length === 0).map((s) => s.id));
  const result = apply();
  // Execuções agendadas em andamento deixam de excluir se a exclusão confirmada não vale mais.
  scheduler?.reviewAll();
  const names = affected.filter((s) => fine.has(s.id) && store.getSchedule(s.id) && scheduleProblems(store, s).length > 0).map((s) => `"${s.name}"`);
  if (!names.length) return { result, warning: '' };
  const one = names.length === 1;
  return {
    result,
    warning: `A exclusão automática ${one ? 'do agendamento' : 'dos agendamentos'} ${names.join(', ')} ficou suspensa com esta alteração: ${one ? 'ele não exclui' : 'eles não excluem'} nada até ${one ? 'ser salvo e confirmado' : 'serem salvos e confirmados'} de novo em Agendamentos.`,
  };
}

/** O que a análise alcança (sem o período): muda quando a ação, os locais, os termos ou as opções mudam. */
export function coverageSignature(store, schedule) {
  const kind = schedule.kind;
  const targets = [...(schedule.targetIds || [])].sort().map((id) => {
    const t = targetOf(store, kind, id);
    if (!t) return [id, null];
    return kind === 'mail' ? [id, mailDeletionScope(t), t.excludeMailboxes || [], t.excludeFolders || []] : [id, deletionScope(t), t.exclude || [], t.cloud?.exclude || []];
  });
  const terms = (schedule.listIds || []).flatMap((id) => (store.getList(id)?.terms || []).map(termKey)).sort();
  // eslint-disable-next-line no-unused-vars
  const { concurrency, deleteMatches, modifiedAfter, receivedAfter, ...options } = schedule.options || {};
  // A ação entra: ao passar a excluir, a primeira execução é completa (exclui também o que as
  // execuções anteriores, só de análise, encontraram).
  return hash([kind, schedule.action, targets, terms, Object.entries(options).sort(([a], [b]) => a.localeCompare(b))]);
}

const sameIds = (a = [], b = []) => a.length === b.length && [...a].sort().join('|') === [...b].sort().join('|');

export class Scheduler {
  /**
   * now: relógio (substituído nos testes); tickMs: intervalo das verificações; graceMs: atraso
   * tolerado antes de um horário contar como perdido.
   */
  constructor({ store, manager, now = () => new Date(), tickMs = TICK_MS, graceMs = GRACE_MS } = {}) {
    this.store = store;
    this.manager = manager;
    this.now = now;
    this.tickMs = tickMs;
    this.graceMs = graceMs;
    this.timer = null;
    this.ticking = null;
    this.stopped = false;
    this.firing = new Set(); // agendamentos com uma execução sendo iniciada
    this.pruning = Promise.resolve();
    // Uma execução agendada com exclusão que esperou na fila só exclui se o agendamento ainda
    // autorizar exatamente aquilo quando ela começa.
    if (manager) manager.deletionGuard = (scan) => this.deletionGuard(scan);
  }

  /** Começa a verificar os horários (a primeira verificação trata os horários perdidos). */
  start() {
    this.stopped = false;
    for (const schedule of this.store.listSchedules()) {
      try {
        validateRule(schedule.rule);
      } catch (err) {
        console.warn(`[CLEAN] Agendamento "${schedule.name}" com regra inválida: ${err.message}`);
      }
    }
    if (!this.timer) {
      this.timer = setInterval(() => this.tick(), this.tickMs);
      this.timer.unref?.();
    }
    return this.tick();
  }

  /** Para de iniciar análises (aguarda a verificação em andamento). */
  async stop() {
    this.stopped = true;
    clearInterval(this.timer);
    this.timer = null;
    await this.ticking;
    await this.pruning;
  }

  /** Verifica os agendamentos (uma verificação por vez). Os erros vão para o console. */
  tick() {
    this.ticking ||= this.#tick()
      .catch((err) => console.error('[CLEAN] Falha no agendador:', err))
      .finally(() => {
        this.ticking = null;
      });
    return this.ticking;
  }

  async #tick() {
    this.collectOutcomes();
    await this.pruning;
    for (const { id } of [...this.store.listSchedules()]) {
      if (this.stopped) return;
      const schedule = this.store.getSchedule(id);
      if (!schedule?.enabled || !schedule.nextRunAt) continue;
      try {
        const now = this.now();
        const due = new Date(schedule.nextRunAt);
        // Relógio corrigido para trás (estava adiantado): a próxima execução guardada ficou além
        // da que a regra prevê a partir de agora.
        const expected = nextOccurrence(schedule.rule, now, { remaining: remainingOf(schedule) });
        if (expected && due > expected) {
          console.warn(`[CLEAN] Agendamento "${schedule.name}": próxima execução recalculada para ${fmt(expected)} (o relógio do servidor voltou).`);
          this.store.updateScheduleState(id, { nextRunAt: expected.toISOString() });
          continue;
        }
        if (due > now) continue;
        await this.#due(schedule, due, now);
      } catch (err) {
        // Um agendamento com dados inválidos (ex.: db.json editado) não impede os demais.
        console.error(`[CLEAN] Falha no agendamento "${schedule.name}":`, err);
        this.store.updateScheduleState(id, { nextRunAt: null });
        this.#record(schedule, { status: 'failed', trigger: 'schedule', message: `Falha no agendador: ${err.message}. O agendamento foi parado; edite-o para voltar a executar.` });
      }
    }
  }

  /** Próxima execução de um agendamento ativo a partir de agora (ISO) ou null. */
  plan(schedule, now = this.now()) {
    if (!schedule.enabled || !schedule.rule) return null; // sem regra: só manualmente
    try {
      return nextOccurrence(schedule.rule, now, { remaining: remainingOf(schedule) })?.toISOString() || null;
    } catch {
      return null; // regra inválida: aparece como problema na lista
    }
  }

  async #due(schedule, due, now) {
    const rule = JSON.stringify(schedule.rule);
    // A próxima execução é marcada antes de iniciar esta: nunca executa o mesmo horário duas vezes.
    const next = nextOccurrence(schedule.rule, now, { remaining: remainingOf(schedule) });
    this.store.updateScheduleState(schedule.id, { nextRunAt: next ? next.toISOString() : null });
    const late = now - due > this.graceMs;
    let scan = null;
    if (!late) scan = await this.#fire(schedule, { trigger: 'schedule', plannedFor: due });
    else {
      const others = countBetween(schedule.rule, due, now);
      const also = others ? ` Outro${others > 1 ? 's' : ''} ${others} horário${others > 1 ? 's' : ''} também ${others > 1 ? 'foram perdidos' : 'foi perdido'}.` : '';
      const lost = `O horário de ${fmt(due)} foi perdido: o CLEAN estava parado (ou o computador, desligado).${also}`;
      if (!schedule.catchUp) {
        this.#record(schedule, { status: 'missed', trigger: 'schedule', plannedFor: due.toISOString(), message: lost });
        return;
      }
      scan = await this.#fire(schedule, { trigger: 'catch-up', plannedFor: due, note: `Execução atrasada. ${lost}` });
    }
    if (!scan) return;
    // "Depois de N execuções": esta conta; ao chegar a N, o agendamento termina. Se a regra mudou
    // enquanto a execução era iniciada, a nova regra começa a contagem do zero.
    const current = this.store.getSchedule(schedule.id);
    if (!current || JSON.stringify(current.rule) !== rule) return;
    const countDone = (current.countDone || 0) + 1;
    const ended = current.rule?.end === 'count' && current.rule.frequency !== 'once' && countDone >= current.rule.count;
    this.store.updateScheduleState(schedule.id, ended ? { countDone, nextRunAt: null } : { countDone });
  }

  /**
   * Executa agora ("Executar agora" na interface). Lança ScanError se não for possível (a falha
   * também fica no histórico). Não muda a próxima execução programada nem conta como execução
   * do término "depois de N execuções".
   */
  async runNow(id, by, { simulate = false } = {}) {
    this.collectOutcomes();
    const schedule = this.store.getSchedule(id);
    if (!schedule) throw new ScanError('Agendamento não encontrado.', 404);
    return this.#fire(schedule, { trigger: 'manual', by, simulate });
  }

  /** Execução deste agendamento na fila ou em andamento, se houver. */
  #activeRun(schedule) {
    return this.store.listScans().find((s) => s.scheduleId === schedule.id && this.manager.isActive(s.id)) || null;
  }

  /** Inicia a análise do agendamento e registra o resultado no histórico. */
  async #fire(schedule, { trigger, plannedFor = null, note = '', by = null, simulate = false }) {
    const base = { trigger, plannedFor: plannedFor ? plannedFor.toISOString() : null };
    const manual = trigger === 'manual';
    // Verificação e reserva sem pausa (sem await): duas execuções do mesmo agendamento nunca
    // começam juntas (ex.: "Executar agora" duas vezes, ou no mesmo instante do horário).
    if (this.firing.has(schedule.id) || this.#activeRun(schedule)) {
      const message = 'A execução anterior deste agendamento ainda está em andamento (ou na fila).';
      if (manual) throw new ScanError(message, 409);
      this.#record(schedule, { ...base, status: 'skipped', message });
      return null;
    }
    this.firing.add(schedule.id);
    try {
      return await this.#start(schedule, { base, manual, note, by, simulate });
    } finally {
      this.firing.delete(schedule.id);
    }
  }

  async #start(schedule, { base, manual, note, by, simulate = false }) {
    const now = this.now();
    let scan;
    let period;
    let signature;
    try {
      // Simulação ("Simular agora"): lista o que seria excluído, sem excluir nem conferir a exclusão.
      const problems = scheduleProblems(this.store, simulate ? { ...schedule, action: 'analyze' } : schedule);
      if (problems.length) throw new ScanError(problems[0]);
      signature = coverageSignature(this.store, schedule);
      period = this.#period(schedule, now, signature);
      const mail = schedule.kind === 'mail';
      const deleting = schedule.action === 'delete' && !simulate;
      const retention = schedule.purpose === 'retention';
      // Critérios da exclusão com que esta execução começa (iguais aos confirmados: sem problemas).
      const criteria = deleting ? deletionCriteria(this.store, schedule) : null;
      const confirmed = schedule.deleteConfirmation;
      const what = retention ? 'política de retenção' : 'agendamento';
      const origin = `${what} "${schedule.name}"${deleting && confirmed ? ` (${retention ? 'exclusão' : 'exclusão automática'} confirmada por ${confirmed.by} em ${fmt(confirmed.at)})` : ''}${simulate ? ' (simulação)' : ''}`;
      const request = retention
        ? { retention: schedule.retention, options: { ...schedule.options, deleteMatches: deleting } }
        : { listIds: schedule.listIds, options: { ...schedule.options, deleteMatches: deleting, [mail ? 'receivedAfter' : 'modifiedAfter']: period.from ? period.from.toISOString() : null } };
      scan = await this.manager.start(
        {
          kind: schedule.kind,
          name: `${schedule.name}${simulate ? ' (simulação)' : ''} – ${fmt(now)}`,
          [mail ? 'sourceIds' : 'repositoryIds']: schedule.targetIds,
          ...request,
          confirmDelete: deleting ? 'EXCLUIR' : '',
        },
        { by: manual ? `${by} (Executar agora, ${origin})` : origin, schedule: { id: schedule.id, name: schedule.name, criteria, enabled: schedule.enabled } },
      );
    } catch (err) {
      const message = err instanceof ScanError ? err.message : `Falha ao iniciar a análise: ${err.message}`;
      if (!(err instanceof ScanError)) console.error('[CLEAN] Falha ao iniciar a análise agendada:', err);
      this.#record(schedule, { ...base, status: 'failed', by: manual ? by : null, message });
      if (manual) throw err instanceof ScanError ? err : new ScanError(message, 500);
      return null;
    }
    const notes = [note, period.text].filter(Boolean);
    for (const message of notes) this.store.appendLog(scan.id, { level: 'info', message: `Agendamento "${schedule.name}": ${message}` });
    this.#record(schedule, {
      ...base,
      status: 'started',
      scanId: scan.id,
      scanName: scan.name,
      signature,
      full: period.full,
      base: period.base,
      from: period.from ? period.from.toISOString() : null,
      by: manual ? by : null,
      ...(simulate ? { simulated: true } : {}),
      message: notes.join(' '),
    });
    return scan;
  }

  /**
   * Período desta execução: todos os itens, os alterados nos últimos N dias ou, na incremental,
   * os alterados desde o início da última execução concluída sem falhas de acesso e com a mesma
   * configuração (com uma margem), com uma análise completa a cada `fullEvery` execuções.
   */
  #period(schedule, now, signature) {
    const p = schedule.period || { type: 'all' };
    const noun = schedule.kind === 'mail' ? 'mensagens recebidas' : 'arquivos alterados';
    if (p.type === 'days') {
      const from = new Date(+now - p.days * DAY_MS);
      return { from, full: false, base: false, text: `somente ${noun} desde ${fmt(from)} (últimos ${p.days} dias).` };
    }
    if (p.type !== 'since-last') return { from: null, full: true, base: true, text: '' };
    const runs = (schedule.history || []).filter((h) => h.status === 'started' && h.signature === signature);
    // Uma execução que não conseguiu ler um repositório, uma conta, um site ou uma caixa inteira
    // não serve de base: o que mudou antes dela ficaria de fora das próximas.
    const complete = (h) => h.outcome?.status === 'completed' && !h.outcome.gaps && h.outcome.startedAt && (schedule.action !== 'delete' || h.outcome.deleting !== false);
    const baseline = runs.find((h) => h.base && complete(h));
    if (!baseline) {
      return { from: null, full: true, base: true, text: 'análise completa (não há execução anterior concluída, sem falhas de acesso, com os mesmos locais, termos e opções).' };
    }
    const every = p.fullEvery || 7;
    const lastFull = runs.findIndex((h) => h.full && complete(h));
    if (lastFull === -1 || lastFull >= every - 1) return { from: null, full: true, base: true, text: `análise completa periódica (a cada ${every} execuções).` };
    const from = new Date(Date.parse(baseline.outcome.startedAt) - INCREMENTAL_MARGIN_MS);
    return { from, full: false, base: true, text: `análise incremental: somente ${noun} desde ${fmt(from)} (1 hora antes do início da última execução concluída).` };
  }

  #record(schedule, entry) {
    const current = this.store.getSchedule(schedule.id);
    if (!current) return;
    const item = { id: crypto.randomUUID(), at: this.now().toISOString(), ...entry };
    const patch = { history: [item, ...(current.history || [])].slice(0, MAX_HISTORY) };
    if (entry.status === 'started') patch.runCount = (current.runCount || 0) + 1;
    this.store.updateScheduleState(schedule.id, patch);
  }

  /**
   * Anota no histórico o resultado das análises que terminaram (o relatório pode ser excluído
   * depois) e, para os agendamentos com retenção, exclui os relatórios antigos.
   */
  collectOutcomes() {
    for (const schedule of this.store.listSchedules()) {
      let changed = false;
      for (const h of schedule.history || []) {
        if (h.status !== 'started' || h.outcome) continue;
        const scan = this.store.getScan(h.scanId);
        if (scan && !FINISHED.has(scan.status)) continue;
        if (scan && this.manager.isActive(scan.id)) continue;
        h.outcome = scan
          ? {
              status: scan.status,
              startedAt: scan.startedAt,
              finishedAt: scan.finishedAt,
              matched: (scan.kind === 'mail' ? scan.stats?.messagesMatched : scan.stats?.filesMatched) || 0,
              deleted: scan.stats?.deleted || 0,
              errors: scan.stats?.errors || 0,
              gaps: scan.stats?.gaps || 0,
              // Com exclusão: se ela valeu até o fim (não foi desligada na fila nem durante a execução).
              deleting: Boolean(scan.options?.deleteMatches) && !scan.deletionRevoked,
              error: scan.error || null,
            }
          : { status: 'removed' };
        changed = true;
      }
      if (!changed) continue;
      this.store.updateScheduleState(schedule.id, { history: [...schedule.history] });
      if (schedule.keepLast) {
        const id = schedule.id;
        this.pruning = this.pruning.then(() => this.#prune(id)).catch((err) => console.error('[CLEAN] Falha ao limpar os relatórios do agendamento:', err));
      }
    }
  }

  /**
   * Retenção, depois que uma execução termina: guarda os relatórios até o `keepLast`-ésimo
   * concluído (os mais novos, de qualquer situação, ficam) e sempre o da última análise completa
   * concluída; os mais antigos são excluídos. Relatórios em uso não são tocados.
   */
  async #prune(id) {
    const schedule = this.store.getSchedule(id);
    const keep = Number(schedule?.keepLast) || 0;
    if (!keep) return;
    // Mais recentes primeiro (no empate da data, a posição no cadastro decide).
    const own = this.store
      .listScans()
      .map((scan, index) => ({ scan, index }))
      .filter(({ scan }) => scan.scheduleId === id)
      .sort((a, b) => String(b.scan.createdAt).localeCompare(String(a.scan.createdAt)) || b.index - a.index)
      .map(({ scan }) => scan);
    // Concluído sem falhas de acesso (um relatório de quando um local estava fora do ar não conta).
    const good = (scan) => scan.status === 'completed' && !scan.stats?.gaps;
    const lastFull = (schedule.history || []).find((h) => h.status === 'started' && h.full && h.outcome?.status === 'completed' && !h.outcome.gaps)?.scanId;
    let completed = 0;
    let removed = 0;
    for (const scan of own) {
      const kept = completed < keep;
      if (good(scan)) completed++;
      if (kept || scan.id === lastFull || !FINISHED.has(scan.status)) continue;
      if (this.manager.isActive(scan.id) || this.manager.hasItemDeletion(scan.id)) continue;
      try {
        await this.store.deleteScan(scan.id);
        removed++;
      } catch (err) {
        console.error(`[CLEAN] Não foi possível excluir o relatório antigo "${scan.name}":`, err.message);
      }
    }
    if (removed && own[0]) {
      this.store.appendLog(own[0].id, {
        level: 'info',
        message: `Agendamento "${schedule.name}": ${removed} relatório(s) antigo(s) excluído(s) (o agendamento guarda até o ${keep}º relatório concluído mais recente e o da última análise completa; o registro geral de exclusões é mantido).`,
      });
    }
  }

  /**
   * Motivo para uma execução agendada NÃO excluir (ou null): o agendamento foi excluído, pausado
   * ou passou a "Somente analisar", mudou de locais, listas ou opções, ou a confirmação da
   * exclusão não vale mais. Consultado quando a execução sai da fila e quando o agendamento muda.
   */
  deletionGuard(scan) {
    const schedule = this.store.getSchedule(scan.scheduleId);
    if (!schedule) return 'o agendamento foi excluído';
    // Pausado depois de a execução ser criada ("Executar agora" num agendamento pausado vale).
    if (!schedule.enabled && scan.scheduleEnabled !== false) return 'o agendamento foi pausado';
    if (schedule.action !== 'delete') return 'o agendamento passou a "Somente analisar"';
    // Os critérios com que a execução começou (termos, pastas ignoradas, locais protegidos) precisam
    // ser os confirmados agora: uma nova confirmação com outros critérios não vale para ela.
    if (!scan.scheduleCriteria || scan.scheduleCriteria !== schedule.deleteConfirmation?.criteria) {
      return 'a exclusão do agendamento foi confirmada de novo com outros termos, pastas ignoradas ou locais protegidos depois que esta execução foi criada';
    }
    const targets = scan.kind === 'mail' ? scan.sourceIds : scan.repositoryIds;
    if (!sameIds(targets, schedule.targetIds) || !sameIds(scan.listIds, schedule.listIds)) return 'os locais ou as listas do agendamento foram alterados';
    if (schedule.purpose === 'retention') {
      // A política (critério, idade máxima, nomes, limite e forma de exclusão) precisa ser a mesma.
      // eslint-disable-next-line no-unused-vars
      const { cutoff, ...policy } = scan.retention || {};
      if (JSON.stringify(policy) !== JSON.stringify(schedule.retention)) return 'a política de retenção foi alterada';
      const problems = scheduleProblems(this.store, schedule);
      return problems.length ? problems[0] : null;
    }
    // eslint-disable-next-line no-unused-vars
    const strip = ({ deleteMatches, modifiedAfter, receivedAfter, concurrency, ...rest } = {}) => JSON.stringify(Object.entries(rest).sort(([a], [b]) => a.localeCompare(b)));
    let current;
    try {
      current = (scan.kind === 'mail' ? sanitizeMailOptions : sanitizeOptions)({ ...schedule.options, deleteMatches: false });
    } catch {
      return 'as opções do agendamento são inválidas';
    }
    if (strip(scan.options) !== strip(current)) return 'as opções do agendamento foram alteradas';
    const problems = scheduleProblems(this.store, schedule);
    return problems.length ? problems[0] : null;
  }

  /**
   * Depois de uma alteração no agendamento (edição, pausa, exclusão): as execuções dele em
   * andamento deixam de excluir se a exclusão não vale mais. As que estão na fila são conferidas
   * quando começam.
   */
  reviewRuns(scheduleId) {
    for (const scan of this.store.listScans()) {
      if (scan.scheduleId !== scheduleId || !scan.options?.deleteMatches || scan.deletionRevoked || !this.manager.isActive(scan.id)) continue;
      const reason = this.deletionGuard(scan);
      if (reason) this.manager.revokeScanDeletion(scan.id, reason);
    }
  }

  /** O mesmo para todas as execuções agendadas com exclusão (depois de alterar um cadastro). */
  reviewAll() {
    const ids = new Set(this.store.listScans().filter((s) => s.scheduleId && s.options?.deleteMatches && this.manager.isActive(s.id)).map((s) => s.scheduleId));
    for (const id of ids) this.reviewRuns(id);
  }
}
