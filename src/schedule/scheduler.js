// Agendador: inicia as análises agendadas nos horários das regras de recorrência.
//
// A cada poucos segundos, os agendamentos ativos cujo horário chegou são executados (a análise
// entra na fila do gerenciador como qualquer outra). Regras:
// - Sobreposição: se a execução anterior do mesmo agendamento ainda estiver em andamento (ou na
//   fila), a nova é pulada e fica registrada no histórico.
// - Horário perdido (CLEAN parado ou computador desligado): na volta, o agendamento é executado
//   uma vez (se "executar assim que possível" estiver marcado) ou o horário fica registrado como
//   perdido; em ambos os casos, os demais horários perdidos não são repetidos.
// - Exclusão automática: vale somente para o que foi confirmado ao salvar o agendamento. Se o
//   cadastro de um repositório ou conexão mudar depois (exclusão desligada, outro caminho, outras
//   contas, sites ou caixas, ou exclusão que passou a ser definitiva), a execução falha com a
//   explicação, até o agendamento ser salvo e confirmado de novo.
import crypto from 'node:crypto';
import { nextOccurrence, countBetween } from './recurrence.js';
import { ScanError } from '../scan/manager.js';
import { deletionScope, mailDeletionScope, isCloudRepo } from '../scan/delete.js';

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

/**
 * Problemas que impedem a execução do agendamento agora (cadastros excluídos, listas sem termos,
 * exclusão automática que não vale mais). Lista vazia: pode executar.
 */
export function scheduleProblems(store, schedule) {
  const problems = [];
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
  if (lists.length && lists.every((l) => !(l.terms || []).length)) problems.push('As listas de referência do agendamento não têm termos.');
  if (schedule.action === 'delete') {
    const pins = schedule.deleteConfirmation?.targets || {};
    const again = 'Edite o agendamento e confirme a exclusão de novo (ou escolha "Somente analisar").';
    for (const t of targets) {
      const pin = pins[t.id];
      if (!t.allowDelete) problems.push(`A exclusão foi desligada no cadastro ${noun.registry} "${t.name}". Permita a exclusão de novo ou edite o agendamento para "Somente analisar".`);
      else if (!pin || pin.scope !== scopeOf(schedule.kind, t)) {
        const what = schedule.kind === 'mail' ? 'a conta, o servidor ou as caixas' : 'o caminho, as contas ou os sites';
        problems.push(`O cadastro ${noun.registry} "${t.name}" mudou (${what}) depois que a exclusão automática foi confirmada neste agendamento. ${again}`);
      } else if (pin.mode !== 'permanent' && modeOf(schedule.kind, t) === 'permanent') {
        problems.push(`A exclusão em "${t.name}" passou a ser definitiva depois que a exclusão automática foi confirmada neste agendamento. ${again}`);
      }
    }
  }
  return problems;
}

/** O que a análise alcança (sem o período): muda quando os locais, os termos ou as opções mudam. */
export function coverageSignature(store, schedule) {
  const kind = schedule.kind;
  const targets = [...(schedule.targetIds || [])].sort().map((id) => {
    const t = targetOf(store, kind, id);
    if (!t) return [id, null];
    return kind === 'mail' ? [id, mailDeletionScope(t), t.excludeMailboxes || [], t.excludeFolders || []] : [id, deletionScope(t), t.exclude || [], t.cloud?.exclude || []];
  });
  const terms = (schedule.listIds || [])
    .flatMap((id) => (store.getList(id)?.terms || []).map((t) => JSON.stringify([t.type, t.value, Boolean(t.wholeWord), t.validator || null])))
    .sort();
  // eslint-disable-next-line no-unused-vars
  const { concurrency, deleteMatches, modifiedAfter, receivedAfter, ...options } = schedule.options || {};
  const json = JSON.stringify([kind, targets, terms, Object.entries(options).sort(([a], [b]) => a.localeCompare(b))]);
  return crypto.createHash('sha256').update(json).digest('hex').slice(0, 32);
}

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
  }

  /** Começa a verificar os horários (a primeira verificação trata os horários perdidos). */
  start() {
    this.stopped = false;
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
    for (const { id } of [...this.store.listSchedules()]) {
      if (this.stopped) return;
      const schedule = this.store.getSchedule(id);
      if (!schedule?.enabled || !schedule.nextRunAt) continue;
      const due = new Date(schedule.nextRunAt);
      const now = this.now();
      if (due > now) continue;
      await this.#due(schedule, due, now);
    }
  }

  /** Próxima execução de um agendamento ativo a partir de agora (ISO) ou null. */
  plan(schedule, now = this.now()) {
    if (!schedule.enabled) return null;
    return nextOccurrence(schedule.rule, now)?.toISOString() || null;
  }

  async #due(schedule, due, now) {
    // A próxima execução é marcada antes de iniciar esta: nunca executa o mesmo horário duas vezes.
    const next = nextOccurrence(schedule.rule, now);
    this.store.updateScheduleState(schedule.id, { nextRunAt: next ? next.toISOString() : null });
    const late = now - due > this.graceMs;
    if (!late) return this.#fire(schedule, { trigger: 'schedule', plannedFor: due });
    const others = countBetween(schedule.rule, due, now);
    const also = others ? ` Outro${others > 1 ? 's' : ''} ${others} horário${others > 1 ? 's' : ''} também ${others > 1 ? 'foram perdidos' : 'foi perdido'}.` : '';
    const lost = `O horário de ${fmt(due)} foi perdido: o CLEAN estava parado (ou o computador, desligado).${also}`;
    if (!schedule.catchUp) {
      this.#record(schedule, { status: 'missed', trigger: 'schedule', plannedFor: due.toISOString(), message: lost });
      return null;
    }
    return this.#fire(schedule, { trigger: 'catch-up', plannedFor: due, note: `Execução atrasada. ${lost}` });
  }

  /**
   * Executa agora ("Executar agora" na interface). Lança ScanError se não for possível (a falha
   * também fica no histórico). Não muda a próxima execução programada.
   */
  async runNow(id, by) {
    this.collectOutcomes();
    const schedule = this.store.getSchedule(id);
    if (!schedule) throw new ScanError('Agendamento não encontrado.', 404);
    return this.#fire(schedule, { trigger: 'manual', by });
  }

  /** Inicia a análise do agendamento e registra o resultado no histórico. */
  async #fire(schedule, { trigger, plannedFor = null, note = '', by = null }) {
    const now = this.now();
    const base = { trigger, plannedFor: plannedFor ? plannedFor.toISOString() : null };
    const manual = trigger === 'manual';
    const previous = (schedule.history || []).find((h) => h.status === 'started');
    if (previous && this.manager.isActive(previous.scanId)) {
      const message = 'A execução anterior deste agendamento ainda está em andamento (ou na fila).';
      if (manual) throw new ScanError(message, 409);
      this.#record(schedule, { ...base, status: 'skipped', message });
      return null;
    }
    try {
      const problems = scheduleProblems(this.store, schedule);
      if (problems.length) throw new ScanError(problems[0]);
      const signature = coverageSignature(this.store, schedule);
      const period = this.#period(schedule, now, signature);
      const mail = schedule.kind === 'mail';
      const deleting = schedule.action === 'delete';
      const confirmed = schedule.deleteConfirmation;
      const origin = `agendamento "${schedule.name}"${deleting && confirmed ? ` (exclusão automática confirmada por ${confirmed.by} em ${fmt(confirmed.at)})` : ''}`;
      const scan = await this.manager.start(
        {
          kind: schedule.kind,
          name: `${schedule.name} – ${fmt(now)}`,
          [mail ? 'sourceIds' : 'repositoryIds']: schedule.targetIds,
          listIds: schedule.listIds,
          options: { ...schedule.options, deleteMatches: deleting, [mail ? 'receivedAfter' : 'modifiedAfter']: period.from ? period.from.toISOString() : null },
          confirmDelete: deleting ? 'EXCLUIR' : '',
        },
        { by: manual ? `${by} (Executar agora, ${origin})` : origin, schedule: { id: schedule.id, name: schedule.name } },
      );
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
        message: notes.join(' '),
      });
      await this.#prune(schedule, scan.id);
      return scan;
    } catch (err) {
      const message = err instanceof ScanError ? err.message : `Falha ao iniciar a análise: ${err.message}`;
      if (!(err instanceof ScanError)) console.error('[CLEAN] Falha ao iniciar a análise agendada:', err);
      this.#record(schedule, { ...base, status: 'failed', by: manual ? by : null, message });
      if (manual) throw err instanceof ScanError ? err : new ScanError(message, 500);
      return null;
    }
  }

  /**
   * Período desta execução: todos os itens, os alterados nos últimos N dias ou, na incremental,
   * os alterados desde o início da última execução concluída com a mesma configuração (com uma
   * margem), com uma análise completa a cada `fullEvery` execuções.
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
    const baseline = runs.find((h) => h.base && h.outcome?.status === 'completed' && h.outcome.startedAt);
    if (!baseline) {
      return { from: null, full: true, base: true, text: 'análise completa (não há execução anterior concluída com os mesmos locais, termos e opções).' };
    }
    if (p.fullEvery > 0) {
      const lastFull = runs.findIndex((h) => h.full && h.outcome?.status === 'completed');
      if (lastFull === -1 || lastFull >= p.fullEvery - 1) return { from: null, full: true, base: true, text: `análise completa periódica (a cada ${p.fullEvery} execuções).` };
    }
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

  /** Anota no histórico o resultado das análises que terminaram (o relatório pode ser excluído depois). */
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
              error: scan.error || null,
            }
          : { status: 'removed' };
        changed = true;
      }
      if (changed) this.store.updateScheduleState(schedule.id, { history: [...schedule.history] });
    }
  }

  /** Mantém só os `keepLast` relatórios mais recentes do agendamento (0 = todos). */
  async #prune(schedule, currentScanId) {
    const keep = Number(schedule.keepLast) || 0;
    if (!keep) return;
    // Mais recentes primeiro (no empate da data, a posição no cadastro decide).
    const own = this.store
      .listScans()
      .map((scan, index) => ({ scan, index }))
      .filter(({ scan }) => scan.scheduleId === schedule.id)
      .sort((a, b) => String(b.scan.createdAt).localeCompare(String(a.scan.createdAt)) || b.index - a.index)
      .map(({ scan }) => scan);
    this.collectOutcomes(); // o resultado das execuções fica no histórico mesmo sem o relatório
    let removed = 0;
    for (const scan of own.slice(keep)) {
      if (scan.id === currentScanId || this.manager.isActive(scan.id) || this.manager.hasItemDeletion(scan.id)) continue;
      await this.store.deleteScan(scan.id);
      removed++;
    }
    if (removed) {
      this.store.appendLog(currentScanId, {
        level: 'info',
        message: `Agendamento "${schedule.name}": ${removed} relatório(s) antigo(s) excluído(s) (o agendamento guarda os ${keep} mais recentes; o registro geral de exclusões é mantido).`,
      });
    }
  }
}
