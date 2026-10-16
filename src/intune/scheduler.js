// Agendador do reinício dos dispositivos no Intune (Microsoft Graph): executa os agendamentos de
// reinício nos horários das regras de recorrência e registra o resultado (sucesso/falha) por
// aparelho no histórico de cada agendamento.
//
// A cada poucos segundos, os agendamentos ativos cujo horário chegou são executados. Regras (iguais
// às das análises agendadas):
// - Sobreposição: se uma execução do mesmo agendamento ainda estiver em andamento, a nova é pulada
//   e fica registrada no histórico.
// - Horário perdido (CLEAN parado ou computador desligado): na volta, o agendamento é executado uma
//   vez (se "executar assim que possível" estiver marcado) ou o horário fica registrado como
//   perdido; em ambos os casos, os demais horários perdidos não são repetidos.
// - "Depois de N execuções": contam as execuções de fato iniciadas pelo agendamento.
import crypto from 'node:crypto';
import { nextOccurrence, countBetween, validateRule } from '../schedule/recurrence.js';
import { IntuneConnector } from './connector.js';
import { runReboot, STALE_SYNC_HOURS } from './reboot.js';

const TICK_MS = 15000;
// Atraso tolerado (servidor ocupado): acima disso, o horário conta como perdido.
const GRACE_MS = 5 * 60000;
export const MAX_HISTORY = 50;
// Teto de aparelhos guardados por execução no histórico (o resumo e os totais sempre contam todos).
export const MAX_RUN_DEVICES = 500;

const pad = (n) => String(n).padStart(2, '0');

/** Data e hora do servidor (dd/mm/aaaa hh:mm). */
export function fmt(value) {
  const d = new Date(value);
  return `${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${d.getFullYear()} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

/** Quantas execuções ainda faltam no término "depois de N execuções" (sem esse término: infinitas). */
export function remainingOf(job) {
  const rule = job.rule || {};
  if (rule.frequency === 'once' || rule.end !== 'count') return Infinity;
  return Math.max(0, rule.count - (job.countDone || 0));
}

/**
 * Problemas que impedem a execução do agendamento de reinício agora (regra inválida, conexão
 * excluída ou que não é do Microsoft 365). Lista vazia: pode executar.
 */
export function jobProblems(store, job) {
  const problems = [];
  try {
    validateRule(job.rule);
  } catch (err) {
    problems.push(`A regra de recorrência salva é inválida (${err.message}). Edite o agendamento.`);
  }
  const source = store.getMailSource(job.sourceId);
  if (!source) problems.push(`A conexão Microsoft 365 "${job.sourceName || job.sourceId}" foi excluída do cadastro. Edite o agendamento.`);
  else if (source.type !== 'graph') problems.push(`A conexão "${source.name}" não é do Microsoft 365 (Graph): o reinício pelo Intune precisa de uma conexão do Microsoft 365.`);
  return problems;
}

export class IntuneScheduler {
  /**
   * now: relógio (substituído nos testes); tickMs: intervalo das verificações; graceMs: atraso
   * tolerado antes de um horário contar como perdido; endpoints: troca os endereços do Graph (testes).
   */
  constructor({ store, endpoints = {}, now = () => new Date(), tickMs = TICK_MS, graceMs = GRACE_MS } = {}) {
    this.store = store;
    this.endpoints = endpoints;
    this.now = now;
    this.tickMs = tickMs;
    this.graceMs = graceMs;
    this.timer = null;
    this.ticking = null;
    this.stopped = false;
    this.firing = new Set(); // agendamentos com uma execução em andamento
  }

  /** Começa a verificar os horários (a primeira verificação trata os horários perdidos). */
  start() {
    this.stopped = false;
    for (const job of this.store.listIntuneJobs()) {
      try {
        validateRule(job.rule);
      } catch (err) {
        console.warn(`[CLEAN] Reinício agendado "${job.name}" com regra inválida: ${err.message}`);
      }
    }
    if (!this.timer) {
      this.timer = setInterval(() => this.tick(), this.tickMs);
      this.timer.unref?.();
    }
    return this.tick();
  }

  /** Para de iniciar reinícios (aguarda a verificação em andamento). */
  async stop() {
    this.stopped = true;
    clearInterval(this.timer);
    this.timer = null;
    await this.ticking;
  }

  /** Verifica os agendamentos (uma verificação por vez). Os erros vão para o console. */
  tick() {
    this.ticking ||= this.#tick()
      .catch((err) => console.error('[CLEAN] Falha no agendador de reinício:', err))
      .finally(() => {
        this.ticking = null;
      });
    return this.ticking;
  }

  async #tick() {
    for (const { id } of [...this.store.listIntuneJobs()]) {
      if (this.stopped) return;
      const job = this.store.getIntuneJob(id);
      if (!job?.enabled || !job.nextRunAt) continue;
      try {
        const now = this.now();
        const due = new Date(job.nextRunAt);
        // Relógio corrigido para trás: a próxima execução guardada ficou além da que a regra prevê agora.
        const expected = nextOccurrence(job.rule, now, { remaining: remainingOf(job) });
        if (expected && due > expected) {
          console.warn(`[CLEAN] Reinício agendado "${job.name}": próxima execução recalculada para ${fmt(expected)} (o relógio do servidor voltou).`);
          this.store.updateIntuneJobState(id, { nextRunAt: expected.toISOString() });
          continue;
        }
        if (due > now) continue;
        await this.#due(job, due, now);
      } catch (err) {
        // Um agendamento com dados inválidos (ex.: db.json editado) não impede os demais.
        console.error(`[CLEAN] Falha no reinício agendado "${job.name}":`, err);
        this.store.updateIntuneJobState(id, { nextRunAt: null });
        this.#record(job, { status: 'failed', trigger: 'schedule', message: `Falha no agendador: ${err.message}. O agendamento foi parado; edite-o para voltar a executar.` });
      }
    }
  }

  /** Próxima execução de um agendamento ativo a partir de agora (ISO) ou null. */
  plan(job, now = this.now()) {
    if (!job.enabled || !job.rule) return null;
    try {
      return nextOccurrence(job.rule, now, { remaining: remainingOf(job) })?.toISOString() || null;
    } catch {
      return null; // regra inválida: aparece como problema na lista
    }
  }

  async #due(job, due, now) {
    const rule = JSON.stringify(job.rule);
    // A próxima execução é marcada antes de iniciar esta: nunca executa o mesmo horário duas vezes.
    const next = nextOccurrence(job.rule, now, { remaining: remainingOf(job) });
    this.store.updateIntuneJobState(job.id, { nextRunAt: next ? next.toISOString() : null });
    const late = now - due > this.graceMs;
    let run = null;
    if (!late) run = await this.#fire(job, { trigger: 'schedule', plannedFor: due });
    else {
      const others = countBetween(job.rule, due, now);
      const also = others ? ` Outro${others > 1 ? 's' : ''} ${others} horário${others > 1 ? 's' : ''} também ${others > 1 ? 'foram perdidos' : 'foi perdido'}.` : '';
      const lost = `O horário de ${fmt(due)} foi perdido: o CLEAN estava parado (ou o computador, desligado).${also}`;
      if (!job.catchUp) {
        this.#record(job, { status: 'missed', trigger: 'schedule', plannedFor: due.toISOString(), message: lost });
        return;
      }
      run = await this.#fire(job, { trigger: 'catch-up', plannedFor: due, note: `Execução atrasada. ${lost}` });
    }
    if (!run || run.status === 'skipped') return;
    // "Depois de N execuções": esta conta; ao chegar a N, o agendamento termina. Se a regra mudou
    // enquanto a execução era iniciada, a nova regra começa a contagem do zero.
    const current = this.store.getIntuneJob(job.id);
    if (!current || JSON.stringify(current.rule) !== rule) return;
    const countDone = (current.countDone || 0) + 1;
    const ended = current.rule?.end === 'count' && current.rule.frequency !== 'once' && countDone >= current.rule.count;
    this.store.updateIntuneJobState(job.id, ended ? { countDone, nextRunAt: null } : { countDone });
  }

  /**
   * Executa agora ("Executar agora" na interface). simulate: só lista os aparelhos do escopo (não
   * reinicia). Não muda a próxima execução programada nem conta no término "depois de N execuções".
   */
  async runNow(id, by, { simulate = false } = {}) {
    const job = this.store.getIntuneJob(id);
    if (!job) {
      const err = new Error('Agendamento de reinício não encontrado.');
      err.status = 404;
      throw err;
    }
    return this.#fire(job, { trigger: 'manual', by, simulate });
  }

  /** Verifica a sobreposição e executa o reinício, registrando o resultado no histórico. */
  async #fire(job, { trigger, plannedFor = null, note = '', by = null, simulate = false }) {
    const base = { trigger, plannedFor: plannedFor ? plannedFor.toISOString() : null, by: trigger === 'manual' ? by : null, simulate };
    // Reserva sem pausa (sem await): duas execuções do mesmo agendamento nunca começam juntas.
    if (this.firing.has(job.id)) {
      const message = 'A execução anterior deste agendamento ainda está em andamento.';
      if (trigger === 'manual') {
        const err = new Error(message);
        err.status = 409;
        throw err;
      }
      return this.#record(job, { ...base, status: 'skipped', message });
    }
    this.firing.add(job.id);
    try {
      return await this.#run(job, { base, note });
    } finally {
      this.firing.delete(job.id);
    }
  }

  async #run(job, { base, note }) {
    const problems = jobProblems(this.store, job);
    if (problems.length) {
      const message = problems[0];
      if (base.trigger === 'manual') {
        const err = new Error(message);
        err.status = 400;
        throw err;
      }
      return this.#record(job, { ...base, status: 'failed', message });
    }
    const source = this.store.getMailSource(job.sourceId);
    const logs = [];
    const log = (level, message) => {
      logs.push({ level, message });
      if (logs.length > MAX_RUN_DEVICES) logs.shift();
    };
    if (note) log('info', note);
    const grantId = source.graph?.account?.grantId;
    const onRefreshToken = grantId ? (token) => this.store.saveRefreshToken(source.id, grantId, token) : undefined;
    const controller = new AbortController();
    const connector = new IntuneConnector(
      { ...source, secrets: this.store.openMailSecrets(source) },
      { endpoints: this.endpoints, signal: controller.signal, onRefreshToken, log },
    );
    let result;
    try {
      result = await runReboot(connector, job.scope, { log, simulate: base.simulate });
    } catch (err) {
      await connector.close().catch(() => {});
      const message = `Falha ao executar o reinício: ${err.message}`;
      if (base.trigger === 'manual') {
        const wrapped = new Error(message);
        wrapped.status = err.status && err.status < 500 ? err.status : 502;
        throw wrapped;
      }
      log('error', message);
      return this.#record(job, { ...base, status: 'failed', message, log: logs.slice(0, MAX_RUN_DEVICES) });
    }
    await connector.close().catch(() => {});
    return this.#record(job, {
      ...base,
      status: runStatus(result, base.simulate),
      startedAt: result.startedAt,
      finishedAt: result.finishedAt,
      total: result.total,
      ok: result.ok,
      failed: result.failed,
      skipped: result.skipped,
      message: summaryText(result, base.simulate),
      devices: result.devices.slice(0, MAX_RUN_DEVICES),
      devicesTruncated: result.devices.length > MAX_RUN_DEVICES,
      log: logs.slice(0, MAX_RUN_DEVICES),
    });
  }

  /** Registra uma execução no histórico do agendamento e devolve o item registrado. */
  #record(job, entry) {
    const current = this.store.getIntuneJob(job.id);
    if (!current) return null;
    const item = { id: crypto.randomUUID(), at: this.now().toISOString(), ...entry };
    const patch = { history: [item, ...(current.history || [])].slice(0, MAX_HISTORY) };
    // Execuções que de fato começaram (não as puladas nem os horários perdidos) contam no total.
    if (!['skipped', 'missed'].includes(entry.status)) patch.runCount = (current.runCount || 0) + 1;
    this.store.updateIntuneJobState(job.id, patch);
    return item;
  }
}

/** Situação da execução a partir dos totais por aparelho. */
function runStatus(result, simulate) {
  if (simulate) return 'ok';
  if (result.total === 0) return 'ok';
  if (result.failed && result.ok) return 'partial';
  if (result.failed && !result.ok) return 'failed';
  return 'ok';
}

/** Aviso sobre aparelhos provavelmente offline (quando houver). */
function offlineNote(result) {
  if (!result.stale) return '';
  const one = result.stale === 1;
  return ` ${result.stale} ${one ? 'aparelho não sincroniza' : 'aparelhos não sincronizam'} com o Intune há mais de ${STALE_SYNC_HOURS} horas e ${one ? 'pode estar' : 'podem estar'} offline — o reinício só ocorre quando ${one ? 'ele reconectar' : 'reconectarem'}.`;
}

/** Resumo em português do resultado da execução. */
function summaryText(result, simulate) {
  if (simulate) return `Simulação: ${result.total} aparelho(s) no escopo (nenhum foi reiniciado).${offlineNote(result)}`;
  if (result.total === 0) return 'Nenhum aparelho no escopo.';
  // "Reinício enviado": o rebootNow é aceito pelo Intune e entregue na próxima comunicação do aparelho.
  const parts = [`${result.ok} com reinício enviado`];
  if (result.failed) parts.push(`${result.failed} com falha`);
  if (result.skipped) parts.push(`${result.skipped} ignorado(s)`);
  return `${parts.join(', ')} de ${result.total} aparelho(s).${offlineNote(result)}`;
}
