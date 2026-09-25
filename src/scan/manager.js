// Fila de análises: cada análise roda em uma worker thread; o resultado é gravado conforme chega.
import { Worker } from 'node:worker_threads';
import { newStats, DEFAULT_OPTIONS } from './scanner.js';
import { newMailStats, MAIL_DEFAULT_OPTIONS } from '../mail/scanner.js';

export class ScanError extends Error {
  constructor(message, status = 400) {
    super(message);
    this.status = status;
  }
}

/** Normaliza as opções recebidas da interface. */
export function sanitizeOptions(input) {
  if (input === null || typeof input !== 'object' || Array.isArray(input)) input = {};
  const o = { ...DEFAULT_OPTIONS };
  if (input.checkName !== undefined) o.checkName = Boolean(input.checkName);
  if (input.checkContent !== undefined) o.checkContent = Boolean(input.checkContent);
  if (input.nameTarget === 'path' || input.nameTarget === 'file') o.nameTarget = input.nameTarget;
  if (input.resolveOwner !== undefined) o.resolveOwner = Boolean(input.resolveOwner);
  const size = Number(input.maxFileSizeMB);
  if (Number.isFinite(size) && size > 0) o.maxFileSizeMB = Math.min(size, 2048);
  const concurrency = Number(input.concurrency);
  if (Number.isInteger(concurrency) && concurrency > 0) o.concurrency = Math.min(concurrency, 16);
  if (input.modifiedAfter) {
    const d = new Date(input.modifiedAfter);
    if (Number.isNaN(d.getTime())) throw new ScanError('Data "modificados a partir de" inválida.');
    o.modifiedAfter = d.toISOString();
  }
  if (!o.checkName && !o.checkContent) throw new ScanError('Selecione ao menos uma verificação: nome ou conteúdo.');
  return o;
}

const MAIL_CHECKS = ['checkSubject', 'checkBody', 'checkAttachmentNames', 'checkAttachments', 'checkAddresses'];

/** Normaliza as opções de uma análise de e-mail. */
export function sanitizeMailOptions(input) {
  if (input === null || typeof input !== 'object' || Array.isArray(input)) input = {};
  const o = { ...MAIL_DEFAULT_OPTIONS };
  for (const k of [...MAIL_CHECKS, 'includeTrash', 'includeJunk']) if (input[k] !== undefined) o[k] = Boolean(input[k]);
  const size = Number(input.maxMessageSizeMB);
  if (Number.isFinite(size) && size > 0) o.maxMessageSizeMB = Math.min(size, 500);
  const concurrency = Number(input.concurrency);
  if (Number.isInteger(concurrency) && concurrency > 0) o.concurrency = Math.min(concurrency, 8);
  if (input.receivedAfter) {
    const d = new Date(input.receivedAfter);
    if (Number.isNaN(d.getTime())) throw new ScanError('Data "recebidas a partir de" inválida.');
    o.receivedAfter = d.toISOString();
  }
  if (!MAIL_CHECKS.some((k) => o[k])) throw new ScanError('Selecione ao menos uma verificação: assunto, corpo ou anexos.');
  return o;
}

/** Configuração da conexão usada pela análise (sem os segredos). */
function mailSnapshot(source) {
  const { id, name, type, scope, mailboxes, excludeMailboxes, excludeFolders, graph, gmail, imap } = source;
  return { id, name, type, scope, mailboxes, excludeMailboxes, excludeFolders, graph, gmail, imap };
}

const ids = (value) => (Array.isArray(value) ? [...new Set(value.filter((v) => typeof v === 'string'))] : []);

export class ScanManager {
  /** mailEndpoints: endereços alternativos das APIs de e-mail (usado nos testes). */
  constructor(store, { maxConcurrent = 1, workerUrl = new URL('./worker.js', import.meta.url), mailEndpoints = {} } = {}) {
    this.store = store;
    this.mailEndpoints = mailEndpoints;
    this.maxConcurrent = Math.max(1, maxConcurrent);
    this.workerUrl = workerUrl;
    this.running = new Map();
    this.queue = [];
  }

  /** Termos das listas escolhidas (o identificador do termo inclui o da lista). */
  #terms(listIds) {
    const lists = ids(listIds).map((id) => this.store.getList(id));
    if (lists.length === 0 || lists.some((l) => !l)) throw new ScanError('Selecione listas de referência válidas.');
    const terms = lists.flatMap((list) =>
      (list.terms || []).map((term) => ({ ...term, id: `${list.id}:${term.id}`, listName: list.name })),
    );
    if (terms.length === 0) throw new ScanError('As listas selecionadas não possuem termos.');
    return { lists, terms };
  }

  #scanName(name, prefix) {
    const stamp = new Date().toLocaleString('pt-BR', { dateStyle: 'short', timeStyle: 'short' });
    return String(name || '').trim().slice(0, 200) || `${prefix} de ${stamp}`;
  }

  /** Inicia uma análise. body.kind = 'mail' para caixas de e-mail; senão, repositórios de arquivos. */
  async start(body = {}) {
    if (body?.kind === 'mail') return this.#startMail(body);
    const { name, repositoryIds, listIds, options } = body || {};
    const repositories = ids(repositoryIds).map((id) => this.store.getRepository(id));
    if (repositories.length === 0 || repositories.some((r) => !r)) throw new ScanError('Selecione repositórios válidos.');
    const { lists, terms } = this.#terms(listIds);
    const opts = sanitizeOptions(options);
    const scan = this.store.createScan({
      kind: 'files',
      name: this.#scanName(name, 'Análise'),
      status: 'queued',
      repositoryIds: repositories.map((r) => r.id),
      listIds: lists.map((l) => l.id),
      options: opts,
      summary: {
        repositories: repositories.map((r) => ({ id: r.id, name: r.name, path: r.path })),
        lists: lists.map((l) => ({ id: l.id, name: l.name, termCount: (l.terms || []).length })),
        termCount: terms.length,
      },
      stats: newStats(repositories.length),
      current: null,
      startedAt: null,
      finishedAt: null,
      error: null,
    });
    await this.store.writeScanConfig(scan.id, {
      repositories: repositories.map(({ id, name, path, exclude, audit }) => ({ id, name, path, exclude, audit })),
      terms,
      options: opts,
    });
    this.queue.push(scan.id);
    this.#pump();
    return scan;
  }

  async #startMail({ name, sourceIds, listIds, options }) {
    const sources = ids(sourceIds).map((id) => this.store.getMailSource(id));
    if (sources.length === 0 || sources.some((s) => !s)) throw new ScanError('Selecione conexões de e-mail válidas.');
    const { lists, terms } = this.#terms(listIds);
    const opts = sanitizeMailOptions(options);
    const scan = this.store.createScan({
      kind: 'mail',
      name: this.#scanName(name, 'Análise de e-mail'),
      status: 'queued',
      sourceIds: sources.map((s) => s.id),
      listIds: lists.map((l) => l.id),
      options: opts,
      summary: {
        sources: sources.map((s) => ({ id: s.id, name: s.name, type: s.type, scope: s.scope, mailboxCount: s.scope === 'all' ? null : (s.mailboxes || []).length })),
        lists: lists.map((l) => ({ id: l.id, name: l.name, termCount: (l.terms || []).length })),
        termCount: terms.length,
      },
      stats: newMailStats(sources.length),
      current: null,
      startedAt: null,
      finishedAt: null,
      error: null,
    });
    await this.store.writeScanConfig(scan.id, { kind: 'mail', sources: sources.map(mailSnapshot), terms, options: opts });
    this.queue.push(scan.id);
    this.#pump();
    return scan;
  }

  /**
   * Configuração entregue à thread. Nas análises de e-mail, os segredos são decifrados só agora e
   * vão apenas para a memória da thread (o config.json da análise não os contém).
   */
  async #workerConfig(id) {
    const config = await this.store.readScanConfig(id);
    if (config.kind !== 'mail') return config;
    const sources = config.sources.map((s) => {
      const current = this.store.getMailSource(s.id);
      if (!current) throw new ScanError(`A conexão de e-mail "${s.name}" foi excluída antes do início da análise.`);
      return { ...mailSnapshot(current), secrets: this.store.openMailSecrets(current) };
    });
    return { ...config, sources, endpoints: this.mailEndpoints };
  }

  #pump() {
    while (this.running.size < this.maxConcurrent && this.queue.length > 0) {
      const id = this.queue.shift();
      // A vaga é reservada antes de qualquer await, para respeitar o limite de análises simultâneas.
      const entry = { worker: null, done: false, cancelRequested: false, fatal: null };
      this.running.set(id, entry);
      this.#run(id, entry).catch((err) => {
        this.running.delete(id);
        this.#fail(id, err.message);
        this.#pump();
      });
    }
  }

  #fail(id, message) {
    this.store.updateScan(id, { status: 'failed', error: message, finishedAt: new Date().toISOString(), current: null });
    this.store.appendLog(id, { level: 'error', message });
  }

  async #run(id, entry) {
    const config = await this.#workerConfig(id);
    if (entry.cancelRequested) {
      this.running.delete(id);
      this.store.updateScan(id, { status: 'cancelled', finishedAt: new Date().toISOString(), current: null });
      this.#pump();
      return;
    }
    this.store.updateScan(id, { status: 'running', startedAt: new Date().toISOString() });
    const worker = new Worker(this.workerUrl, { workerData: config });
    entry.worker = worker;
    worker.on('message', (message) => this.#onMessage(id, entry, message));
    worker.on('error', (err) => {
      entry.fatal = err?.stack || String(err);
    });
    worker.on('exit', async (code) => {
      this.running.delete(id);
      if (!entry.done) {
        await this.store.flushScan(id);
        const scan = this.store.getScan(id);
        if (scan) {
          if (entry.cancelRequested) {
            this.store.updateScan(id, { status: 'cancelled', finishedAt: new Date().toISOString(), current: null });
          } else {
            const label = scan.kind === 'mail' ? 'Última mensagem em leitura' : 'Último arquivo em leitura';
            const where = scan.current?.path ? ` ${label}: ${scan.current.path}` : '';
            const reason = entry.fatal ? `Falha na análise: ${entry.fatal.split('\n')[0]}` : `A análise terminou inesperadamente (código ${code}).`;
            this.#fail(id, `${reason}${where}`);
          }
        }
      }
      this.#pump();
    });
  }

  #onMessage(id, entry, message) {
    const { store } = this;
    switch (message?.type) {
      case 'progress':
        store.updateScan(id, { stats: message.stats, current: message.current });
        break;
      case 'results':
        store.appendResults(id, message.records);
        break;
      case 'errors':
        store.appendErrors(id, message.items);
        break;
      case 'log':
        store.appendLog(id, message);
        break;
      case 'fatal':
        entry.fatal = message.message;
        break;
      case 'done':
        entry.done = true;
        store.flushScan(id).then(() => {
          store.updateScan(id, {
            status: message.cancelled ? 'cancelled' : 'completed',
            finishedAt: new Date().toISOString(),
            stats: message.stats,
            current: null,
          });
          store.saveNow().catch(() => {});
        });
        // A thread termina sozinha; se algo a mantiver viva, encerra depois de alguns segundos.
        setTimeout(() => entry.worker.terminate(), 5000).unref();
        break;
      default:
        break;
    }
  }

  isActive(id) {
    return this.running.has(id) || this.queue.includes(id);
  }

  cancel(id) {
    const queued = this.queue.indexOf(id);
    if (queued !== -1) {
      this.queue.splice(queued, 1);
      this.store.updateScan(id, { status: 'cancelled', finishedAt: new Date().toISOString() });
      return true;
    }
    const entry = this.running.get(id);
    if (!entry) return false;
    entry.cancelRequested = true;
    if (!entry.worker) return true; // ainda preparando: #run encerra antes de criar a thread
    entry.worker.postMessage({ type: 'cancel' });
    // Se a thread não terminar sozinha (ex.: arquivo enorme sendo lido), força o encerramento.
    setTimeout(() => {
      if (!entry.done) entry.worker.terminate();
    }, 60000).unref();
    return true;
  }

  async shutdown() {
    this.queue.length = 0;
    const exits = [];
    for (const [id, entry] of this.running) {
      entry.cancelRequested = true;
      this.store.updateScan(id, { status: 'interrupted', finishedAt: new Date().toISOString(), current: null });
      entry.done = true;
      if (entry.worker) {
        exits.push(new Promise((resolve) => entry.worker.once('exit', resolve)));
        entry.worker.terminate();
      }
    }
    await Promise.all(exits);
    await this.store.close();
  }
}
