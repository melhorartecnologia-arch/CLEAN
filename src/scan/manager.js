// Fila de análises: cada análise roda em uma worker thread; o resultado é gravado conforme chega.
import { Worker } from 'node:worker_threads';
import { newStats, DEFAULT_OPTIONS } from './scanner.js';

export class ScanError extends Error {
  constructor(message, status = 400) {
    super(message);
    this.status = status;
  }
}

/** Normaliza as opções recebidas da interface. */
export function sanitizeOptions(input = {}) {
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

export class ScanManager {
  constructor(store, { maxConcurrent = 1, workerUrl = new URL('./worker.js', import.meta.url) } = {}) {
    this.store = store;
    this.maxConcurrent = Math.max(1, maxConcurrent);
    this.workerUrl = workerUrl;
    this.running = new Map();
    this.queue = [];
  }

  async start({ name, repositoryIds = [], listIds = [], options = {} }) {
    const repositories = [...new Set(repositoryIds)].map((id) => this.store.getRepository(id));
    const lists = [...new Set(listIds)].map((id) => this.store.getList(id));
    if (repositories.length === 0 || repositories.some((r) => !r)) throw new ScanError('Selecione repositórios válidos.');
    if (lists.length === 0 || lists.some((l) => !l)) throw new ScanError('Selecione listas de referência válidas.');
    const terms = lists.flatMap((list) =>
      (list.terms || []).map((term) => ({ ...term, id: `${list.id}:${term.id}`, listName: list.name })),
    );
    if (terms.length === 0) throw new ScanError('As listas selecionadas não possuem termos.');
    const opts = sanitizeOptions(options);
    const stamp = new Date().toLocaleString('pt-BR', { dateStyle: 'short', timeStyle: 'short' });
    const scan = this.store.createScan({
      name: String(name || '').trim().slice(0, 200) || `Análise de ${stamp}`,
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

  #pump() {
    while (this.running.size < this.maxConcurrent && this.queue.length > 0) {
      const id = this.queue.shift();
      this.#run(id).catch((err) => this.#fail(id, err.message));
    }
  }

  #fail(id, message) {
    this.store.updateScan(id, { status: 'failed', error: message, finishedAt: new Date().toISOString(), current: null });
    this.store.appendLog(id, { level: 'error', message });
  }

  async #run(id) {
    const config = await this.store.readScanConfig(id);
    this.store.updateScan(id, { status: 'running', startedAt: new Date().toISOString() });
    const worker = new Worker(this.workerUrl, { workerData: config });
    const entry = { worker, done: false, cancelRequested: false, fatal: null };
    this.running.set(id, entry);
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
            this.#fail(id, entry.fatal ? `Falha na análise: ${entry.fatal.split('\n')[0]}` : `A análise terminou inesperadamente (código ${code}).`);
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
    entry.worker.postMessage({ type: 'cancel' });
    // Se a thread estiver presa (ex.: arquivo muito grande), força o encerramento.
    setTimeout(() => {
      if (!entry.done) entry.worker.terminate();
    }, 15000).unref();
    return true;
  }

  async shutdown() {
    this.queue.length = 0;
    const exits = [];
    for (const [id, entry] of this.running) {
      entry.cancelRequested = true;
      exits.push(new Promise((resolve) => entry.worker.once('exit', resolve)));
      this.store.updateScan(id, { status: 'interrupted', finishedAt: new Date().toISOString(), current: null });
      entry.done = true;
      entry.worker.terminate();
    }
    await Promise.all(exits);
    await this.store.close();
  }
}
