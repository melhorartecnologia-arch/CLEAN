// Fila de análises: cada análise roda em uma worker thread; o resultado é gravado conforme chega.
import { Worker } from 'node:worker_threads';
import { newStats, DEFAULT_OPTIONS } from './scanner.js';
import { newMailStats, MAIL_DEFAULT_OPTIONS } from '../mail/scanner.js';
import { cleanPaths, keptPaths, isCloudRepo, deletionScope, mailDeletionScope } from './delete.js';
import { keptCloud } from '../cloud/drives.js';
import { PROJECT_ROOT } from '../config.js';

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
  o.deleteMatches = input.deleteMatches === true; // só com o valor exato: nunca por engano
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
  o.deleteMatches = input.deleteMatches === true; // só com o valor exato: nunca por engano
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

/** Repositório usado pela análise (sem os segredos). */
function repoSnapshot(repo) {
  const { id, name, path, exclude, audit, allowDelete } = repo;
  if (!isCloudRepo(repo)) return { id, type: 'local', name, path, exclude, audit, allowDelete: Boolean(allowDelete) };
  const { type, graph, cloud, deleteMode } = repo;
  return { id, type, name, path, exclude, allowDelete: Boolean(allowDelete), deleteMode: deleteMode === 'permanent' ? 'permanent' : 'trash', graph, cloud };
}

/** Configuração da conexão usada pela análise (sem os segredos). */
function mailSnapshot(source) {
  const { id, name, type, scope, mailboxes, excludeMailboxes, excludeFolders, graph, gmail, imap, allowDelete, deleteMode } = source;
  return { id, name, type, scope, mailboxes, excludeMailboxes, excludeFolders, graph, gmail, imap, allowDelete: Boolean(allowDelete), deleteMode: deleteMode === 'trash' ? 'trash' : 'permanent' };
}

/**
 * "Analisar e excluir": só com a exclusão permitida em todos os locais escolhidos e com a
 * confirmação digitada na tela ("EXCLUIR").
 */
function checkDeletion(opts, body, targets, noun) {
  if (!opts.deleteMatches) return;
  const blocked = targets.filter((t) => !t.allowDelete).map((t) => `"${t.name}"`);
  if (blocked.length) throw new ScanError(`A exclusão não está permitida em ${blocked.join(', ')}. Ative "Permitir exclusão" no cadastro ${noun} ou escolha "Somente analisar".`);
  if (String(body?.confirmDelete || '').trim().toUpperCase() !== 'EXCLUIR') throw new ScanError('Para analisar e excluir, digite EXCLUIR na confirmação.');
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
    // Exclusões manuais em andamento pelo relatório ("análise:item"): o relatório não pode ser
    // excluído (nem pela limpeza dos agendamentos) enquanto houver uma.
    this.itemDeletions = new Set();
    // Conferência das análises agendadas com exclusão (definida pelo agendador): devolve o motivo
    // para não excluir, ou null.
    this.deletionGuard = null;
  }

  hasItemDeletion(scanId) {
    for (const key of this.itemDeletions) if (key.startsWith(`${scanId}:`)) return true;
    return false;
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

  /**
   * Inicia uma análise. body.kind = 'mail' para caixas de e-mail; senão, repositórios de arquivos.
   * by: quem iniciou (registrado nas exclusões automáticas); schedule: { id, name } do agendamento
   * que iniciou a análise.
   */
  async start(body = {}, { by = null, schedule = null } = {}) {
    const origin = { startedBy: by, ...(schedule ? { scheduleId: schedule.id, scheduleName: schedule.name } : {}) };
    if (body?.kind === 'mail') return this.#startMail(body, origin);
    const { name, repositoryIds, listIds, options } = body || {};
    const repositories = ids(repositoryIds).map((id) => this.store.getRepository(id));
    if (repositories.length === 0 || repositories.some((r) => !r)) throw new ScanError('Selecione repositórios válidos.');
    const { lists, terms } = this.#terms(listIds);
    const opts = sanitizeOptions(options);
    checkDeletion(opts, body, repositories, 'do repositório');
    const scan = this.store.createScan({
      kind: 'files',
      name: this.#scanName(name, 'Análise'),
      status: 'queued',
      repositoryIds: repositories.map((r) => r.id),
      listIds: lists.map((l) => l.id),
      options: opts,
      ...origin,
      summary: {
        repositories: repositories.map((r) => ({ id: r.id, name: r.name, path: r.path, type: r.type || 'local' })),
        lists: lists.map((l) => ({ id: l.id, name: l.name, termCount: (l.terms || []).length })),
        termCount: terms.length,
      },
      stats: newStats(repositories.length),
      current: null,
      startedAt: null,
      finishedAt: null,
      error: null,
    });
    await this.store.writeScanConfig(scan.id, { repositories: repositories.map(repoSnapshot), terms, options: opts });
    this.queue.push(scan.id);
    this.#pump();
    return scan;
  }

  async #startMail(body, origin) {
    const { name, sourceIds, listIds, options } = body;
    const sources = ids(sourceIds).map((id) => this.store.getMailSource(id));
    if (sources.length === 0 || sources.some((s) => !s)) throw new ScanError('Selecione conexões de e-mail válidas.');
    const { lists, terms } = this.#terms(listIds);
    const opts = sanitizeMailOptions(options);
    checkDeletion(opts, body, sources, 'da conexão de e-mail');
    const scan = this.store.createScan({
      kind: 'mail',
      name: this.#scanName(name, 'Análise de e-mail'),
      status: 'queued',
      sourceIds: sources.map((s) => s.id),
      listIds: lists.map((l) => l.id),
      options: opts,
      ...origin,
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
   * A exclusão automática é conferida de novo com o cadastro atual: só continua se ainda for
   * permitida (no mesmo caminho) e, no e-mail, do modo confirmado ao criar a análise (ou para a
   * lixeira, se o cadastro passou a ser assim). Público para os testes.
   */
  async workerConfig(id) {
    const config = await this.store.readScanConfig(id);
    const scan = this.store.getScan(id);
    let deleting = Boolean(config.options?.deleteMatches);
    const warn = (message) => this.store.appendLog(id, { level: 'warn', message });
    // Execução agendada que esperou na fila: só exclui se o agendamento ainda autorizar.
    const blocked = deleting && scan?.scheduleId && this.deletionGuard ? this.deletionGuard(scan) : null;
    if (blocked) {
      deleting = false;
      warn(`Exclusão automática desativada nesta execução: ${blocked}. A análise continua sem excluir.`);
      this.store.updateScan(id, { options: { ...scan.options, deleteMatches: false } });
    }
    const extra = { startedBy: scan?.startedBy || null, protect: cleanPaths({ dataDir: this.store.dataDir, appDir: PROJECT_ROOT }) };
    if (config.kind !== 'mail') {
      const all = this.store.listRepositories();
      const repositories = config.repositories.map((r) => {
        const current = this.store.getRepository(r.id);
        let repo = { ...r, allowDelete: false };
        if (isCloudRepo(r)) {
          // OneDrive/SharePoint: as credenciais (e o que analisar) vêm do cadastro atual.
          if (!current || !isCloudRepo(current)) throw new ScanError(`O repositório "${r.name}" foi excluído ou alterado antes do início da análise.`);
          repo = { ...repoSnapshot(current), allowDelete: false, deleteMode: r.deleteMode, secrets: this.store.openRepositorySecrets(current) };
        }
        if (!deleting || !r.allowDelete) return repo;
        const reason = !current
          ? 'o repositório foi removido do cadastro'
          : deletionScope(current) !== deletionScope(r)
            ? isCloudRepo(r)
              ? 'as contas, os sites ou as credenciais do repositório foram alterados'
              : 'o caminho do repositório foi alterado'
            : !current.allowDelete
              ? 'a opção "Permitir exclusão" foi desligada'
              : null;
        if (reason) {
          warn(`Exclusão automática desativada para "${r.name}": ${reason} depois que a análise foi criada.`);
          return repo;
        }
        if (!isCloudRepo(r)) return { ...repo, allowDelete: true, keep: keptPaths(current, all) };
        // Vale a forma confirmada ao criar a análise (ou a lixeira, se o cadastro passou a ser assim).
        const deleteMode = r.deleteMode === 'trash' || current.deleteMode !== 'permanent' ? 'trash' : 'permanent';
        return { ...repo, allowDelete: true, deleteMode, keep: keptCloud(current, all) };
      });
      return { ...config, options: { ...config.options, deleteMatches: deleting }, repositories, endpoints: this.mailEndpoints, ...extra };
    }
    const sources = config.sources.map((s) => {
      const current = this.store.getMailSource(s.id);
      if (!current) throw new ScanError(`A conexão de e-mail "${s.name}" foi excluída antes do início da análise.`);
      const live = mailSnapshot(current);
      // Como nos repositórios: só exclui no mesmo alcance (conta, servidor e caixas) de quando a
      // análise foi criada.
      const reason = !live.allowDelete ? 'a opção "Permitir exclusão" foi desligada' : mailDeletionScope(live) !== mailDeletionScope(s) ? 'a conta, o servidor ou as caixas da conexão foram alterados' : null;
      const allowDelete = deleting && s.allowDelete && !reason;
      if (deleting && s.allowDelete && reason) warn(`Exclusão automática desativada para "${s.name}": ${reason} depois que a análise foi criada.`);
      const deleteMode = s.deleteMode === 'trash' || live.deleteMode === 'trash' ? 'trash' : 'permanent';
      return { ...live, allowDelete, deleteMode, secrets: this.store.openMailSecrets(current) };
    });
    return { ...config, options: { ...config.options, deleteMatches: deleting }, sources, endpoints: this.mailEndpoints, ...extra };
  }

  /**
   * A exclusão foi desligada (ou o cadastro mudou) durante uma análise: as threads em andamento
   * deixam de excluir itens daquele repositório (kind 'repository') ou conexão (kind 'mail').
   */
  revokeDeletion(kind, id, reason) {
    for (const entry of this.running.values()) entry.worker?.postMessage({ type: 'revoke-delete', kind, id, reason });
  }

  /**
   * Uma análise deixa de excluir em todos os locais (ex.: o agendamento que a iniciou foi pausado
   * ou excluído). Na fila, a conferência é feita quando ela começa.
   */
  revokeScanDeletion(id, reason) {
    const entry = this.running.get(id);
    if (!entry) return;
    entry.revokeAll = reason; // a thread ainda pode estar sendo criada
    entry.worker?.postMessage({ type: 'revoke-delete', all: true, reason });
  }

  #pump() {
    while (this.running.size < this.maxConcurrent && this.queue.length > 0) {
      const id = this.queue.shift();
      // A vaga é reservada antes de qualquer await, para respeitar o limite de análises simultâneas.
      const entry = { worker: null, done: false, cancelRequested: false, shuttingDown: false, fatal: null };
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
    const config = await this.workerConfig(id);
    if (entry.cancelRequested) {
      this.running.delete(id);
      this.store.updateScan(id, { status: 'cancelled', finishedAt: new Date().toISOString(), current: null });
      this.#pump();
      return;
    }
    this.store.updateScan(id, { status: 'running', startedAt: new Date().toISOString() });
    const worker = new Worker(this.workerUrl, { workerData: config });
    entry.worker = worker;
    if (entry.revokeAll) worker.postMessage({ type: 'revoke-delete', all: true, reason: entry.revokeAll });
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
            this.store.updateScan(id, { status: entry.shuttingDown ? 'interrupted' : 'cancelled', finishedAt: new Date().toISOString(), current: null });
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
      case 'deletions':
        store.appendDeletions(id, message.items).catch((err) => {
          console.error('[CLEAN] Falha ao gravar o registro de exclusões:', err.message);
          const items = message.items.map((d) => d.item).join(' | ');
          store.appendLog(id, { level: 'error', message: `Falha ao gravar o registro de ${message.items.length} exclusão(ões): ${err.message}. Itens: ${items.slice(0, 2000)}` });
        });
        break;
      case 'log':
        store.appendLog(id, message);
        break;
      case 'fatal':
        entry.fatal = message.message;
        break;
      case 'done':
        entry.done = true;
        entry.finishing = store.flushScan(id).then(() => {
          entry.finished = true; // resultados gravados: o relatório já pode ser usado (e excluir itens)
          store.updateScan(id, {
            status: entry.shuttingDown ? 'interrupted' : message.cancelled ? 'cancelled' : 'completed',
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

  /** Na fila ou em andamento (até os resultados estarem gravados; a thread pode demorar a sair). */
  isActive(id) {
    const entry = this.running.get(id);
    return Boolean(entry && !entry.finished) || this.queue.includes(id);
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

  /**
   * Encerramento do servidor: as análises são canceladas e têm alguns segundos para registrar o que
   * já fizeram (inclusive exclusões) antes de a thread ser encerrada.
   */
  async shutdown({ graceMs = 5000 } = {}) {
    this.queue.length = 0;
    const exits = [];
    const entries = [...this.running.values()];
    for (const [id, entry] of this.running) {
      entry.cancelRequested = true;
      entry.shuttingDown = true;
      this.store.updateScan(id, { status: 'interrupted', finishedAt: new Date().toISOString(), current: null });
      if (entry.worker) {
        exits.push(new Promise((resolve) => entry.worker.once('exit', resolve)));
        entry.worker.postMessage({ type: 'cancel' });
        setTimeout(() => entry.worker.terminate(), graceMs).unref();
      } else {
        entry.done = true;
      }
    }
    await Promise.all(exits);
    await Promise.all(entries.map((entry) => entry.finishing));
    await this.store.close();
  }
}
