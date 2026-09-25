// Motor da análise: percorre os repositórios, procura os termos no nome e no conteúdo e monta os
// registros do relatório (apenas arquivos com ocorrências) com o último usuário de cada arquivo.
import fs from 'node:fs/promises';
import path from 'node:path';
import { Matcher } from './matcher.js';
import { walk, compileExclusions, DEFAULT_EXCLUDES } from './walker.js';
import { extractFile } from './extractors/index.js';
import { OwnerResolver } from './owner.js';
import { AuditIndex, queryAuditEvents, pickLastUser } from './audit.js';

export const DEFAULT_OPTIONS = {
  checkName: true,
  nameTarget: 'file', // 'file' = nome do arquivo; 'path' = caminho relativo (inclui pastas)
  checkContent: true,
  maxFileSizeMB: 50,
  modifiedAfter: null, // ISO: analisa só arquivos modificados a partir desta data
  resolveOwner: true,
  concurrency: 4,
  maxSamples: 3,
};

const ERROR_MESSAGES = {
  EACCES: 'Acesso negado',
  EPERM: 'Acesso negado (permissão)',
  ENOENT: 'Não encontrado (pode ter sido removido durante a análise)',
  EBUSY: 'Arquivo em uso ou bloqueado',
  ENAMETOOLONG: 'Caminho muito longo',
  ELOOP: 'Laço de atalhos',
  EIO: 'Erro de leitura (E/S)',
  ETIMEDOUT: 'Tempo esgotado ao acessar a rede',
  EHOSTUNREACH: 'Servidor inacessível',
};

export function friendlyError(err) {
  const base = ERROR_MESSAGES[err?.code];
  return base ? `${base} (${err.code})` : String(err?.message || err);
}

function iso(date) {
  return date instanceof Date && !Number.isNaN(date.getTime()) && date.getTime() > 0 ? date.toISOString() : null;
}

function uncHost(p) {
  return /^\\\\([^\\]+)\\/.exec(String(p).replace(/\//g, '\\'))?.[1] || '';
}

export function newStats(repositoriesTotal = 0) {
  return {
    repositoriesTotal,
    repositoriesDone: 0,
    directories: 0,
    filesSeen: 0,
    filesSkippedByDate: 0,
    filesMatched: 0,
    occurrences: 0,
    contentAnalyzed: 0,
    contentPartial: 0,
    contentEncrypted: 0,
    contentUnsupported: 0,
    contentSkippedSize: 0,
    contentEmpty: 0,
    contentErrors: 0,
    bytesAnalyzed: 0,
    errors: 0,
  };
}

export class Scanner {
  /**
   * config: { repositories: [{ id, name, path, exclude, audit }], terms: [...], options: {...} }
   * emit(message): recebe { type: 'log'|'progress'|'results'|'errors'|'done', ... }
   */
  constructor(config, emit, { ownerResolver, auditQuery = queryAuditEvents } = {}) {
    this.repositories = config.repositories || [];
    this.options = { ...DEFAULT_OPTIONS, ...(config.options || {}) };
    this.matcher = new Matcher(config.terms || [], { maxSamples: this.options.maxSamples });
    this.emit = emit;
    this.ownerResolver = ownerResolver || new OwnerResolver();
    this.auditQuery = auditQuery;
    this.stats = newStats(this.repositories.length);
    this.cancelled = false;
    this.seq = 0;
    this.pendingOwners = [];
    this.pendingErrors = [];
    this.lastProgress = 0;
    this.current = null;
    this.auditCache = new Map();
    const maxBytes = Math.max(1, Number(this.options.maxFileSizeMB) || 50) * 1024 * 1024;
    this.limits = { maxBytes, maxChars: 20_000_000 };
    this.modifiedAfter = this.options.modifiedAfter ? new Date(this.options.modifiedAfter).getTime() : null;
    if (Number.isNaN(this.modifiedAfter)) this.modifiedAfter = null;
  }

  cancel() {
    this.cancelled = true;
  }

  log(level, message) {
    this.emit({ type: 'log', level, message, time: new Date().toISOString() });
  }

  error(filePath, err) {
    this.stats.errors++;
    this.pendingErrors.push({ path: filePath, message: typeof err === 'string' ? err : friendlyError(err), time: new Date().toISOString() });
    if (this.pendingErrors.length >= 200) this.flushErrors();
  }

  flushErrors() {
    if (this.pendingErrors.length === 0) return;
    this.emit({ type: 'errors', items: this.pendingErrors.splice(0) });
  }

  progress(force = false) {
    const now = Date.now();
    if (!force && now - this.lastProgress < 400) return;
    this.lastProgress = now;
    this.emit({ type: 'progress', stats: { ...this.stats }, current: this.current });
  }

  async run() {
    const started = Date.now();
    for (const { term, error } of this.matcher.invalid) this.log('warn', `Termo ignorado "${term.value}": ${error}`);
    if (this.matcher.size === 0) this.log('warn', 'Nenhum termo válido nas listas selecionadas.');
    this.log('info', `Análise iniciada com ${this.matcher.size} termo(s) em ${this.repositories.length} repositório(s).`);
    for (const repo of this.repositories) {
      if (this.cancelled) break;
      await this.scanRepository(repo);
      this.stats.repositoriesDone++;
      this.progress(true);
    }
    await this.flushOwners();
    this.flushErrors();
    this.current = null;
    const seconds = Math.round((Date.now() - started) / 1000);
    this.log('info', `${this.cancelled ? 'Análise cancelada' : 'Análise concluída'} em ${seconds}s: ${this.stats.filesSeen} arquivo(s) verificados, ${this.stats.filesMatched} com ocorrências.`);
    this.emit({ type: 'done', stats: { ...this.stats }, cancelled: this.cancelled });
    return this.stats;
  }

  async scanRepository(repo) {
    this.current = { repository: repo.name, path: repo.path };
    this.progress(true);
    try {
      const st = await fs.stat(repo.path);
      if (!st.isDirectory()) throw Object.assign(new Error('O caminho não é uma pasta'), { code: 'ENOTDIR' });
    } catch (err) {
      this.error(repo.path, err);
      this.log('error', `Repositório "${repo.name}" inacessível: ${friendlyError(err)}`);
      return;
    }
    this.log('info', `Analisando "${repo.name}" (${repo.path})`);
    const auditIndex = await this.loadAudit(repo);
    const isExcluded = compileExclusions([...DEFAULT_EXCLUDES, ...(repo.exclude || [])]);
    const iterator = walk(repo.path, { isExcluded, shouldStop: () => this.cancelled });
    const worker = async () => {
      for (;;) {
        if (this.cancelled) return;
        const { value: entry, done } = await iterator.next();
        if (done) return;
        if (entry.type === 'error') {
          this.error(entry.path, entry.error);
        } else if (entry.type === 'dir') {
          this.stats.directories++;
        } else {
          try {
            await this.processFile(repo, entry, auditIndex);
          } catch (err) {
            this.error(entry.path, err);
          }
        }
        this.progress();
      }
    };
    const n = Math.min(Math.max(1, Number(this.options.concurrency) || 4), 16);
    await Promise.all(Array.from({ length: n }, worker));
    await iterator.return?.();
  }

  async loadAudit(repo) {
    const cfg = repo.audit || {};
    if (!cfg.enabled) return null;
    const computer = String(cfg.computer || uncHost(repo.path) || '').trim();
    const days = Math.min(Math.max(Number(cfg.days) || 30, 1), 365);
    const maxEvents = Math.min(Math.max(Number(cfg.maxEvents) || 200000, 100), 5_000_000);
    const ignoreUsers = Array.isArray(cfg.ignoreUsers) ? cfg.ignoreUsers : [];
    const key = `${computer.toLowerCase()}|${days}|${maxEvents}|${ignoreUsers.join(';').toLowerCase()}`;
    if (!this.auditCache.has(key)) {
      this.log('info', `Consultando o log de auditoria de ${computer || 'este computador'} (últimos ${days} dias)...`);
      this.auditCache.set(
        key,
        this.auditQuery({ computer, days, maxEvents, ignoreUsers }).then(
          (events) => {
            const index = new AuditIndex(events);
            this.log('info', `Log de auditoria de ${computer || 'este computador'}: ${index.any.size} arquivo(s) com eventos.`);
            return index;
          },
          (err) => {
            this.log('warn', `Não foi possível ler o log de auditoria de ${computer || 'este computador'}: ${err.message}`);
            return null;
          },
        ),
      );
    }
    return this.auditCache.get(key);
  }

  async processFile(repo, entry, auditIndex) {
    this.stats.filesSeen++;
    this.current = { repository: repo.name, path: entry.path };
    let st;
    try {
      st = await fs.stat(entry.path);
    } catch (err) {
      this.error(entry.path, err);
      return;
    }
    if (this.modifiedAfter && st.mtimeMs < this.modifiedAfter) {
      this.stats.filesSkippedByDate++;
      return;
    }
    const { options, matcher } = this;
    const matches = [];
    if (options.checkName) {
      const label = options.nameTarget === 'path' ? 'Caminho' : 'Nome do arquivo';
      const text = options.nameTarget === 'path' ? entry.relativePath : entry.name;
      matches.push(...matcher.match([{ text, label }], 'name'));
    }
    let content = null;
    if (options.checkContent) {
      content = await extractFile(entry.path, { size: st.size, limits: this.limits });
      this.countContent(content, st.size);
      if (content.status === 'error') this.error(entry.path, content.note);
      matches.push(...matcher.match(content.segments, 'content'));
      content.segments = null;
    }
    if (matches.length === 0) return;

    if (!content || content.status === 'skipped-size') {
      // Somente os metadados (para o "salvo por último por"), sem ler o texto.
      const meta = await extractFile(entry.path, { size: st.size, limits: this.limits, withText: false }).catch(() => null);
      content = { ...(content || {}), type: content?.type || meta?.type, metadata: meta?.metadata || {} };
    }
    const occurrences = matches.reduce((sum, m) => sum + m.count, 0);
    this.stats.filesMatched++;
    this.stats.occurrences += occurrences;
    const audit = auditIndex ? auditIndex.lookup(repo.path, entry.relativePath, repo.audit?.localPath || '') : null;
    const record = {
      id: ++this.seq,
      repositoryId: repo.id,
      repositoryName: repo.name,
      path: entry.path,
      relativePath: entry.relativePath,
      name: entry.name,
      extension: path.extname(entry.name).toLowerCase(),
      size: st.size,
      created: iso(st.birthtime),
      modified: iso(st.mtime),
      accessed: iso(st.atime),
      contentType: content?.type || null,
      contentStatus: options.checkContent ? content?.status || null : 'not-requested',
      contentNote: content?.note || null,
      metadata: content?.metadata || {},
      audit,
      owner: null,
      ownerError: null,
      lastUser: null,
      lastUserSource: null,
      occurrences,
      terms: [...new Set(matches.map((m) => m.term))],
      matches,
    };
    if (options.resolveOwner) {
      this.pendingOwners.push(record);
      if (this.pendingOwners.length >= 200) await this.flushOwners();
    } else {
      this.finishRecords([record]);
    }
  }

  countContent(content, size) {
    const s = this.stats;
    switch (content.status) {
      case 'ok':
        s.contentAnalyzed++;
        s.bytesAnalyzed += size;
        break;
      case 'partial':
        s.contentAnalyzed++;
        s.contentPartial++;
        s.bytesAnalyzed += Math.min(size, this.limits.maxBytes);
        break;
      case 'encrypted':
        s.contentEncrypted++;
        break;
      case 'skipped-size':
        s.contentSkippedSize++;
        break;
      case 'empty':
        s.contentEmpty++;
        break;
      case 'error':
        s.contentErrors++;
        break;
      default:
        s.contentUnsupported++;
    }
  }

  async flushOwners() {
    while (this.pendingOwners.length > 0) {
      const batch = this.pendingOwners.splice(0, 400);
      const owners = await this.ownerResolver.resolve(batch.map((r) => r.path));
      for (const record of batch) {
        const info = owners.get(record.path);
        record.owner = info?.owner || null;
        record.ownerError = info?.error || null;
      }
      this.finishRecords(batch);
    }
  }

  finishRecords(records) {
    for (const record of records) {
      const pick = pickLastUser({ audit: record.audit, metadata: record.metadata, owner: record.owner });
      record.lastUser = pick.user;
      record.lastUserSource = pick.source;
    }
    this.emit({ type: 'results', records });
  }
}
