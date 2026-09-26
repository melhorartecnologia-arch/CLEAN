// Motor da análise: percorre os repositórios, procura os termos no nome e no conteúdo e monta os
// registros do relatório (apenas arquivos com ocorrências) com o último usuário de cada arquivo.
import fs from 'node:fs/promises';
import path from 'node:path';
import vm from 'node:vm';
import { Matcher } from './matcher.js';
import { walk, compileExclusions, DEFAULT_EXCLUDES } from './walker.js';
import { extractFile, extractBuffer } from './extractors/index.js';
import { OwnerResolver } from './owner.js';
import { AuditIndex, queryAuditEvents, pickLastUser } from './audit.js';
import { friendlyError, withTimeout } from './errors.js';
import { deleteFile, deletionEvent, isWithin, isCloudRepo, guardFor } from './delete.js';
import { DrivesConnector, person, keptCloudTarget, cloudTarget } from '../cloud/drives.js';
import { fileDate, cloudDate, ageDays, patternMatcher } from '../retention/policy.js';

// Tempo máximo para ler um arquivo (o que passar disso é registrado como erro e a análise segue).
const FILE_TIMEOUT = 5 * 60 * 1000;

// Arquivos de texto maiores que o limite: só o início é baixado e analisado (os demais formatos
// precisam do arquivo inteiro e ficam só com o nome verificado).
const TEXT_EXTENSIONS = new Set(['.txt', '.csv', '.tsv', '.log', '.md', '.json', '.xml', '.html', '.htm', '.ini', '.cfg', '.conf', '.sql', '.yaml', '.yml']);

export { isCloudRepo };

/** Endereço web legível (sem os códigos %20 etc.), parte por parte. */
function readableUrl(url) {
  return String(url)
    .split('/')
    .map((part) => {
      try {
        return decodeURIComponent(part);
      } catch {
        return part;
      }
    })
    .join('/');
}

export { friendlyError };

export const DEFAULT_OPTIONS = {
  checkName: true,
  nameTarget: 'file', // 'file' = nome do arquivo; 'path' = caminho relativo (inclui pastas)
  checkContent: true,
  maxFileSizeMB: 50,
  modifiedAfter: null, // ISO: analisa só arquivos modificados a partir desta data
  resolveOwner: true,
  concurrency: 4,
  maxSamples: 3,
  deleteMatches: false, // exclui automaticamente os arquivos em que algum termo for encontrado
};

/**
 * Executa uma função com tempo limite. Uma expressão regular com retrocesso excessivo não pode
 * ser interrompida de outra forma dentro da mesma thread; o módulo vm consegue.
 */
export function createGuard(timeoutMs) {
  const holder = { fn: null };
  const context = vm.createContext(holder);
  const script = new vm.Script('fn()');
  return (fn) => {
    holder.fn = fn;
    try {
      script.runInContext(context, { timeout: timeoutMs });
    } finally {
      holder.fn = null;
    }
  };
}

function iso(date) {
  return date instanceof Date && !Number.isNaN(date.getTime()) && date.getTime() > 0 ? date.toISOString() : null;
}

/**
 * Data usada no filtro "a partir de": a mais recente entre a modificação e a criação (um arquivo
 * copiado para o repositório mantém a data de modificação original, mas a criação é a da cópia).
 * Assim, as análises incrementais dos agendamentos não deixam de ver arquivos copiados com uma
 * data de modificação antiga. Mudanças só de permissões ou atributos não contam.
 */
export function changedAt(st) {
  return Math.max(st.mtimeMs || 0, st.birthtimeMs || 0);
}

/** O mesmo para os arquivos do OneDrive/SharePoint; sem nenhuma data válida, conta como alterado. */
export function cloudChangedAt(item) {
  const time = Math.max(Date.parse(item.lastModifiedDateTime) || 0, Date.parse(item.createdDateTime) || 0);
  return time || Infinity;
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
    libraries: 0, // bibliotecas do OneDrive/SharePoint analisadas
    accountsSkipped: 0, // contas sem OneDrive
    gaps: 0, // repositórios, contas ou sites que não puderam ser lidos (a análise ficou incompleta)
    bytesExpired: 0, // retenção: tamanho dos arquivos expirados
    retentionUnknown: 0, // retenção: arquivos sem a data do critério (não expiram)
    deleteSkipped: 0, // retenção: expirados não excluídos por causa do limite da execução
    deleteProtected: 0, // retenção: expirados em locais protegidos (não são excluídos)
    deleted: 0, // excluídos na análise ("analisar e excluir")
    deleteMissing: 0, // já não existiam na hora da exclusão
    deleteChanged: 0, // alterados depois de analisados: mantidos
    deleteErrors: 0,
  };
}

/** Conta o resultado de uma exclusão nas estatísticas da análise. */
export function countDeletion(stats, status) {
  if (status === 'deleted') stats.deleted++;
  else if (status === 'missing') stats.deleteMissing++;
  else if (status === 'changed') stats.deleteChanged++;
  else stats.deleteErrors++;
}

export class Scanner {
  /**
   * config: { repositories: [{ id, name, path, exclude, audit }], terms: [...], options: {...} }
   * emit(message): recebe { type: 'log'|'progress'|'results'|'errors'|'done', ... }
   */
  constructor(config, emit, { ownerResolver, auditQuery = queryAuditEvents, cloudConnectorFactory } = {}) {
    this.repositories = config.repositories || [];
    this.repoById = new Map(this.repositories.map((r) => [r.id, r]));
    this.options = { ...DEFAULT_OPTIONS, ...(config.options || {}) };
    this.matcher = new Matcher(config.terms || [], { maxSamples: this.options.maxSamples, guard: createGuard(config.regexTimeoutMs || 30000) });
    this.abort = new AbortController();
    this.emit = emit;
    this.ownerResolver = ownerResolver || new OwnerResolver();
    this.auditQuery = auditQuery;
    this.stats = newStats(this.repositories.length);
    this.cancelled = false;
    this.seq = 0;
    this.pendingOwners = [];
    this.inFlight = new Set(); // arquivos sendo lidos no momento (para mensagens de erro)
    this.pendingErrors = [];
    this.lastProgress = 0;
    this.current = null;
    this.auditCache = new Map();
    const maxBytes = Math.max(1, Number(this.options.maxFileSizeMB) || 50) * 1024 * 1024;
    this.limits = { maxBytes, maxChars: 20_000_000 };
    this.fileTimeoutMs = config.fileTimeoutMs || FILE_TIMEOUT;
    this.modifiedAfter = this.options.modifiedAfter ? new Date(this.options.modifiedAfter).getTime() : null;
    if (Number.isNaN(this.modifiedAfter)) this.modifiedAfter = null;
    // Pastas do próprio CLEAN (dados e instalação): nunca têm arquivos excluídos.
    this.protect = config.protect || [];
    // Política de retenção: arquivos mais antigos que a data de corte, pelo critério escolhido
    // (sem termos: o conteúdo não é lido).
    const r = config.retention || null;
    this.retention = r ? { ...r, cutoffMs: Date.parse(r.cutoff), matchName: patternMatcher(r.patterns || []) } : null;
    // Limite de exclusões da execução: vagas em uso (exclusões feitas ou em andamento) e falhas.
    this.deleteAttempts = 0;
    this.deleteFailures = 0;
    this.limitLogged = false;
    // OneDrive e SharePoint: um conector por repositório (usado também na exclusão automática).
    this.endpoints = config.endpoints || {};
    this.cloudConnectorFactory = cloudConnectorFactory || ((repo, options) => new DrivesConnector(repo, options));
    this.cloudConnectors = new Map();
    this.cloudKept = new Map(); // proteção por outros repositórios, conferida uma vez por repositório
    this.deletedBy = config.startedBy || null; // quem iniciou a análise com exclusão automática
  }

  cancel() {
    this.cancelled = true;
    this.abort.abort();
  }

  /** A exclusão foi desligada no cadastro durante a análise: nada mais é excluído daquele repositório. */
  revokeDeletion({ kind, id, reason, all = false }) {
    if (all) {
      const active = this.repositories.filter((r) => r.allowDelete);
      for (const repo of active) repo.allowDelete = false;
      if (active.length && this.options.deleteMatches) this.log('warn', `Exclusão automática desativada: ${reason}. A análise continua sem excluir.`);
      return;
    }
    const repo = kind === 'repository' ? this.repoById.get(id) : null;
    if (!repo?.allowDelete) return;
    repo.allowDelete = false;
    if (this.options.deleteMatches) this.log('warn', `Exclusão automática desativada para "${repo.name}": ${reason || 'o cadastro do repositório foi alterado'}.`);
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
    if (this.retention) {
      const cutoff = new Date(this.retention.cutoffMs).toLocaleString('pt-BR', { dateStyle: 'short', timeStyle: 'short' });
      this.log('info', `Retenção iniciada em ${this.repositories.length} repositório(s): arquivos com a data do critério anterior a ${cutoff}.`);
    } else {
      for (const { term, error } of this.matcher.invalid) this.log('warn', `Termo ignorado "${term.value}": ${error}`);
      if (this.matcher.size === 0) this.log('warn', 'Nenhum termo válido nas listas selecionadas.');
      this.log('info', `Análise iniciada com ${this.matcher.size} termo(s) em ${this.repositories.length} repositório(s).`);
    }
    for (const repo of this.repositories) {
      if (this.cancelled) break;
      if (isCloudRepo(repo)) await this.scanCloudRepository(repo);
      else await this.scanRepository(repo);
      // Proprietários (e exclusões) de cada repositório logo ao fim dele, e não só no fim da análise.
      await this.flushOwners();
      this.stats.repositoriesDone++;
      this.progress(true);
    }
    await this.flushOwners();
    this.flushErrors();
    this.current = null;
    const seconds = Math.round((Date.now() - started) / 1000);
    const found = this.retention ? 'expirado(s)' : 'com ocorrências';
    if (this.stats.deleteProtected) {
      this.log('info', `${this.stats.deleteProtected} arquivo(s) expirado(s) em locais protegidos (repositórios sem "Permitir exclusão" dentro dos analisados, contas ou sites protegidos, pastas do CLEAN) não foram excluídos.`);
    }
    this.log('info', `${this.cancelled ? 'Análise cancelada' : 'Análise concluída'} em ${seconds}s: ${this.stats.filesSeen} arquivo(s) verificados, ${this.stats.filesMatched} ${found}.`);
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
      this.stats.gaps++;
      this.log('error', `Repositório "${repo.name}" inacessível: ${friendlyError(err)}`);
      return;
    }
    this.log('info', `Analisando "${repo.name}" (${repo.path})`);
    // Na retenção o log de auditoria não é carregado (o último usuário vem do proprietário).
    const auditIndex = this.retention ? null : await this.loadAudit(repo);
    const isExcluded = compileExclusions([...DEFAULT_EXCLUDES, ...(repo.exclude || [])]);
    // A pasta de dados do CLEAN (com os relatórios, que contêm os próprios termos) não é analisada.
    const dataDir = this.protect.find((p) => p.data)?.path;
    const skipDir = (dir) => {
      if (!dataDir || !isWithin(dataDir, dir)) return false;
      this.log('info', `Pasta de dados do CLEAN ignorada: ${dir}`);
      return true;
    };
    const iterator = walk(repo.path, { isExcluded, skipDir, shouldStop: () => this.cancelled });
    const worker = async () => {
      for (;;) {
        if (this.cancelled) return;
        const { value: entry, done } = await iterator.next();
        if (done) return;
        if (entry.type === 'error') {
          this.error(entry.path, entry.error);
          if (entry.path === repo.path) this.stats.gaps++; // a raiz do repositório não pôde ser listada
        } else if (entry.type === 'dir') {
          this.stats.directories++;
        } else {
          this.inFlight.add(entry.path);
          try {
            await this.processFile(repo, entry, auditIndex);
          } catch (err) {
            this.error(entry.path, err);
          } finally {
            this.inFlight.delete(entry.path);
          }
        }
        this.progress();
      }
    };
    const n = Math.min(Math.max(1, Number(this.options.concurrency) || 4), 16);
    await Promise.all(Array.from({ length: n }, worker));
    await iterator.return?.();
  }

  // ---------------------------------------------------------------------------------------------
  // OneDrive e SharePoint

  cloudConnector(repo) {
    if (!this.cloudConnectors.has(repo.id)) {
      const log = (level, message) => this.log(level, `${repo.name}: ${message}`);
      this.cloudConnectors.set(repo.id, this.cloudConnectorFactory(repo, { signal: this.abort.signal, endpoints: this.endpoints, log }));
    }
    return this.cloudConnectors.get(repo.id);
  }

  async scanCloudRepository(repo) {
    const kind = repo.type === 'sharepoint' ? 'SharePoint' : 'OneDrive';
    this.current = { repository: repo.name, path: `${kind}: ${repo.name}` };
    this.progress(true);
    this.log('info', `Analisando "${repo.name}" (${kind})`);
    const connector = this.cloudConnector(repo);
    const isExcluded = compileExclusions([...DEFAULT_EXCLUDES, ...(repo.exclude || [])]);
    let skipped = 0;
    try {
      for await (const drive of connector.drives()) {
        if (this.cancelled) break;
        if (drive.skip) {
          this.stats.accountsSkipped++;
          if (++skipped <= 20) this.log('info', `${drive.account}: ignorado, ${drive.reason}.`);
          continue;
        }
        if (drive.error) {
          this.error(drive.account, drive.error);
          this.stats.gaps++;
          this.log('warn', `${drive.account} inacessível: ${friendlyError(drive.error)}`);
          continue;
        }
        // Biblioteca ignorada pelo nome (ex.: "Site Assets") ou pelo caminho ("Documentos/Antigo").
        if (isExcluded(drive.library, drive.library)) continue;
        await this.scanDrive(repo, connector, drive, isExcluded);
        this.stats.libraries++;
        this.progress(true);
      }
    } catch (err) {
      if (this.cancelled) return;
      this.error(repo.name, err);
      this.stats.gaps++;
      this.log('error', `Repositório "${repo.name}" inacessível: ${friendlyError(err)}`);
    }
    if (skipped > 20) this.log('info', `${skipped} conta(s) sem OneDrive foram ignoradas.`);
  }

  async scanDrive(repo, connector, drive, isExcluded) {
    this.current = { repository: repo.name, path: drive.label };
    // Os padrões valem para o caminho dentro da biblioteca e também com o nome dela na frente.
    const excluded = (name, rel) => isExcluded(name, rel) || isExcluded(name, `${drive.library}/${rel}`);
    const iterator = connector.walk(drive, { isExcluded: excluded, shouldStop: () => this.cancelled });
    const worker = async () => {
      for (;;) {
        if (this.cancelled) return;
        const { value: entry, done } = await iterator.next();
        if (done) return;
        if (entry.type === 'error') {
          this.error(entry.path, entry.error);
          if (entry.path === drive.label) this.stats.gaps++; // a raiz da biblioteca não pôde ser listada
        } else if (entry.type === 'dir') {
          this.stats.directories++;
        } else {
          const key = `${drive.label} › ${entry.relativePath}`;
          this.inFlight.add(key);
          try {
            await this.processCloudFile(repo, connector, drive, entry);
          } catch (err) {
            this.error(key, err);
          } finally {
            this.inFlight.delete(key);
          }
        }
        this.progress();
      }
    };
    const n = Math.min(Math.max(1, Number(this.options.concurrency) || 4), 16);
    try {
      await Promise.all(Array.from({ length: n }, worker));
    } finally {
      await iterator.return?.();
    }
  }

  /** Conteúdo de um arquivo da nuvem: baixado para a memória (até o limite de tamanho) e lido. */
  async cloudContent(connector, drive, item, signal) {
    const size = Number(item.size) || 0;
    const ext = path.extname(item.name).toLowerCase();
    const type = ext.slice(1) || 'arquivo';
    if (size === 0) return { type, status: 'empty', segments: [], metadata: {} };
    const tooBig = size > this.limits.maxBytes;
    const mb = Math.round(this.limits.maxBytes / 1048576);
    if (tooBig && !TEXT_EXTENSIONS.has(ext)) {
      return { type, status: 'skipped-size', segments: [], metadata: {}, note: `Conteúdo não analisado: arquivo maior que ${mb} MB.` };
    }
    const res = await connector.download(drive.id, item.id, { maxBytes: this.limits.maxBytes, signal });
    const content = await extractBuffer(res.data, { name: item.name, limits: this.limits });
    if ((tooBig || res.truncated) && content.status === 'ok') {
      content.status = 'partial';
      content.note = `Arquivo com mais de ${mb} MB: apenas o início foi analisado.`;
    }
    return content;
  }

  async processCloudFile(repo, connector, drive, entry) {
    this.stats.filesSeen++;
    const { item, relativePath } = entry;
    const label = `${drive.label} › ${relativePath}`;
    this.current = { repository: repo.name, path: label };
    if (this.modifiedAfter && cloudChangedAt(item) < this.modifiedAfter) {
      this.stats.filesSkippedByDate++;
      return;
    }
    if (this.retention) return this.processExpiredCloudFile(repo, drive, entry);
    const size = Number(item.size) || 0;
    const { options, matcher } = this;
    let content = null;
    if (options.checkContent) {
      // Cada download tem o seu sinal: no tempo esgotado (ou ao cancelar), ele é interrompido de fato.
      const controller = new AbortController();
      const job = this.cloudContent(connector, drive, item, AbortSignal.any([this.abort.signal, controller.signal]));
      content = await withTimeout(job, this.fileTimeoutMs, 'Tempo esgotado ao baixar ou ler o conteúdo do arquivo.').catch((err) => {
        controller.abort(err);
        return { type: path.extname(item.name).slice(1), status: 'error', segments: [], metadata: {}, note: `Falha ao ler o conteúdo: ${friendlyError(err)}` };
      });
      this.countContent(content, size);
      if (content.status === 'error') this.error(label, content.note);
    }
    const groups = [];
    if (options.checkName) {
      const nameLabel = options.nameTarget === 'path' ? 'Caminho' : 'Nome do arquivo';
      groups.push({ segments: [{ text: options.nameTarget === 'path' ? relativePath : item.name, label: nameLabel }], location: 'name' });
    }
    if (content) groups.push({ segments: content.segments, location: 'content' });
    const matches = matcher.matchGroups(groups).flat();
    if (matcher.timedOut) this.error(label, 'Tempo limite ao procurar as expressões regulares neste arquivo (possível retrocesso excessivo); resultado parcial.');
    if (content) content.segments = null;
    if (matches.length === 0) return;

    const occurrences = matches.reduce((sum, m) => sum + m.count, 0);
    this.stats.filesMatched++;
    this.stats.occurrences += occurrences;
    const record = this.cloudRecord(repo, drive, entry, {
      contentType: content?.type || null,
      contentStatus: options.checkContent ? content?.status || null : 'not-requested',
      contentNote: content?.note || null,
      metadata: content?.metadata || {},
      occurrences,
      terms: [...new Set(matches.map((m) => m.term))],
      matches,
    });
    this.finishRecords([record]);
    await this.deleteRecords([record]);
  }

  /** Registro de um arquivo do OneDrive/SharePoint no relatório. */
  cloudRecord(repo, drive, { item, relativePath }, fields) {
    return {
      id: ++this.seq,
      repositoryId: repo.id,
      repositoryName: repo.name,
      path: item.webUrl ? readableUrl(item.webUrl) : `${drive.label} › ${relativePath}`,
      relativePath,
      name: item.name,
      extension: path.extname(item.name).toLowerCase(),
      size: Number(item.size) || 0,
      created: iso(new Date(item.createdDateTime)),
      modified: iso(new Date(item.lastModifiedDateTime)),
      accessed: null,
      audit: null,
      // OneDrive: o dono da conta (no SharePoint, quem criou o arquivo fica em cloud.createdBy).
      owner: drive.owner || null,
      ownerError: null,
      lastUser: null,
      lastUserSource: null,
      cloud: {
        kind: drive.kind,
        tenant: repo.graph?.tenantId || null,
        driveId: drive.id,
        itemId: item.id,
        cTag: item.cTag || null,
        eTag: item.eTag || null,
        parentId: item.parentReference?.id || null,
        webUrl: item.webUrl || null,
        account: drive.account,
        accountName: drive.accountName || '',
        aliases: drive.aliases || [],
        library: drive.library,
        lastModifiedBy: person(item.lastModifiedBy),
        createdBy: person(item.createdBy),
      },
      ...fields,
    };
  }

  /** Campos de um item expirado (retenção), sem termos. */
  expiredFields(when) {
    return {
      contentType: null,
      contentStatus: 'not-requested',
      contentNote: null,
      metadata: {},
      occurrences: 0,
      terms: [],
      matches: [],
      retention: { criterion: this.retention.criterion, date: iso(new Date(when)), ageDays: ageDays(when) },
    };
  }

  /** Retenção: arquivo de pasta do Windows expirado (listado e, com exclusão, excluído). */
  async processExpiredFile(repo, entry, st) {
    const r = this.retention;
    if (!r.matchName(entry.name)) return;
    const when = fileDate(st, r.criterion);
    if (when === null) {
      this.stats.retentionUnknown++; // sem a data do critério: nunca é excluído
      return;
    }
    if (when >= r.cutoffMs) return;
    this.stats.filesMatched++;
    this.stats.bytesExpired += st.size;
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
      audit: null,
      owner: null,
      ownerError: null,
      lastUser: null,
      lastUserSource: null,
      ...this.expiredFields(when),
    };
    if (this.options.resolveOwner) {
      this.pendingOwners.push(record);
      if (this.pendingOwners.length >= 200) await this.flushOwners();
    } else {
      this.finishRecords([record]);
      await this.deleteRecords([record]);
    }
  }

  /** Retenção: arquivo do OneDrive/SharePoint expirado (sem baixar o conteúdo). */
  async processExpiredCloudFile(repo, drive, entry) {
    const r = this.retention;
    const { item } = entry;
    if (!r.matchName(item.name)) return;
    const when = cloudDate(item, r.criterion);
    if (when === null) {
      this.stats.retentionUnknown++;
      return;
    }
    if (when >= r.cutoffMs) return;
    this.stats.filesMatched++;
    this.stats.bytesExpired += Number(item.size) || 0;
    const record = this.cloudRecord(repo, drive, entry, this.expiredFields(when));
    this.finishRecords([record]);
    await this.deleteRecords([record]);
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
    if (this.modifiedAfter && changedAt(st) < this.modifiedAfter) {
      this.stats.filesSkippedByDate++;
      return;
    }
    if (this.retention) return this.processExpiredFile(repo, entry, st);
    const { options, matcher } = this;
    let content = null;
    if (options.checkContent) {
      content = await withTimeout(
        extractFile(entry.path, { size: st.size, limits: this.limits }),
        this.fileTimeoutMs,
        'Tempo esgotado ao ler o conteúdo do arquivo.',
      ).catch((err) => ({ type: path.extname(entry.name).slice(1), status: 'error', segments: [], metadata: {}, note: friendlyError(err) }));
      this.countContent(content, st.size);
      if (content.status === 'error') this.error(entry.path, content.note);
    }
    const groups = [];
    if (options.checkName) {
      const label = options.nameTarget === 'path' ? 'Caminho' : 'Nome do arquivo';
      const text = options.nameTarget === 'path' ? entry.relativePath : entry.name;
      groups.push({ segments: [{ text, label }], location: 'name' });
    }
    if (content) groups.push({ segments: content.segments, location: 'content' });
    const matches = matcher.matchGroups(groups).flat();
    if (matcher.timedOut) {
      this.error(entry.path, 'Tempo limite ao procurar as expressões regulares neste arquivo (possível retrocesso excessivo); resultado parcial.');
    }
    if (content) content.segments = null;
    if (matches.length === 0) return;

    if (!content || content.status === 'skipped-size') {
      // Somente os metadados (para o "salvo por último por"), sem ler o texto.
      const meta = await withTimeout(extractFile(entry.path, { size: st.size, limits: this.limits, withText: false }), this.fileTimeoutMs).catch(() => null);
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
      await this.deleteRecords([record]);
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
      // Ao cancelar, os registros são entregues sem esperar o proprietário (nada se perde).
      const owners = this.cancelled ? new Map() : await this.ownerResolver.resolve(batch.map((r) => r.path), { signal: this.abort.signal });
      for (const record of batch) {
        const info = owners.get(record.path);
        record.owner = info?.owner || null;
        record.ownerError = info?.error || (this.cancelled ? 'Análise cancelada antes de identificar o proprietário.' : null);
      }
      this.finishRecords(batch);
      await this.deleteRecords(batch);
    }
  }

  /**
   * Exclusão automática ("analisar e excluir"): os arquivos são excluídos depois de registrados no
   * relatório, com o proprietário e o último usuário já identificados, e só se continuarem iguais ao
   * que foi analisado (mesmo tamanho e data de modificação). Cada exclusão é registrada logo em
   * seguida; ao cancelar, nada mais é excluído.
   */
  async deleteRecords(records) {
    if (!this.options.deleteMatches || this.cancelled || records.length === 0) return;
    for (const record of records) {
      if (this.cancelled) break;
      const repo = this.repoById.get(record.repositoryId);
      const cloud = Boolean(record.cloud);
      const method = cloud ? (repo?.deleteMode === 'permanent' ? 'permanent' : 'trash') : 'file';
      if (this.retention) {
        await this.deleteExpired(record, repo, method);
        continue;
      }
      let result;
      if (!repo?.allowDelete) {
        result = { status: 'failed', error: 'A exclusão não está permitida neste repositório.' };
      } else if (cloud) {
        const connector = this.cloudConnector(repo);
        if (!this.cloudKept.has(repo.id)) this.cloudKept.set(repo.id, connector.resolveKept(repo.keep));
        const kept = keptCloudTarget(record.cloud, await this.cloudKept.get(repo.id));
        // O conector da análise usa o sinal de cancelamento; a exclusão em andamento termina mesmo assim.
        result = kept ? { status: 'failed', error: kept.error } : await connector.deleteItem(cloudTarget(record), method, { signal: null });
      } else {
        result = await deleteFile(record.path, {
          root: repo.path,
          expected: { size: record.size, modified: record.modified },
          protect: [...this.protect, ...(repo.keep || [])],
        });
      }
      countDeletion(this.stats, result.status);
      if (result.status === 'failed') this.error(record.path, `Falha ao excluir: ${result.error}`);
      this.emit({ type: 'deletions', items: [deletionEvent(record.id, result, { mode: 'auto', method, by: this.deletedBy, item: record.path })] });
    }
  }

  /**
   * Retenção: exclui um item expirado. Com a exclusão desligada no repositório (o aviso é registrado
   * uma vez) ou num local protegido, nada é tentado nem registrado como falha. O limite da execução
   * conta as exclusões feitas (a vaga fica reservada enquanto a exclusão está em andamento e volta
   * se ela não acontecer); as falhas têm o mesmo limite, para não repetir a mesma falha milhares de
   * vezes (permissão, rótulo de retenção...).
   */
  async deleteExpired(record, repo, method) {
    if (!repo?.allowDelete) return;
    const cloud = Boolean(record.cloud);
    const connector = cloud ? this.cloudConnector(repo) : null;
    let kept;
    if (cloud) {
      if (!this.cloudKept.has(repo.id)) this.cloudKept.set(repo.id, connector.resolveKept(repo.keep));
      kept = keptCloudTarget(record.cloud, await this.cloudKept.get(repo.id));
    } else {
      kept = guardFor([...this.protect, ...(repo.keep || [])], record.path);
    }
    if (kept) {
      this.stats.deleteProtected++;
      return;
    }
    const limit = this.retention.maxDeletions || 0;
    if (limit && (this.deleteAttempts >= limit || this.deleteFailures >= limit)) {
      this.stats.deleteSkipped++;
      if (!this.limitLogged) {
        this.limitLogged = true;
        const count = `${limit} ${limit === 1 ? 'exclusão' : 'exclusões'}`;
        this.log(
          'warn',
          this.deleteFailures >= limit
            ? `${limit === 1 ? 'Uma falha' : `${limit} falhas`} de exclusão nesta execução (o limite da política): os demais arquivos expirados foram apenas listados. Confira a aba Erros (permissões, arquivos em uso, rótulos de retenção...).`
            : `Limite de ${count} desta execução atingido: os demais arquivos expirados foram apenas listados. Confira o relatório e, se estiver certo, aumente o limite na política.`,
        );
      }
      return;
    }
    this.deleteAttempts++; // vaga reservada
    const result = cloud
      ? await connector.deleteItem(cloudTarget(record), method, { signal: null })
      : await deleteFile(record.path, {
          root: repo.path,
          expected: { size: record.size, modified: record.modified },
          protect: [...this.protect, ...(repo.keep || [])],
          check: (st) => this.stillExpired(st),
        });
    if (result.status !== 'deleted') this.deleteAttempts--;
    if (result.status === 'failed') this.deleteFailures++;
    countDeletion(this.stats, result.status);
    if (result.status === 'failed') this.error(record.path, `Falha ao excluir: ${result.error}`);
    this.emit({ type: 'deletions', items: [deletionEvent(record.id, result, { mode: 'retention', method, by: this.deletedBy, item: record.path })] });
  }

  /**
   * Retenção: o arquivo continua expirado na hora de excluir? A data do critério pode ter mudado
   * depois da listagem (ex.: o arquivo foi aberto). Devolve o motivo para mantê-lo, ou null.
   */
  stillExpired(st) {
    const when = fileDate(st, this.retention.criterion);
    if (when !== null && when < this.retention.cutoffMs) return null;
    return 'O arquivo deixou de estar expirado depois da listagem: a data do critério da política mudou (por exemplo, ele foi aberto ou recriado).';
  }

  finishRecords(records) {
    for (const record of records) {
      const pick = pickLastUser({ audit: record.audit, cloud: record.cloud?.lastModifiedBy, metadata: record.metadata, owner: record.owner });
      record.lastUser = pick.user;
      record.lastUserSource = pick.source;
    }
    this.emit({ type: 'results', records });
  }
}
