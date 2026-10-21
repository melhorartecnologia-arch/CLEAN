// Persistência em arquivos: configuração em data/db.json e resultados de cada análise em
// data/scans/<id>/ (NDJSON, uma linha por arquivo encontrado), sem dependência de banco de dados.
import fs from 'node:fs/promises';
import path from 'node:path';
import crypto from 'node:crypto';
import { SecretBox } from './secrets.js';

const EMPTY_DB = { version: 1, repositories: [], lists: [], mailSources: [], scans: [], schedules: [], intuneJobs: [], pendingRemovals: [] };
const MAX_LOG = 200;

function now() {
  return new Date().toISOString();
}

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

/** Repete a operação quando o arquivo está bloqueado (no Windows, antivírus e indexação bloqueiam arquivos por instantes). */
async function withRetry(operation) {
  for (let attempt = 0; ; attempt++) {
    try {
      return await operation();
    } catch (err) {
      if (attempt >= 8 || !['EPERM', 'EACCES', 'EBUSY'].includes(err.code)) throw err;
      await sleep(50 * (attempt + 1));
    }
  }
}

const renameWithRetry = (from, to) => withRetry(() => fs.rename(from, to));

// Leitura do registro geral de exclusões: blocos lidos em fluxo e, na leitura de linhas pela
// posição, o maior trecho lido de uma vez (linhas próximas vêm numa leitura só).
const LOG_CHUNK = 1024 * 1024;
const LOG_SPAN = 4 * 1024 * 1024;
// O registro não existe (ainda) ou o caminho não é um arquivo: lido como vazio.
const MISSING_LOG = new Set(['ENOENT', 'EISDIR', 'ENOTDIR']);

/** Uma linha NDJSON (bytes de start a end): o objeto, ou null se estiver vazia ou corrompida. */
function parseLine(buf, start, end) {
  const text = buf.toString('utf8', start, end);
  if (!text.trim()) return null;
  try {
    const value = JSON.parse(text);
    return value && typeof value === 'object' && !Array.isArray(value) ? value : null;
  } catch {
    return null; // linha corrompida: ignora
  }
}

/**
 * Contexto da análise copiado para cada linha do registro geral de exclusões, para que ele continue
 * completo depois que o relatório for excluído: tipo da análise, finalidade, agendamento ou política
 * e a regra da política de retenção (só os campos que a análise tiver).
 */
function deletionContext(scan) {
  const r = scan?.retention;
  const context = {
    scanKind: scan?.kind || 'files',
    purpose: r ? 'retention' : scan?.fileTypes ? 'types' : 'terms',
    scheduleId: scan?.scheduleId || null,
    scheduleName: scan?.scheduleName || null,
  };
  if (r) {
    const policy = {};
    const copy = (key) => {
      if (r[key] !== undefined && r[key] !== null && r[key] !== '') policy[key] = r[key];
    };
    copy('criterion');
    copy('cutoff');
    if (r.since) copy('since');
    else {
      copy('amount');
      copy('unit');
    }
    if (Array.isArray(r.folders) && r.folders.length) policy.folders = r.folders;
    context.policy = policy;
  }
  return context;
}

export class Store {
  constructor(dataDir) {
    this.dataDir = path.resolve(dataDir);
    this.dbFile = path.join(this.dataDir, 'db.json');
    this.scansDir = path.join(this.dataDir, 'scans');
    // Registro geral de exclusões: continua existindo mesmo que o relatório da análise seja excluído.
    this.deletionLog = path.join(this.dataDir, 'exclusoes.ndjson');
    this.db = null;
    this.saveChain = Promise.resolve();
    this.saveTimer = null;
    this.appendChains = new Map();
    this.resultCache = new Map();
    this.readChains = new Map();
    this.secrets = null;
  }

  async init() {
    await fs.mkdir(this.scansDir, { recursive: true });
    try {
      this.db = JSON.parse(await fs.readFile(this.dbFile, 'utf8'));
    } catch (err) {
      if (err.code !== 'ENOENT') throw new Error(`Não foi possível ler ${this.dbFile}: ${err.message}`);
      this.db = structuredClone(EMPTY_DB);
    }
    for (const key of ['repositories', 'lists', 'mailSources', 'scans', 'schedules', 'intuneJobs', 'pendingRemovals']) if (!Array.isArray(this.db[key])) this.db[key] = [];
    this.secrets = await SecretBox.open(this.dataDir);
    // Análises que estavam em andamento quando o servidor parou
    for (const scan of this.db.scans) {
      if (scan.status === 'running' || scan.status === 'queued') {
        scan.status = 'interrupted';
        scan.finishedAt ??= now();
        scan.current = null;
      }
    }
    // Pastas de relatórios excluídos que não puderam ser apagadas (arquivo bloqueado): nova tentativa.
    for (const id of this.db.pendingRemovals.splice(0)) await this.#removeScanDir(id);
    await this.saveNow();
    return this;
  }

  // -- gravação --------------------------------------------------------------------------------

  scheduleSave() {
    if (this.saveTimer) return;
    this.saveTimer = setTimeout(() => {
      this.saveTimer = null;
      this.saveNow().catch((err) => console.error('[CLEAN] Falha ao salvar dados:', err.message));
    }, 500);
  }

  saveNow() {
    if (this.saveTimer) {
      clearTimeout(this.saveTimer);
      this.saveTimer = null;
    }
    const json = JSON.stringify(this.db, null, 2);
    this.saveChain = this.saveChain
      .catch(() => {})
      .then(async () => {
        const tmp = `${this.dbFile}.${process.pid}.tmp`;
        await fs.writeFile(tmp, json, 'utf8');
        await renameWithRetry(tmp, this.dbFile);
      });
    return this.saveChain;
  }

  /** Grava as alterações pendentes (use antes de encerrar o processo). */
  async close() {
    await Promise.all([...this.appendChains.values()]);
    await this.saveNow();
  }

  // -- coleções genéricas ----------------------------------------------------------------------

  #list(kind) {
    return this.db[kind];
  }

  #get(kind, id) {
    return this.db[kind].find((item) => item.id === id) || null;
  }

  #create(kind, data) {
    const item = { id: crypto.randomUUID(), ...data, createdAt: now(), updatedAt: now() };
    this.db[kind].push(item);
    this.scheduleSave();
    return item;
  }

  #update(kind, id, patch, { touch = true } = {}) {
    const item = this.#get(kind, id);
    if (!item) return null;
    Object.assign(item, patch, touch ? { updatedAt: now() } : {});
    this.scheduleSave();
    return item;
  }

  #delete(kind, id) {
    const index = this.db[kind].findIndex((item) => item.id === id);
    if (index === -1) return false;
    this.db[kind].splice(index, 1);
    this.scheduleSave();
    return true;
  }

  listRepositories() {
    return this.#list('repositories');
  }
  getRepository(id) {
    return this.#get('repositories', id);
  }
  createRepository(data) {
    return this.#create('repositories', data);
  }
  updateRepository(id, data) {
    return this.#update('repositories', id, data);
  }
  deleteRepository(id) {
    return this.#delete('repositories', id);
  }

  listLists() {
    return this.#list('lists');
  }
  getList(id) {
    return this.#get('lists', id);
  }
  createList(data) {
    return this.#create('lists', data);
  }
  updateList(id, data) {
    return this.#update('lists', id, data);
  }
  deleteList(id) {
    return this.#delete('lists', id);
  }

  listMailSources() {
    return this.#list('mailSources');
  }
  getMailSource(id) {
    return this.#get('mailSources', id);
  }
  createMailSource(data) {
    return this.#create('mailSources', data);
  }
  updateMailSource(id, data) {
    return this.#update('mailSources', id, data);
  }
  deleteMailSource(id) {
    return this.#delete('mailSources', id);
  }

  /** Segredos decifrados de uma conexão de e-mail (nunca são gravados em claro nem enviados ao navegador). */
  openMailSecrets(source) {
    const s = source?.secrets || {};
    const open = (value) => (value ? this.secrets.open(value) : '');
    return {
      clientSecret: open(s.clientSecret),
      privateKey: open(s.privateKey),
      certificateKey: open(s.certificateKey),
      refreshToken: open(s.refreshToken),
      defaultPassword: open(s.defaultPassword),
      passwords: Object.fromEntries(Object.entries(s.passwords || {}).map(([address, value]) => [address, open(value)])),
    };
  }

  /**
   * Conta Microsoft conectada: grava (cifrado) o novo token de atualização que a Microsoft devolve a
   * cada renovação — só se a conexão ainda estiver com a mesma entrada da conta (grantId), para que
   * uma análise antiga não troque o token de uma conta conectada depois. Não muda a data de alteração.
   */
  saveRefreshToken(sourceId, grantId, token) {
    const source = this.getMailSource(sourceId);
    if (!source || !grantId || !token || source.graph?.account?.grantId !== grantId) return false;
    this.#update('mailSources', sourceId, { secrets: { ...(source.secrets || {}), refreshToken: this.secrets.seal(token) } }, { touch: false });
    return true;
  }

  /** Segredos do repositório do OneDrive/SharePoint, decifrados (nunca vão para o navegador). */
  openRepositorySecrets(repo) {
    const s = repo?.secrets || {};
    const open = (value) => (value ? this.secrets.open(value) : '');
    return { clientSecret: open(s.clientSecret), certificateKey: open(s.certificateKey) };
  }

  listSchedules() {
    return this.#list('schedules');
  }
  getSchedule(id) {
    return this.#get('schedules', id);
  }
  createSchedule(data) {
    return this.#create('schedules', data);
  }
  updateSchedule(id, data) {
    return this.#update('schedules', id, data);
  }
  /** Estado mantido pelo agendador (próxima execução, histórico): não muda a data de alteração. */
  updateScheduleState(id, patch) {
    return this.#update('schedules', id, patch, { touch: false });
  }
  deleteSchedule(id) {
    return this.#delete('schedules', id);
  }

  /** Agendamentos que usam um repositório ('files'), uma conexão de e-mail ('mail') ou uma lista ('list'). */
  schedulesUsing(kind, id) {
    return this.#list('schedules').filter((s) => (kind === 'list' ? s.listIds : s.kind === kind ? s.targetIds : [])?.includes(id));
  }

  // Reinício agendado de dispositivos no Intune (Microsoft Graph).
  listIntuneJobs() {
    return this.#list('intuneJobs');
  }
  getIntuneJob(id) {
    return this.#get('intuneJobs', id);
  }
  createIntuneJob(data) {
    return this.#create('intuneJobs', data);
  }
  updateIntuneJob(id, data) {
    return this.#update('intuneJobs', id, data);
  }
  /** Estado mantido pelo agendador (próxima execução, histórico): não muda a data de alteração. */
  updateIntuneJobState(id, patch) {
    return this.#update('intuneJobs', id, patch, { touch: false });
  }
  deleteIntuneJob(id) {
    return this.#delete('intuneJobs', id);
  }
  /** Agendamentos de reinício que usam uma conexão de e-mail (Graph). */
  intuneJobsUsing(sourceId) {
    return this.#list('intuneJobs').filter((j) => j.sourceId === sourceId);
  }

  listScans() {
    return this.#list('scans');
  }
  getScan(id) {
    return this.#get('scans', id);
  }
  createScan(data) {
    return this.#create('scans', { log: [], ...data });
  }
  updateScan(id, patch) {
    return this.#update('scans', id, patch, { touch: false });
  }

  appendLog(id, entry) {
    const scan = this.getScan(id);
    if (!scan) return;
    scan.log ||= [];
    scan.log.push({ time: entry.time || now(), level: entry.level || 'info', message: entry.message });
    if (scan.log.length > MAX_LOG) scan.log.splice(0, scan.log.length - MAX_LOG);
    this.scheduleSave();
  }

  /**
   * Exclui o relatório: sai do cadastro na hora; a pasta é apagada em seguida (com novas tentativas:
   * no Windows, antivírus e indexação bloqueiam arquivos por instantes). Se ainda assim não puder ser
   * apagada, fica anotada e é tentada de novo quando o CLEAN iniciar.
   */
  async deleteScan(id) {
    const dir = this.scanDir(id);
    const existed = this.#delete('scans', id);
    for (const key of [...this.resultCache.keys()]) if (key.startsWith(`${id}/`)) this.resultCache.delete(key);
    await this.#removeScanDir(id, dir);
    return existed;
  }

  async #removeScanDir(id, dir = this.scanDir(id)) {
    try {
      await fs.rm(dir, { recursive: true, force: true, maxRetries: 5, retryDelay: 100 });
    } catch (err) {
      console.error(`[CLEAN] Não foi possível apagar a pasta do relatório ${id} (nova tentativa ao iniciar): ${err.message}`);
      if (!this.db.pendingRemovals.includes(id)) this.db.pendingRemovals.push(id);
      this.scheduleSave();
    }
  }

  // -- arquivos de cada análise ----------------------------------------------------------------

  scanDir(id) {
    if (!/^[0-9a-f-]{36}$/i.test(id)) throw new Error('Identificador de análise inválido.');
    return path.join(this.scansDir, id);
  }

  async writeScanConfig(id, config) {
    await fs.mkdir(this.scanDir(id), { recursive: true });
    await fs.writeFile(path.join(this.scanDir(id), 'config.json'), JSON.stringify(config), 'utf8');
  }

  async readScanConfig(id) {
    return JSON.parse(await fs.readFile(path.join(this.scanDir(id), 'config.json'), 'utf8'));
  }

  /**
   * Acrescenta linhas a um arquivo NDJSON, em ordem. Com strict, uma falha de gravação é devolvida a
   * quem chamou (registro de exclusões); sem, ela só é informada no console.
   */
  #appendTo(target, items, { strict = false } = {}) {
    const data = `${items.map((item) => JSON.stringify(item)).join('\n')}\n`;
    const run = (this.appendChains.get(target) || Promise.resolve()).then(async () => {
      await fs.mkdir(path.dirname(target), { recursive: true });
      await withRetry(() => fs.appendFile(target, data, 'utf8'));
    });
    const settled = run.catch((err) => console.error(`[CLEAN] Falha ao gravar ${target}:`, err.message));
    this.appendChains.set(target, settled);
    return strict ? run : settled;
  }

  #append(id, file, items, options) {
    return this.#appendTo(path.join(this.scanDir(id), file), items, options);
  }

  appendResults(id, records) {
    return this.#append(id, 'results.ndjson', records);
  }

  appendErrors(id, errors) {
    return this.#append(id, 'errors.ndjson', errors);
  }

  /**
   * Registro das exclusões (automáticas e manuais) dos itens de uma análise, copiado também para o
   * registro geral (data/exclusoes.ndjson). No registro geral, cada linha leva também o contexto da
   * análise (tipo, finalidade, agendamento ou política e a regra da retenção); os campos do evento
   * prevalecem. A promessa é rejeitada se a gravação falhar.
   */
  async appendDeletions(id, items) {
    const scan = this.getScan(id);
    const scanName = scan?.name || '';
    const context = deletionContext(scan);
    await Promise.all([
      this.#append(id, 'deletions.ndjson', items, { strict: true }),
      this.#appendTo(this.deletionLog, items.map((item) => ({ scanId: id, scanName, ...context, ...item })), { strict: true }),
    ]);
  }

  /** Aguarda a gravação pendente dos arquivos de uma análise. */
  async flushScan(id) {
    const dir = this.scanDir(id);
    await Promise.all([...this.appendChains.entries()].filter(([file]) => file.startsWith(dir)).map(([, chain]) => chain));
  }

  /** Aguarda a gravação pendente do registro geral de exclusões (inclusive a que chegar enquanto espera). */
  async flushDeletionLog() {
    let waited = null;
    for (let chain = this.appendChains.get(this.deletionLog); chain && chain !== waited; chain = this.appendChains.get(this.deletionLog)) {
      waited = chain;
      await chain;
    }
  }

  /**
   * Lê o registro geral de exclusões em fluxo (ele pode ter centenas de MB): { offset, length, event }
   * para cada linha completa, na ordem do arquivo (offset e length em bytes, sem a quebra de linha).
   * Linhas corrompidas são ignoradas, assim como a última se ainda não tiver a quebra de linha (em
   * gravação). Sem o arquivo (ou se o caminho não for um arquivo), não devolve nada. end: lê só até
   * essa posição (ex.: o tamanho do arquivo quando a leitura começou).
   */
  async *readDeletionLog({ end = Infinity } = {}) {
    await this.flushDeletionLog();
    let fh;
    try {
      fh = await fs.open(this.deletionLog, 'r');
    } catch (err) {
      if (MISSING_LOG.has(err.code)) return;
      throw err;
    }
    try {
      const st = await fh.stat();
      if (!st.isFile()) return;
      const limit = Math.min(st.size, end);
      const block = Buffer.allocUnsafe(LOG_CHUNK);
      let rest = null; // começo de uma linha que continua no próximo bloco
      let restOffset = 0;
      let position = 0;
      while (position < limit) {
        let bytesRead;
        try {
          ({ bytesRead } = await fh.read(block, 0, Math.min(LOG_CHUNK, limit - position), position));
        } catch (err) {
          if (MISSING_LOG.has(err.code)) return;
          throw err;
        }
        if (!bytesRead) break;
        let data = block.subarray(0, bytesRead);
        let base = position;
        if (rest) {
          data = Buffer.concat([rest, data]);
          base = restOffset;
          rest = null;
        }
        position += bytesRead;
        let start = 0;
        for (let nl = data.indexOf(0x0a, start); nl !== -1; nl = data.indexOf(0x0a, start)) {
          const event = parseLine(data, start, nl);
          if (event) yield { offset: base + start, length: nl - start, event };
          start = nl + 1;
        }
        if (start < data.length) {
          rest = Buffer.from(data.subarray(start)); // cópia: o bloco é reaproveitado na próxima leitura
          restOffset = base + start;
        }
      }
    } finally {
      await fh.close();
    }
  }

  /**
   * Lê linhas do registro geral pela posição ({ offset, length } de readDeletionLog), ex.: uma página
   * depois de uma passada com filtros. Devolve os eventos na ordem pedida (null para uma linha que
   * não pôde ser lida, ex.: o registro foi substituído).
   */
  async readDeletionLines(refs) {
    const out = refs.map(() => null);
    if (!refs.length) return out;
    let fh;
    try {
      fh = await fs.open(this.deletionLog, 'r');
    } catch (err) {
      if (MISSING_LOG.has(err.code)) return out;
      throw err;
    }
    try {
      const order = refs.map((_, i) => i).sort((a, b) => refs[a].offset - refs[b].offset);
      for (let i = 0; i < order.length; ) {
        const first = refs[order[i]].offset;
        let end = first + refs[order[i]].length;
        let j = i;
        while (j + 1 < order.length && refs[order[j + 1]].offset + refs[order[j + 1]].length - first <= LOG_SPAN) {
          j++;
          end = Math.max(end, refs[order[j]].offset + refs[order[j]].length);
        }
        const buf = Buffer.alloc(end - first);
        let bytesRead = 0;
        try {
          if (buf.length) ({ bytesRead } = await fh.read(buf, 0, buf.length, first));
        } catch (err) {
          if (MISSING_LOG.has(err.code)) return out;
          throw err;
        }
        for (let k = i; k <= j; k++) {
          const ref = refs[order[k]];
          const start = ref.offset - first;
          if (start + ref.length <= bytesRead) out[order[k]] = parseLine(buf, start, start + ref.length);
        }
        i = j + 1;
      }
    } finally {
      await fh.close();
    }
    return out;
  }

  /**
   * Lê um arquivo NDJSON de forma incremental (útil durante a análise, quando ele ainda cresce).
   * Leituras do mesmo arquivo são enfileiradas: duas leituras simultâneas partiriam do mesmo
   * ponto do cache e duplicariam (ou pulariam) registros.
   */
  #readNdjson(id, file) {
    const key = `${id}/${file}`;
    const previous = this.readChains.get(key) || Promise.resolve();
    const next = previous.catch(() => {}).then(() => this.#readNdjsonNow(id, file));
    this.readChains.set(key, next);
    const release = () => {
      if (this.readChains.get(key) === next) this.readChains.delete(key);
    };
    next.then(release, release);
    return next;
  }

  async #readNdjsonNow(id, file) {
    const target = path.join(this.scanDir(id), file);
    const key = `${id}/${file}`;
    let st;
    try {
      st = await fs.stat(target);
    } catch {
      return [];
    }
    let cache = this.resultCache.get(key);
    if (!cache || st.size < cache.offset) cache = { offset: 0, rest: Buffer.alloc(0), records: [] };
    if (st.size > cache.offset) {
      const fh = await fs.open(target, 'r');
      try {
        const buf = Buffer.alloc(st.size - cache.offset);
        const { bytesRead } = await fh.read(buf, 0, buf.length, cache.offset);
        const data = Buffer.concat([cache.rest, buf.subarray(0, bytesRead)]);
        const lastNewline = data.lastIndexOf(0x0a);
        const complete = lastNewline >= 0 ? data.subarray(0, lastNewline) : Buffer.alloc(0);
        cache.rest = Buffer.from(lastNewline >= 0 ? data.subarray(lastNewline + 1) : data);
        for (const line of complete.toString('utf8').split('\n')) {
          if (!line.trim()) continue;
          try {
            cache.records.push(JSON.parse(line));
          } catch {
            // linha corrompida: ignora
          }
        }
        cache.offset += bytesRead;
      } finally {
        await fh.close();
      }
    }
    // Mantém em memória apenas as análises consultadas mais recentemente.
    this.resultCache.delete(key);
    this.resultCache.set(key, cache);
    while (this.resultCache.size > 8) this.resultCache.delete(this.resultCache.keys().next().value);
    return cache.records;
  }

  readResults(id) {
    return this.#readNdjson(id, 'results.ndjson');
  }

  readErrors(id) {
    return this.#readNdjson(id, 'errors.ndjson');
  }

  readDeletions(id) {
    return this.#readNdjson(id, 'deletions.ndjson');
  }
}
