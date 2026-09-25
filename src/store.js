// Persistência em arquivos: configuração em data/db.json e resultados de cada análise em
// data/scans/<id>/ (NDJSON, uma linha por arquivo encontrado), sem dependência de banco de dados.
import fs from 'node:fs/promises';
import path from 'node:path';
import crypto from 'node:crypto';

const EMPTY_DB = { version: 1, repositories: [], lists: [], scans: [] };
const MAX_LOG = 200;

function now() {
  return new Date().toISOString();
}

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

/** Renomeia com novas tentativas (no Windows, antivírus e indexação bloqueiam arquivos por instantes). */
async function renameWithRetry(from, to) {
  for (let attempt = 0; ; attempt++) {
    try {
      await fs.rename(from, to);
      return;
    } catch (err) {
      if (attempt >= 8 || !['EPERM', 'EACCES', 'EBUSY'].includes(err.code)) throw err;
      await sleep(50 * (attempt + 1));
    }
  }
}

export class Store {
  constructor(dataDir) {
    this.dataDir = path.resolve(dataDir);
    this.dbFile = path.join(this.dataDir, 'db.json');
    this.scansDir = path.join(this.dataDir, 'scans');
    this.db = null;
    this.saveChain = Promise.resolve();
    this.saveTimer = null;
    this.appendChains = new Map();
    this.resultCache = new Map();
    this.readChains = new Map();
  }

  async init() {
    await fs.mkdir(this.scansDir, { recursive: true });
    try {
      this.db = JSON.parse(await fs.readFile(this.dbFile, 'utf8'));
    } catch (err) {
      if (err.code !== 'ENOENT') throw new Error(`Não foi possível ler ${this.dbFile}: ${err.message}`);
      this.db = structuredClone(EMPTY_DB);
    }
    for (const key of ['repositories', 'lists', 'scans']) if (!Array.isArray(this.db[key])) this.db[key] = [];
    // Análises que estavam em andamento quando o servidor parou
    for (const scan of this.db.scans) {
      if (scan.status === 'running' || scan.status === 'queued') {
        scan.status = 'interrupted';
        scan.finishedAt ??= now();
        scan.current = null;
      }
    }
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

  async deleteScan(id) {
    const existed = this.#delete('scans', id);
    for (const key of [...this.resultCache.keys()]) if (key.startsWith(`${id}/`)) this.resultCache.delete(key);
    await fs.rm(this.scanDir(id), { recursive: true, force: true });
    return existed;
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

  #append(id, file, items) {
    const target = path.join(this.scanDir(id), file);
    const data = `${items.map((item) => JSON.stringify(item)).join('\n')}\n`;
    const chain = (this.appendChains.get(target) || Promise.resolve())
      .then(() => fs.mkdir(this.scanDir(id), { recursive: true }))
      .then(() => fs.appendFile(target, data, 'utf8'))
      .catch((err) => console.error(`[CLEAN] Falha ao gravar ${target}:`, err.message));
    this.appendChains.set(target, chain);
    return chain;
  }

  appendResults(id, records) {
    return this.#append(id, 'results.ndjson', records);
  }

  appendErrors(id, errors) {
    return this.#append(id, 'errors.ndjson', errors);
  }

  /** Aguarda a gravação pendente dos arquivos de uma análise. */
  async flushScan(id) {
    const dir = this.scanDir(id);
    await Promise.all([...this.appendChains.entries()].filter(([file]) => file.startsWith(dir)).map(([, chain]) => chain));
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
    while (this.resultCache.size > 4) this.resultCache.delete(this.resultCache.keys().next().value);
    return cache.records;
  }

  readResults(id) {
    return this.#readNdjson(id, 'results.ndjson');
  }

  readErrors(id) {
    return this.#readNdjson(id, 'errors.ndjson');
  }
}
