// Motor da análise de e-mail: percorre as caixas das conexões escolhidas, baixa cada mensagem e
// procura os termos no assunto, no corpo, nos nomes e no conteúdo dos anexos. O relatório traz
// apenas as mensagens com ocorrências, com remetente, destinatários, pasta e data.
import { Matcher } from '../scan/matcher.js';
import { createGuard, countDeletion } from '../scan/scanner.js';
import { extractMessage, DEFAULT_LIMITS } from '../scan/extractors/index.js';
import { formatAddress } from '../scan/extractors/mime.js';
import { friendlyError, withTimeout } from '../scan/errors.js';
import { createConnector } from './connectors.js';
import { validDate } from './common.js';
import { deletionEvent } from '../scan/delete.js';
import { ageDays, MIN_VALID_DATE, limitWarning } from '../retention/policy.js';

// Tempo máximo para ler uma mensagem já baixada (anexos incluídos).
const MESSAGE_TIMEOUT = 5 * 60 * 1000;

export const MAIL_DEFAULT_OPTIONS = {
  checkSubject: true,
  checkBody: true,
  checkAttachmentNames: true,
  checkAttachments: true,
  checkAddresses: false,
  receivedAfter: null, // ISO: somente mensagens recebidas a partir desta data
  includeTrash: true,
  includeJunk: false,
  maxMessageSizeMB: 50, // mensagens maiores: só o início é baixado
  concurrency: 4,
  maxSamples: 3,
  deleteMatches: false, // exclui automaticamente as mensagens em que algum termo for encontrado
};

export function newMailStats(sourcesTotal = 0) {
  return {
    sourcesTotal,
    sourcesDone: 0,
    mailboxesTotal: 0,
    mailboxesDone: 0,
    mailboxesSkipped: 0,
    folders: 0,
    messagesSeen: 0,
    messagesMatched: 0,
    occurrences: 0,
    messagesPartial: 0,
    messagesEncrypted: 0,
    attachmentsSeen: 0,
    attachmentsAnalyzed: 0,
    attachmentsEncrypted: 0,
    attachmentsUnsupported: 0,
    attachmentsSkippedSize: 0,
    attachmentsErrors: 0,
    bytesDownloaded: 0,
    errors: 0,
    gaps: 0, // conexões ou caixas que não puderam ser lidas (a análise ficou incompleta)
    bytesExpired: 0, // retenção: tamanho das mensagens expiradas
    deleteSkipped: 0, // retenção: expiradas não excluídas por causa do limite da execução
    retentionUnknown: 0, // retenção: sem data de recebimento válida (não expiram)
    alreadyInTrash: 0, // retenção "para a lixeira": expiradas que já estavam na Lixeira (não são movidas de novo)
    deleted: 0, // excluídas na análise ("analisar e excluir")
    deleteMissing: 0, // já não existiam na hora da exclusão
    deleteChanged: 0,
    deleteErrors: 0,
  };
}

/** Números de uma listagem do catálogo de contas do domínio (sem varrer mensagens). */
export function newDirectoryStats(sourcesTotal = 0) {
  return { sourcesTotal, sourcesDone: 0, accounts: 0, errors: 0, gaps: 0 };
}

const mb = (bytes) => `${Math.round((bytes / 1048576) * 10) / 10} MB`;
const plural = (n, one, many) => `${n} ${n === 1 ? one : many}`;

// Raio-X das caixas: censo de TODAS as mensagens analisadas (não só as com ocorrências) — quantidade
// por pasta e por caixa, mais antiga/recente e distribuição por mês. Limites do tamanho em memória e no
// registro da análise: além deles, as mensagens são só contadas (em "outras pastas/caixas").
const MAX_PROFILE_FOLDERS = 2000;
const MAX_PROFILE_MAILBOXES = 5000;
const PROFILE_TOP = 500; // pastas e caixas guardadas no relatório (as demais entram só nos totais)
const monthKey = (ms) => new Date(ms).toISOString().slice(0, 7); // 'AAAA-MM' (UTC)

function newProfile() {
  return { total: 0, withoutDate: 0, folders: new Map(), mailboxes: new Map(), timeline: new Map(), oldest: null, newest: null, folderOverflow: 0, mailboxOverflow: 0 };
}

export class MailScanner {
  /**
   * config: { sources: [conexões com secrets], terms: [...], options: {...}, endpoints? }
   * emit(message): recebe { type: 'log'|'progress'|'results'|'errors'|'deletions'|'credentials'|'done', ... }
   */
  constructor(config, emit, { connectorFactory = createConnector } = {}) {
    this.sources = config.sources || [];
    this.options = { ...MAIL_DEFAULT_OPTIONS, ...(config.options || {}) };
    // Listagens (sem termos): catálogo de contas do domínio ('directory') ou mensagens por caixa ('messages').
    this.listing = config.listing || null;
    this.matcher = new Matcher(config.terms || [], { maxSamples: this.options.maxSamples, guard: createGuard(config.regexTimeoutMs || 30000) });
    this.abort = new AbortController();
    this.emit = emit;
    this.connectorFactory = connectorFactory;
    this.endpoints = config.endpoints || {};
    this.stats = this.listing?.kind === 'directory' ? newDirectoryStats(this.sources.length) : newMailStats(this.sources.length);
    this.cancelled = false;
    this.seq = 0;
    this.inFlight = new Set(); // mensagens sendo lidas (para mensagens de erro)
    this.pendingErrors = [];
    this.pendingResults = [];
    this.pendingDeletes = []; // mensagens da caixa atual a excluir (exclusão automática)
    this.lastProgress = 0;
    this.current = null;
    this.maxBytes = Math.max(1, Number(this.options.maxMessageSizeMB) || 50) * 1048576;
    this.limits = { ...DEFAULT_LIMITS, maxBytes: this.maxBytes };
    this.messageTimeoutMs = config.messageTimeoutMs || MESSAGE_TIMEOUT;
    this.since = validDate(this.options.receivedAfter);
    this.deletedBy = config.startedBy || null; // quem iniciou a análise com exclusão automática
    // Política de retenção: mensagens recebidas antes da data de corte (só os cabeçalhos são lidos).
    const r = config.retention || null;
    this.retention = r ? { ...r, cutoffMs: Date.parse(r.cutoff) } : null;
    // Limite de exclusões da execução: vagas em uso (exclusões feitas ou na fila) e falhas.
    this.deleteQueued = 0;
    this.deleteFailures = 0;
    // Raio-X das caixas (censo), só nas análises por termos — na retenção e nas listagens não se aplica.
    this.profile = this.retention || this.listing ? null : newProfile();
    this.lastProfile = 0;
  }

  /**
   * Raio-X: soma a mensagem ao censo (quantidade por pasta e por caixa, mais antiga/recente e o mês da
   * distribuição ao longo do tempo). Vale para toda mensagem real, com ou sem ocorrências.
   */
  census(mailbox, item) {
    const p = this.profile;
    if (!p) return;
    p.total++;
    const path = item.folder || '(sem pasta)';
    let f = p.folders.get(path);
    if (!f && p.folders.size >= MAX_PROFILE_FOLDERS) p.folderOverflow++;
    else if (!f) p.folders.set(path, (f = { count: 0, oldestMs: null, newestMs: null }));
    if (f) f.count++;
    const addr = mailbox.address;
    let mbx = p.mailboxes.get(addr);
    if (!mbx && p.mailboxes.size >= MAX_PROFILE_MAILBOXES) p.mailboxOverflow++;
    else if (!mbx) p.mailboxes.set(addr, (mbx = { name: mailbox.name || '', count: 0, oldestMs: null }));
    if (mbx) {
      mbx.count++;
      if (!mbx.name && mailbox.name) mbx.name = mailbox.name;
    }
    const t = Date.parse(item.receivedAt);
    // Sem data de recebimento válida (inclusive datas zeradas de mensagens migradas): fora da linha do tempo.
    if (!Number.isFinite(t) || t < MIN_VALID_DATE) {
      p.withoutDate++;
      return;
    }
    p.timeline.set(monthKey(t), (p.timeline.get(monthKey(t)) || 0) + 1);
    if (p.oldest === null || t < p.oldest.ms) p.oldest = { ms: t, mailbox: addr, folder: path };
    if (p.newest === null || t > p.newest.ms) p.newest = { ms: t, mailbox: addr, folder: path };
    if (f) {
      if (f.oldestMs === null || t < f.oldestMs) f.oldestMs = t;
      if (f.newestMs === null || t > f.newestMs) f.newestMs = t;
    }
    if (mbx && (mbx.oldestMs === null || t < mbx.oldestMs)) mbx.oldestMs = t;
  }

  /** Snapshot do raio-x para o relatório: ordenado por quantidade e limitado ao tamanho guardado. */
  buildProfile() {
    const p = this.profile;
    if (!p) return null;
    const iso = (ms) => (ms == null ? null : new Date(ms).toISOString());
    const folders = [...p.folders.entries()]
      .map(([path, f]) => ({ path, count: f.count, oldest: iso(f.oldestMs), newest: iso(f.newestMs) }))
      .sort((a, b) => b.count - a.count);
    const mailboxes = [...p.mailboxes.entries()]
      .map(([mailbox, m]) => ({ mailbox, name: m.name, count: m.count, oldest: iso(m.oldestMs) }))
      .sort((a, b) => b.count - a.count);
    const timeline = [...p.timeline.entries()].map(([month, count]) => ({ month, count })).sort((a, b) => a.month.localeCompare(b.month));
    const place = (o) => (o ? { date: iso(o.ms), mailbox: o.mailbox, folder: o.folder } : null);
    return {
      total: p.total,
      withoutDate: p.withoutDate,
      oldest: place(p.oldest),
      newest: place(p.newest),
      timeline,
      folders: folders.slice(0, PROFILE_TOP),
      foldersTotal: folders.length,
      folderOverflow: p.folderOverflow,
      mailboxes: mailboxes.slice(0, PROFILE_TOP),
      mailboxesTotal: mailboxes.length,
      mailboxOverflow: p.mailboxOverflow,
      since: this.options.receivedAfter || null,
    };
  }

  /** Emite o raio-x (para o gerenciador gravar em scan.profile). Limitado no tempo, a não ser no fim. */
  emitProfile(force = false) {
    if (!this.profile) return;
    const now = Date.now();
    if (!force && now - this.lastProfile < 5000) return;
    this.lastProfile = now;
    this.emit({ type: 'profile', profile: this.buildProfile() });
  }

  cancel() {
    this.cancelled = true;
    this.abort.abort(new Error('Análise cancelada.'));
  }

  /** A exclusão foi desligada no cadastro durante a análise: nada mais é excluído daquela conexão. */
  revokeDeletion({ kind, id, reason, all = false }) {
    if (all) {
      const active = this.sources.filter((s) => s.allowDelete);
      for (const source of active) source.allowDelete = false;
      if (active.length && this.options.deleteMatches) this.log('warn', `Exclusão automática desativada: ${reason}. A análise continua sem excluir.`);
      return;
    }
    if (kind !== 'mail') return;
    for (const source of this.sources) {
      if (source.id !== id || !source.allowDelete) continue;
      source.allowDelete = false;
      if (this.options.deleteMatches) this.log('warn', `Exclusão automática desativada para "${source.name}": ${reason || 'o cadastro da conexão foi alterado'}.`);
    }
  }

  log(level, message) {
    this.emit({ type: 'log', level, message, time: new Date().toISOString() });
  }

  error(where, err) {
    this.stats.errors++;
    this.pendingErrors.push({ path: where, message: typeof err === 'string' ? err : friendlyError(err), time: new Date().toISOString() });
    if (this.pendingErrors.length >= 200) this.flushErrors();
  }

  flushErrors() {
    if (this.pendingErrors.length === 0) return;
    this.emit({ type: 'errors', items: this.pendingErrors.splice(0) });
  }

  flushResults() {
    if (this.pendingResults.length === 0) return;
    this.emit({ type: 'results', records: this.pendingResults.splice(0) });
  }

  progress(force = false) {
    const now = Date.now();
    if (!force && now - this.lastProgress < 400) return;
    this.lastProgress = now;
    this.flushResults();
    this.emit({ type: 'progress', stats: { ...this.stats }, current: this.current });
  }

  async run() {
    const started = Date.now();
    if (this.listing) return this.runListing(started);
    if (this.retention) {
      const cutoff = new Date(this.retention.cutoffMs).toLocaleString('pt-BR', { dateStyle: 'short', timeStyle: 'short' });
      this.log('info', `Retenção iniciada em ${this.sources.length} conexão(ões): mensagens recebidas antes de ${cutoff}.`);
    } else {
      for (const { term, error } of this.matcher.invalid) this.log('warn', `Termo ignorado "${term.value}": ${error}`);
      if (this.matcher.size === 0) this.log('warn', 'Nenhum termo válido nas listas selecionadas.');
      this.log('info', `Análise de e-mail iniciada com ${this.matcher.size} termo(s) em ${this.sources.length} conexão(ões).`);
    }
    for (const source of this.sources) {
      if (this.cancelled) break;
      await this.scanSource(source);
      this.stats.sourcesDone++;
      this.progress(true);
    }
    this.flushResults();
    this.flushErrors();
    this.emitProfile(true); // raio-x final (antes do "done", para o gerenciador gravá-lo)
    this.current = null;
    const s = this.stats;
    const seconds = Math.round((Date.now() - started) / 1000);
    // Retenção: aviso do limite pelos números finais (as exclusões de cada caixa são feitas no fim dela).
    const warning = this.retention ? limitWarning({ limit: this.retention.maxDeletions, deleted: s.deleted, failures: this.deleteFailures, skipped: s.deleteSkipped }, 'mail') : null;
    if (warning) this.log('warn', warning);
    this.log(
      'info',
      `${this.cancelled ? 'Análise cancelada' : 'Análise concluída'} em ${seconds}s: ${s.messagesSeen} mensagem(ns) verificadas em ${s.mailboxesDone} caixa(s), ${s.messagesMatched} ${this.retention ? 'expirada(s)' : 'com ocorrências'}.`,
    );
    this.emit({ type: 'done', stats: { ...s }, cancelled: this.cancelled });
    return s;
  }

  /** Cria o conector da conexão (com o repasse do token renovado das contas Microsoft conectadas). */
  makeConnector(source) {
    const grantId = source.graph?.account?.grantId || null;
    return this.connectorFactory(source, {
      signal: this.abort.signal,
      endpoints: this.endpoints,
      log: (level, message) => this.log(level, `${source.name}: ${message}`),
      // Conta Microsoft conectada: o novo token de atualização vai para o servidor, que o grava.
      onRefreshToken: grantId ? (refreshToken) => this.emit({ type: 'credentials', sourceId: source.id, grantId, refreshToken }) : undefined,
    });
  }

  /**
   * Listagens (sem termos): o catálogo de contas do domínio ('directory') ou a listagem de mensagens
   * por caixa ('messages', só os cabeçalhos, sem baixar as mensagens). Uma linha por conta ou mensagem.
   */
  async runListing(started) {
    const directory = this.listing.kind === 'directory';
    this.log('info', `${directory ? 'Listagem das contas do domínio' : 'Listagem de mensagens'} iniciada em ${plural(this.sources.length, 'conexão', 'conexões')}.`);
    for (const source of this.sources) {
      if (this.cancelled) break;
      if (directory) await this.listDirectory(source);
      else await this.scanSource(source); // reutiliza a varredura das caixas (só os cabeçalhos)
      this.stats.sourcesDone++;
      this.progress(true);
    }
    this.flushResults();
    this.flushErrors();
    this.current = null;
    const s = this.stats;
    const seconds = Math.round((Date.now() - started) / 1000);
    const done = directory ? `${plural(s.accounts, 'conta listada', 'contas listadas')}` : `${plural(s.messagesSeen, 'mensagem listada', 'mensagens listadas')} em ${plural(s.mailboxesDone, 'caixa', 'caixas')}`;
    this.log('info', `${this.cancelled ? 'Listagem cancelada' : 'Listagem concluída'} em ${seconds}s: ${done}.`);
    this.emit({ type: 'done', stats: { ...s }, cancelled: this.cancelled });
    return s;
  }

  /** Catálogo de contas de uma conexão: uma linha por conta registrada no domínio. */
  async listDirectory(source) {
    this.current = { source: source.name, mailbox: null, folder: null, path: source.name };
    this.progress(true);
    let connector;
    try {
      connector = this.makeConnector(source);
    } catch (err) {
      if (this.cancelled) return;
      this.error(source.name, err);
      this.stats.gaps++;
      this.log('error', `Conexão "${source.name}" indisponível: ${friendlyError(err)}`);
      return;
    }
    try {
      let count = 0;
      for await (const acct of connector.directory()) {
        if (this.cancelled) break;
        this.stats.accounts++;
        count++;
        this.pendingResults.push(this.accountRecord(source, acct));
        this.current = { source: source.name, mailbox: acct.address, folder: null, path: acct.address };
        if (this.pendingResults.length >= 100) this.flushResults();
        this.progress();
      }
      if (!this.cancelled) this.log('info', `Conexão "${source.name}": ${plural(count, 'conta listada', 'contas listadas')}.`);
    } catch (err) {
      if (this.cancelled) return;
      this.error(source.name, err);
      this.stats.gaps++;
      this.log('error', `Conexão "${source.name}": ${friendlyError(err)}`);
    } finally {
      await connector.close?.();
    }
  }

  /** Uma conta do catálogo do domínio, no formato do relatório. */
  accountRecord(source, acct) {
    return {
      id: ++this.seq,
      kind: 'mail-account',
      sourceId: source.id,
      sourceName: source.name,
      sourceType: source.type,
      address: acct.address || '',
      addressLower: (acct.address || '').toLowerCase(),
      name: acct.name || '',
      login: acct.login || '',
      aliases: Array.isArray(acct.aliases) ? acct.aliases : [],
      enabled: acct.enabled ?? null,
      type: acct.type || '',
      licensed: acct.licensed ?? null,
      created: acct.created || null,
      lastActivity: acct.lastActivity || null,
      department: acct.department || '',
      title: acct.title || '',
      location: acct.location || '',
      phone: acct.phone || '',
      orgUnit: acct.orgUnit || '',
      admin: acct.admin ?? null,
      suspended: acct.suspended ?? null,
      note: acct.note || '',
    };
  }

  /** Uma mensagem listada (só os cabeçalhos), no formato do relatório. */
  processListedMessage(source, mailbox, item) {
    this.stats.messagesSeen++;
    const from = item.from || null;
    this.pendingResults.push({
      id: ++this.seq,
      kind: 'mail-message',
      sourceId: source.id,
      sourceName: source.name,
      sourceType: source.type,
      mailbox: mailbox.address,
      mailboxName: mailbox.name || '',
      folder: item.folder,
      messageId: item.id,
      internetMessageId: item.internetMessageId || null,
      subject: item.subject || '',
      from: from ? formatAddress(from) : '',
      fromAddress: (from?.address || '').toLowerCase(),
      to: (item.to || []).map(formatAddress),
      cc: (item.cc || []).map(formatAddress),
      date: item.receivedAt || null,
      sent: item.sent || null,
      size: Number(item.size) || 0,
      hasAttachments: item.hasAttachments ?? null,
      inTrash: Boolean(item.inTrash),
      webLink: item.webLink || null,
    });
    if (this.pendingResults.length >= 100) this.flushResults();
  }

  async scanSource(source) {
    this.current = { source: source.name, mailbox: null, folder: null, path: source.name };
    this.progress(true);
    let connector;
    let mailboxes;
    try {
      connector = this.makeConnector(source);
      mailboxes = await connector.mailboxes();
    } catch (err) {
      if (this.cancelled) return;
      this.error(source.name, err);
      this.stats.gaps++;
      this.log('error', `Conexão "${source.name}" indisponível: ${friendlyError(err)}`);
      return;
    }
    this.stats.mailboxesTotal += mailboxes.length;
    this.log('info', `Conexão "${source.name}": ${mailboxes.length} caixa(s) a ${this.listing ? 'listar' : 'analisar'}.`);
    try {
      for (const [index, mailbox] of mailboxes.entries()) {
        if (this.cancelled) break;
        await this.scanMailbox(connector, source, mailbox);
        this.stats.mailboxesDone++;
        this.progress(true);
        this.emitProfile(); // raio-x parcial durante análises longas
        // Recusa definitiva da Microsoft (autorização revogada, segredo inválido...): as outras caixas
        // da conexão falhariam do mesmo jeito — ficam registradas uma vez, como não analisadas.
        const failure = connector.auth?.failure;
        const rest = mailboxes.length - index - 1;
        if (failure && rest > 0) {
          const message = `${plural(rest, 'caixa', 'caixas')} da conexão não ${rest === 1 ? 'foi analisada' : 'foram analisadas'}: ${friendlyError(failure)}`;
          this.error(source.name, message);
          this.stats.gaps++;
          this.log('error', `Conexão "${source.name}": ${message}`);
          break;
        }
      }
    } finally {
      await connector.close?.();
    }
  }

  async scanMailbox(connector, source, mailbox) {
    this.current = { source: source.name, mailbox: mailbox.address, folder: null, path: mailbox.address };
    this.progress(true);
    const before = this.stats.messagesSeen;
    this.pendingDeletes = [];
    const listingMessages = this.listing?.kind === 'messages';
    try {
      const items = connector.messages(mailbox, {
        since: this.since,
        // Retenção: só as recebidas antes da data de corte, sem baixar as mensagens.
        before: this.retention ? new Date(this.retention.cutoffMs) : null,
        // Retenção e listagem de mensagens leem só os cabeçalhos (a listagem pede os destinatários também).
        headersOnly: Boolean(this.retention) || listingMessages,
        fullHeaders: listingMessages,
        includeTrash: this.options.includeTrash,
        includeJunk: this.options.includeJunk,
        maxBytes: this.maxBytes,
        concurrency: Math.min(Math.max(1, Number(this.options.concurrency) || 4), 8),
        onFolder: () => {
          this.stats.folders++;
        },
      });
      for await (const item of items) {
        if (this.cancelled) break;
        if (item.id && this.profile) this.census(mailbox, item); // raio-x: conta todo e-mail real (mesmo os que falharam no download)
        if (item.error) {
          const where = [mailbox.address, item.folder, item.id].filter(Boolean).join(' › ');
          this.error(where, item.id ? `Falha ao ${listingMessages ? 'ler os dados da mensagem' : 'baixar a mensagem'}: ${friendlyError(item.error)}` : friendlyError(item.error));
          if (!item.id) this.stats.gaps++; // uma pasta inteira não pôde ser lida
        } else if (listingMessages) {
          this.processListedMessage(source, mailbox, item);
        } else if (this.retention) {
          this.processExpiredMessage(source, mailbox, item);
        } else {
          await this.processMessage(source, mailbox, item);
        }
        this.progress();
      }
      if (!this.cancelled) this.log('info', `Caixa ${mailbox.address}: ${this.stats.messagesSeen - before} mensagem(ns) ${this.listing ? 'listadas' : 'verificadas'}.`);
    } catch (err) {
      if (this.cancelled) return;
      if (err?.skipMailbox) {
        this.stats.mailboxesSkipped++;
        this.log('warn', `Caixa ${mailbox.address} ignorada: ${err.message}`);
        return;
      }
      this.error(mailbox.address, err);
      this.stats.gaps++;
      this.log('error', `Falha na caixa ${mailbox.address}: ${friendlyError(err)}`);
    } finally {
      // Exclusão automática ao fim de cada caixa: excluir durante a listagem deslocaria a
      // paginação do servidor e mensagens poderiam ficar sem análise.
      const pending = this.pendingDeletes.splice(0);
      if (!this.cancelled && pending.length) await this.deleteMessages(connector, source, mailbox, pending);
    }
  }

  /**
   * Exclusão automática das mensagens encontradas em uma caixa (ao fim dela). Cada resultado é
   * registrado assim que o servidor responde; ao cancelar, as exclusões em andamento terminam (e são
   * registradas) e as demais não começam.
   */
  async deleteMessages(connector, source, mailbox, pending) {
    // Retenção com a exclusão desligada durante a análise (o aviso já foi registrado): nada foi
    // tentado, então nada vira falha; as vagas do limite voltam.
    if (this.retention && !source.allowDelete) {
      this.deleteQueued -= pending.length;
      return;
    }
    const method = source.deleteMode === 'trash' ? 'trash' : 'permanent';
    this.flushResults(); // os registros chegam antes dos eventos de exclusão
    this.current = { source: source.name, mailbox: mailbox.address, folder: null, path: `${mailbox.address} › excluindo ${pending.length} mensagem(ns)` };
    this.progress(true);
    this.log('info', `Caixa ${mailbox.address}: excluindo ${pending.length} mensagem(ns) ${method === 'trash' ? '(movendo para a lixeira)' : '(definitivamente)'}.`);
    // A mesma mensagem pode aparecer duas vezes (ex.: movida durante a listagem): uma exclusão só.
    const byId = new Map();
    for (const p of pending) byId.set(p.messageId, [...(byId.get(p.messageId) || []), p]);
    const record = (messageId, r) => {
      const list = byId.get(messageId);
      if (!list) return;
      byId.delete(messageId);
      const status = r.ok ? 'deleted' : r.missing ? 'missing' : 'failed';
      // Retenção: a vaga no limite da execução volta quando a mensagem não foi excluída.
      if (this.retention && status !== 'deleted') this.deleteQueued -= list.length;
      if (this.retention && status === 'failed') this.deleteFailures += list.length;
      const items = list.map((p) => {
        countDeletion(this.stats, status);
        if (status === 'failed') this.error(`${mailbox.address} › ${p.folder}`, `Falha ao excluir a mensagem "${p.subject || '(sem assunto)'}": ${r.error}`);
        const item = `${mailbox.address} › ${p.folder} › ${p.subject || '(sem assunto)'}`;
        return deletionEvent(p.recordId, { status, error: r.ok ? null : r.error, note: r.note }, { mode: this.retention ? 'retention' : 'auto', method, by: this.deletedBy, item });
      });
      this.emit({ type: 'deletions', items });
      this.progress();
    };
    let failure = source.allowDelete ? null : 'A exclusão não está permitida nesta conexão.';
    if (!failure) {
      try {
        const items = [...byId].map(([id, list]) => ({ id, messageId: list[0].internetMessageId }));
        await connector.deleteMessages(mailbox, items, method, { signal: null, onResult: record, shouldStop: () => this.cancelled || !source.allowDelete });
      } catch (err) {
        failure = friendlyError(err);
        this.log('error', `Falha ao excluir mensagens da caixa ${mailbox.address}: ${failure}`);
      }
    }
    if (this.cancelled) {
      if (byId.size) this.log('warn', `Cancelado: ${byId.size} mensagem(ns) da caixa ${mailbox.address} não foram excluídas.`);
      return;
    }
    if (!failure && !source.allowDelete) {
      if (this.retention) {
        // Retenção: as que ainda não foram tentadas não viram falhas (as vagas voltam).
        for (const list of byId.values()) this.deleteQueued -= list.length;
        return;
      }
      failure = 'A exclusão foi desativada no cadastro da conexão durante a análise.';
    }
    for (const id of [...byId.keys()]) record(id, { ok: false, error: failure || 'O servidor não confirmou a exclusão.' });
    this.progress(true);
  }

  /** Retenção: mensagem recebida antes da data de corte (listada e, com exclusão, excluída). */
  processExpiredMessage(source, mailbox, item) {
    this.stats.messagesSeen++;
    const date = Date.parse(item.receivedAt);
    // O servidor já filtrou pela data. Sem data de recebimento válida (inclusive datas zeradas, como
    // 01/01/1970, de mensagens migradas), a mensagem não expira; depois do corte, também não.
    if (!Number.isFinite(date) || date < MIN_VALID_DATE) {
      this.stats.retentionUnknown++;
      return;
    }
    if (date >= this.retention.cutoffMs) return;
    this.stats.messagesMatched++;
    this.stats.bytesExpired += Number(item.size) || 0;
    const from = item.from || null;
    this.pendingResults.push({
      id: ++this.seq,
      kind: 'mail',
      sourceId: source.id,
      sourceName: source.name,
      sourceType: source.type,
      mailbox: mailbox.address,
      mailboxName: mailbox.name || '',
      folder: item.folder,
      messageId: item.id,
      internetMessageId: item.internetMessageId || null,
      subject: item.subject || '',
      from: from ? formatAddress(from) : '',
      fromAddress: (from?.address || '').toLowerCase(),
      to: [],
      cc: [],
      date: item.receivedAt,
      sent: null,
      size: Number(item.size) || 0,
      attachments: [],
      contentStatus: 'not-requested',
      contentNote: null,
      webLink: item.webLink || null,
      occurrences: 0,
      terms: [],
      matches: [],
      retention: { criterion: 'received', date: item.receivedAt, ageDays: ageDays(date) },
      // Já na Lixeira: com a exclusão "para a lixeira", a mensagem não é movida de novo (o provedor a
      // apaga depois, pela regra da própria Lixeira).
      inTrash: Boolean(item.inTrash),
    });
    // Com a exclusão desligada na conexão (o aviso é registrado uma vez), nada é colocado na fila.
    if (this.options.deleteMatches && source.allowDelete) {
      if (item.inTrash && source.deleteMode === 'trash') {
        this.stats.alreadyInTrash++;
      } else {
        this.queueExpiredDelete(item);
      }
    }
    if (this.pendingResults.length >= 50) this.flushResults();
  }

  /**
   * Retenção: coloca a mensagem na fila de exclusão da caixa, respeitando o limite da execução (que
   * conta as exclusões feitas ou na fila: a vaga volta se a mensagem não for excluída) e o mesmo
   * limite para as falhas.
   */
  queueExpiredDelete(item) {
    const limit = this.retention.maxDeletions || 0;
    if (limit && (this.deleteQueued >= limit || this.deleteFailures >= limit)) {
      this.stats.deleteSkipped++; // o aviso sai no fim, pelos números finais
      return;
    }
    this.deleteQueued++;
    this.pendingDeletes.push({ recordId: this.seq, messageId: item.id, internetMessageId: item.internetMessageId || null, folder: item.folder, subject: item.subject || '' });
  }

  countAttachments(attachments) {
    const s = this.stats;
    for (const a of attachments) {
      if (a.inline && a.status !== 'ok' && a.status !== 'partial') continue; // imagens da assinatura etc.
      s.attachmentsSeen++;
      switch (a.status) {
        case 'ok':
        case 'partial':
        case 'empty':
          s.attachmentsAnalyzed++;
          break;
        case 'encrypted':
          s.attachmentsEncrypted++;
          break;
        case 'skipped-size':
          s.attachmentsSkippedSize++;
          break;
        case 'error':
          s.attachmentsErrors++;
          break;
        case 'not-requested':
          break;
        default:
          s.attachmentsUnsupported++;
      }
    }
  }

  async processMessage(source, mailbox, item) {
    const { options, matcher, stats } = this;
    stats.messagesSeen++;
    stats.bytesDownloaded += item.raw.length;
    const where = `${mailbox.address} › ${item.folder}`;
    this.current = { source: source.name, mailbox: mailbox.address, folder: item.folder, path: where };
    const key = `${where} › ${item.id}`;
    this.inFlight.add(key);
    try {
      const rawLength = item.raw.length;
      let message;
      try {
        message = await withTimeout(
          extractMessage(item.raw, { limits: this.limits, truncated: item.truncated, attachments: options.checkAttachments, omittedToken: item.omittedToken }),
          this.messageTimeoutMs,
          'Tempo esgotado ao ler a mensagem.',
        );
      } catch (err) {
        this.error(key, `Falha ao ler a mensagem: ${friendlyError(err)}`);
        return;
      }
      item.raw = null;
      const unreadable = message.encrypted || message.opaqueSigned;
      const partial = message.partial || Boolean(item.partial);
      if (partial) stats.messagesPartial++;
      if (unreadable) stats.messagesEncrypted++;
      const attachments = message.attachments;
      this.countAttachments(attachments);

      const groups = [];
      if (options.checkSubject && message.subject) groups.push({ location: 'subject', segments: [{ text: message.subject, label: 'Assunto' }] });
      if (options.checkBody && message.body) groups.push({ location: 'body', segments: [{ text: message.body, label: 'Corpo da mensagem' }] });
      if (options.checkAttachmentNames && attachments.length) {
        groups.push({ location: 'attachmentName', segments: attachments.map((a) => ({ text: a.name, label: 'Nome do anexo' })) });
      }
      if (options.checkAttachments) groups.push({ location: 'attachment', segments: attachments.flatMap((a) => a.segments) });
      if (options.checkAddresses) {
        const people = [...message.from, ...message.to, ...message.cc, ...message.bcc].map(formatAddress).filter(Boolean);
        if (people.length) groups.push({ location: 'address', segments: [{ text: people.join('\n'), label: 'Remetente e destinatários' }] });
      }
      const matches = matcher.matchGroups(groups).flat();
      if (matcher.timedOut) this.error(key, 'Tempo limite ao procurar as expressões regulares nesta mensagem (possível retrocesso excessivo); resultado parcial.');
      for (const a of attachments) a.segments = null;
      if (matches.length === 0) return;

      const occurrences = matches.reduce((sum, m) => sum + m.count, 0);
      stats.messagesMatched++;
      stats.occurrences += occurrences;
      const from = message.from[0] || message.sender[0] || null;
      let status = 'ok';
      let note = null;
      if (unreadable && !message.body) {
        status = 'encrypted';
        note = 'Mensagem criptografada (S/MIME ou PGP): apenas o assunto e os remetentes foram verificados.';
      } else if (partial) {
        status = 'partial';
        note =
          item.note ||
          (item.truncated ? `Mensagem com ${mb(item.size)}: apenas os primeiros ${mb(this.maxBytes)} foram analisados.` : 'Mensagem incompleta: apenas parte foi analisada.');
      }
      this.pendingResults.push({
        id: ++this.seq,
        kind: 'mail',
        sourceId: source.id,
        sourceName: source.name,
        sourceType: source.type,
        mailbox: mailbox.address,
        mailboxName: mailbox.name || '',
        folder: item.folder,
        messageId: item.id,
        internetMessageId: message.messageId,
        subject: message.subject,
        from: from ? formatAddress(from) : '',
        fromAddress: (from?.address || '').toLowerCase(),
        to: message.to.map(formatAddress),
        cc: message.cc.map(formatAddress),
        date: item.receivedAt || message.date || null,
        sent: message.date || null,
        size: item.size || rawLength,
        attachments: attachments.map((a) => ({
          name: a.name,
          size: a.size,
          contentType: a.contentType,
          type: a.type,
          status: a.status,
          note: a.note,
          inline: Boolean(a.inline),
        })),
        contentStatus: status,
        contentNote: note,
        webLink: item.webLink || null,
        occurrences,
        terms: [...new Set(matches.map((m) => m.term))],
        matches,
      });
      if (options.deleteMatches) {
        this.pendingDeletes.push({ recordId: this.seq, messageId: item.id, internetMessageId: message.messageId, folder: item.folder, subject: message.subject });
      }
      if (this.pendingResults.length >= 50) this.flushResults();
    } finally {
      this.inFlight.delete(key);
    }
  }
}
