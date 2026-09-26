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
    deleted: 0, // excluídas na análise ("analisar e excluir")
    deleteMissing: 0, // já não existiam na hora da exclusão
    deleteChanged: 0,
    deleteErrors: 0,
  };
}

const mb = (bytes) => `${Math.round((bytes / 1048576) * 10) / 10} MB`;

export class MailScanner {
  /**
   * config: { sources: [conexões com secrets], terms: [...], options: {...}, endpoints? }
   * emit(message): recebe { type: 'log'|'progress'|'results'|'errors'|'done', ... }
   */
  constructor(config, emit, { connectorFactory = createConnector } = {}) {
    this.sources = config.sources || [];
    this.options = { ...MAIL_DEFAULT_OPTIONS, ...(config.options || {}) };
    this.matcher = new Matcher(config.terms || [], { maxSamples: this.options.maxSamples, guard: createGuard(config.regexTimeoutMs || 30000) });
    this.abort = new AbortController();
    this.emit = emit;
    this.connectorFactory = connectorFactory;
    this.endpoints = config.endpoints || {};
    this.stats = newMailStats(this.sources.length);
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
    for (const { term, error } of this.matcher.invalid) this.log('warn', `Termo ignorado "${term.value}": ${error}`);
    if (this.matcher.size === 0) this.log('warn', 'Nenhum termo válido nas listas selecionadas.');
    this.log('info', `Análise de e-mail iniciada com ${this.matcher.size} termo(s) em ${this.sources.length} conexão(ões).`);
    for (const source of this.sources) {
      if (this.cancelled) break;
      await this.scanSource(source);
      this.stats.sourcesDone++;
      this.progress(true);
    }
    this.flushResults();
    this.flushErrors();
    this.current = null;
    const s = this.stats;
    const seconds = Math.round((Date.now() - started) / 1000);
    this.log(
      'info',
      `${this.cancelled ? 'Análise cancelada' : 'Análise concluída'} em ${seconds}s: ${s.messagesSeen} mensagem(ns) verificadas em ${s.mailboxesDone} caixa(s), ${s.messagesMatched} com ocorrências.`,
    );
    this.emit({ type: 'done', stats: { ...s }, cancelled: this.cancelled });
    return s;
  }

  async scanSource(source) {
    this.current = { source: source.name, mailbox: null, folder: null, path: source.name };
    this.progress(true);
    let connector;
    let mailboxes;
    try {
      connector = this.connectorFactory(source, {
        signal: this.abort.signal,
        endpoints: this.endpoints,
        log: (level, message) => this.log(level, `${source.name}: ${message}`),
      });
      mailboxes = await connector.mailboxes();
    } catch (err) {
      if (this.cancelled) return;
      this.error(source.name, err);
      this.stats.gaps++;
      this.log('error', `Conexão "${source.name}" indisponível: ${friendlyError(err)}`);
      return;
    }
    this.stats.mailboxesTotal += mailboxes.length;
    this.log('info', `Conexão "${source.name}": ${mailboxes.length} caixa(s) a analisar.`);
    try {
      for (const mailbox of mailboxes) {
        if (this.cancelled) break;
        await this.scanMailbox(connector, source, mailbox);
        this.stats.mailboxesDone++;
        this.progress(true);
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
    try {
      const items = connector.messages(mailbox, {
        since: this.since,
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
        if (item.error) {
          const where = [mailbox.address, item.folder, item.id].filter(Boolean).join(' › ');
          this.error(where, item.id ? `Falha ao baixar a mensagem: ${friendlyError(item.error)}` : friendlyError(item.error));
        } else {
          await this.processMessage(source, mailbox, item);
        }
        this.progress();
      }
      if (!this.cancelled) this.log('info', `Caixa ${mailbox.address}: ${this.stats.messagesSeen - before} mensagem(ns) verificadas.`);
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
      const items = list.map((p) => {
        countDeletion(this.stats, status);
        if (status === 'failed') this.error(`${mailbox.address} › ${p.folder}`, `Falha ao excluir a mensagem "${p.subject || '(sem assunto)'}": ${r.error}`);
        const item = `${mailbox.address} › ${p.folder} › ${p.subject || '(sem assunto)'}`;
        return deletionEvent(p.recordId, { status, error: r.ok ? null : r.error, note: r.note }, { mode: 'auto', method, by: this.deletedBy, item });
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
    if (!failure && !source.allowDelete) failure = 'A exclusão foi desativada no cadastro da conexão durante a análise.';
    for (const id of [...byId.keys()]) record(id, { ok: false, error: failure || 'O servidor não confirmou a exclusão.' });
    this.progress(true);
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
