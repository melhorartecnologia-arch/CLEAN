// Motor da análise do Microsoft Teams: percorre as conversas (canais das equipes e chats) das
// conexões escolhidas e procura os termos das listas de referência no assunto, no corpo das
// mensagens e nos nomes e no conteúdo dos arquivos anexados. O relatório traz apenas as mensagens
// com ocorrências, com a equipe/canal ou o chat, o autor e a data.
import { Matcher } from '../scan/matcher.js';
import { createGuard, countDeletion } from '../scan/scanner.js';
import { extractBuffer, DEFAULT_LIMITS } from '../scan/extractors/index.js';
import { friendlyError, withTimeout } from '../scan/errors.js';
import { deletionEvent } from '../scan/delete.js';
import { validDate } from '../mail/common.js';
import { TeamsConnector } from './connector.js';

const MESSAGE_TIMEOUT = 5 * 60 * 1000;

export const TEAMS_DEFAULT_OPTIONS = {
  scanChannels: true,
  scanChats: true,
  includeReplies: true,
  checkSubject: true,
  checkBody: true,
  checkAttachmentNames: true,
  checkAttachments: true, // conteúdo dos arquivos anexados
  receivedAfter: null,
  maxMessageSizeMB: 50, // anexos maiores: só o início é baixado
  concurrency: 4,
  maxSamples: 3,
  deleteMatches: false,
};

export function newTeamsStats(sourcesTotal = 0) {
  return {
    sourcesTotal,
    sourcesDone: 0,
    conversationsTotal: 0,
    conversationsDone: 0,
    channels: 0,
    chats: 0,
    messagesSeen: 0,
    messagesMatched: 0,
    occurrences: 0,
    attachmentsSeen: 0,
    attachmentsAnalyzed: 0,
    attachmentsEncrypted: 0,
    attachmentsUnsupported: 0,
    attachmentsSkippedSize: 0,
    attachmentsErrors: 0,
    bytesDownloaded: 0,
    errors: 0,
    gaps: 0, // conexões, equipes, usuários ou conversas que não puderam ser lidos
    deleted: 0,
    deleteMissing: 0,
    deleteChanged: 0,
    deleteErrors: 0,
    deleteSkipped: 0, // mensagens de chat (exclusão não oferecida pelo Graph)
  };
}

const plural = (n, one, many) => `${n} ${n === 1 ? one : many}`;

export class TeamsScanner {
  constructor(config, emit, { connectorFactory = (source, options) => new TeamsConnector(source, options) } = {}) {
    this.sources = config.sources || [];
    this.options = { ...TEAMS_DEFAULT_OPTIONS, ...(config.options || {}) };
    this.matcher = new Matcher(config.terms || [], { maxSamples: this.options.maxSamples, guard: createGuard(config.regexTimeoutMs || 30000) });
    this.abort = new AbortController();
    this.emit = emit;
    this.connectorFactory = connectorFactory;
    this.endpoints = config.endpoints || {};
    this.stats = newTeamsStats(this.sources.length);
    this.cancelled = false;
    this.seq = 0;
    this.inFlight = new Set();
    this.pendingErrors = [];
    this.pendingResults = [];
    this.pendingDeletes = [];
    this.lastProgress = 0;
    this.current = null;
    this.maxBytes = Math.max(1, Number(this.options.maxMessageSizeMB) || 50) * 1048576;
    this.limits = { ...DEFAULT_LIMITS, maxBytes: this.maxBytes };
    this.messageTimeoutMs = config.messageTimeoutMs || MESSAGE_TIMEOUT;
    this.since = validDate(this.options.receivedAfter);
    this.deletedBy = config.startedBy || null;
  }

  cancel() {
    this.cancelled = true;
    this.abort.abort(new Error('Análise cancelada.'));
  }

  /** A exclusão foi desligada no cadastro durante a análise: nada mais é excluído daquela conexão. */
  revokeDeletion({ kind, id, reason, all = false }) {
    if (all) {
      const active = this.sources.filter((s) => s.allowDelete);
      for (const s of active) s.allowDelete = false;
      if (active.length && this.options.deleteMatches) this.log('warn', `Exclusão automática desativada: ${reason}. A análise continua sem excluir.`);
      return;
    }
    if (kind !== 'teams' && kind !== 'mail') return;
    for (const s of this.sources) {
      if (s.id !== id || !s.allowDelete) continue;
      s.allowDelete = false;
      if (this.options.deleteMatches) this.log('warn', `Exclusão automática desativada para "${s.name}": ${reason || 'o cadastro da conexão foi alterado'}.`);
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
    this.log('info', `Análise do Teams iniciada com ${this.matcher.size} termo(s) em ${plural(this.sources.length, 'conexão', 'conexões')}.`);
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
      `${this.cancelled ? 'Análise cancelada' : 'Análise concluída'} em ${seconds}s: ${s.messagesSeen} mensagem(ns) em ${s.conversationsDone} conversa(s), ${s.messagesMatched} com ocorrências.`,
    );
    this.emit({ type: 'done', stats: { ...s }, cancelled: this.cancelled });
    return s;
  }

  async scanSource(source) {
    this.current = { source: source.name, mailbox: null, folder: null, path: source.name };
    this.progress(true);
    let connector;
    try {
      const grantId = source.graph?.account?.grantId || null;
      connector = this.connectorFactory(source, {
        signal: this.abort.signal,
        endpoints: this.endpoints,
        log: (level, message) => this.log(level, `${source.name}: ${message}`),
        onRefreshToken: grantId ? (refreshToken) => this.emit({ type: 'credentials', sourceId: source.id, grantId, refreshToken }) : undefined,
      });
    } catch (err) {
      if (this.cancelled) return;
      this.error(source.name, err);
      this.stats.gaps++;
      this.log('error', `Conexão "${source.name}" indisponível: ${friendlyError(err)}`);
      return;
    }
    try {
      for await (const conv of connector.conversations()) {
        if (this.cancelled) break;
        if (conv.error) {
          this.error(`${source.name} › ${conv.path}`, conv.error);
          this.stats.gaps++;
          continue;
        }
        this.stats.conversationsTotal++;
        this.stats[conv.kind === 'channel' ? 'channels' : 'chats']++;
        await this.scanConversation(connector, source, conv);
        this.stats.conversationsDone++;
        this.progress(true);
        const failure = connector.auth?.failure;
        if (failure) {
          this.error(source.name, `A conexão parou: ${friendlyError(failure)}`);
          this.stats.gaps++;
          this.log('error', `Conexão "${source.name}": ${friendlyError(failure)}`);
          break;
        }
      }
    } catch (err) {
      if (this.cancelled) return;
      this.error(source.name, err);
      this.stats.gaps++;
      this.log('error', `Falha na conexão "${source.name}": ${friendlyError(err)}`);
    } finally {
      await connector.close?.();
    }
  }

  async scanConversation(connector, source, conv) {
    this.current = { source: source.name, mailbox: conv.path, folder: conv.path, path: conv.path };
    this.pendingDeletes = [];
    try {
      for await (const item of connector.messages(conv, { since: this.since })) {
        if (this.cancelled) break;
        await this.processMessage(connector, source, conv, item);
        this.progress();
      }
    } catch (err) {
      if (this.cancelled) return;
      this.error(`${source.name} › ${conv.path}`, err);
      this.stats.gaps++;
      this.log('error', `Falha na conversa "${conv.path}": ${friendlyError(err)}`);
    } finally {
      const pending = this.pendingDeletes.splice(0);
      if (!this.cancelled && pending.length) await this.deleteMessages(connector, source, conv, pending);
    }
  }

  async processMessage(connector, source, conv, item) {
    const { options, matcher, stats } = this;
    stats.messagesSeen++;
    const key = `${conv.path} › ${item.id}`;
    this.inFlight.add(key);
    try {
      const groups = [];
      if (options.checkSubject && item.subject) groups.push({ location: 'subject', segments: [{ text: item.subject, label: 'Assunto' }] });
      if (options.checkBody && item.text) groups.push({ location: 'body', segments: [{ text: item.text, label: 'Mensagem', lines: true }] });
      if (options.checkAttachmentNames && item.attachments.length) {
        groups.push({ location: 'attachmentName', segments: item.attachments.map((a) => ({ text: a.name, label: 'Nome do anexo' })) });
      }
      const attachmentsOut = item.attachments.map((a) => ({ name: a.name, contentType: a.contentType, status: a.reference ? 'not-requested' : 'unsupported', note: null, size: 0 }));
      if (options.checkAttachments) {
        const segs = [];
        for (let i = 0; i < item.attachments.length; i++) {
          const a = item.attachments[i];
          if (!a.reference) continue; // cartões e anexos sem arquivo não são baixados
          const out = attachmentsOut[i];
          stats.attachmentsSeen++;
          try {
            const res = await withTimeout(connector.downloadAttachment(a, { maxBytes: this.maxBytes, signal: this.abort.signal }), this.messageTimeoutMs, 'Tempo esgotado ao baixar o anexo.');
            stats.bytesDownloaded += res.data?.length || 0;
            out.size = res.size || res.data?.length || 0;
            const extracted = await withTimeout(extractBuffer(res.data, { name: a.name, limits: this.limits }), this.messageTimeoutMs, 'Tempo esgotado ao ler o anexo.');
            out.status = res.truncated ? 'partial' : extracted.status;
            if (['ok', 'partial', 'empty'].includes(extracted.status) || res.truncated) stats.attachmentsAnalyzed++;
            else if (extracted.status === 'encrypted') stats.attachmentsEncrypted++;
            else if (extracted.status === 'skipped-size') stats.attachmentsSkippedSize++;
            else stats.attachmentsUnsupported++;
            if (res.truncated) out.note = 'Anexo grande: apenas o início foi lido.';
            for (const seg of extracted.segments || []) segs.push({ text: seg.text, label: `Anexo "${a.name}"`, lines: seg.lines });
          } catch (err) {
            stats.attachmentsErrors++;
            out.status = 'error';
            out.note = friendlyError(err);
          }
        }
        if (segs.length) groups.push({ location: 'attachment', segments: segs });
      }
      const matches = matcher.matchGroups(groups).flat();
      if (matcher.timedOut) this.error(key, 'Tempo limite ao procurar as expressões regulares nesta mensagem; resultado parcial.');
      if (matches.length === 0) return;

      const occurrences = matches.reduce((sum, m) => sum + m.count, 0);
      stats.messagesMatched++;
      stats.occurrences += occurrences;
      const anyBad = attachmentsOut.some((a) => ['error', 'skipped-size', 'partial', 'encrypted'].includes(a.status));
      const recordId = ++this.seq;
      this.pendingResults.push({
        id: recordId,
        kind: 'teams',
        sourceId: source.id,
        sourceName: source.name,
        sourceType: 'graph',
        scopeKind: conv.kind,
        team: conv.teamName || '',
        teamId: conv.teamId || '',
        channel: conv.channelName || '',
        channelId: conv.channelId || '',
        chatType: conv.chatType || '',
        membershipType: conv.membershipType || '',
        folder: conv.path,
        messageId: item.id,
        replyTo: item.replyTo || null,
        subject: item.subject || '',
        from: item.from || '',
        fromId: item.fromId || '',
        date: item.date || null,
        edited: item.edited && item.edited !== item.date ? item.edited : null,
        size: item.size || 0,
        attachments: attachmentsOut,
        contentStatus: anyBad ? 'partial' : 'ok',
        contentNote: anyBad ? 'Um ou mais anexos não puderam ser lidos por completo (veja a lista de anexos).' : null,
        webUrl: item.webUrl || null,
        occurrences,
        terms: [...new Set(matches.map((m) => m.term))],
        matches,
      });
      if (options.deleteMatches && source.allowDelete && conv.kind === 'channel') {
        this.pendingDeletes.push({ recordId, messageId: item.id, replyTo: item.replyTo || null, conv, subject: item.subject || '' });
      } else if (options.deleteMatches && source.allowDelete && conv.kind === 'chat') {
        stats.deleteSkipped++;
      }
      if (this.pendingResults.length >= 50) this.flushResults();
    } finally {
      this.inFlight.delete(key);
    }
  }

  /** Exclusão automática (softDelete) das mensagens de canal encontradas em uma conversa. */
  async deleteMessages(connector, source, conv, pending) {
    this.flushResults();
    this.current = { source: source.name, mailbox: conv.path, folder: conv.path, path: `${conv.path} › excluindo ${pending.length} mensagem(ns)` };
    this.progress(true);
    this.log('info', `${conv.path}: excluindo ${plural(pending.length, 'mensagem', 'mensagens')} do Teams (recuperável).`);
    const byId = new Map(pending.map((p) => [p.recordId, p]));
    const record = (recordId, r) => {
      const p = byId.get(recordId);
      if (!p) return;
      byId.delete(recordId);
      const status = r.ok ? 'deleted' : r.missing ? 'missing' : 'failed';
      countDeletion(this.stats, status);
      if (status === 'failed') this.error(conv.path, `Falha ao excluir a mensagem "${p.subject || '(sem assunto)'}": ${r.error}`);
      const item = `${conv.path} › ${p.subject || '(sem assunto)'}`;
      this.emit({ type: 'deletions', items: [deletionEvent(recordId, { status, error: r.ok ? null : r.error, note: r.note }, { mode: 'auto', method: 'teams', by: this.deletedBy, item })] });
      this.progress();
    };
    let failure = source.allowDelete ? null : 'A exclusão não está permitida nesta conexão.';
    if (!failure) {
      try {
        const items = pending.map((p) => ({ recordId: p.recordId, messageId: p.messageId, replyTo: p.replyTo, conv: p.conv }));
        await connector.deleteMessages(items, { signal: null, onResult: record, shouldStop: () => this.cancelled || !source.allowDelete });
      } catch (err) {
        failure = friendlyError(err);
        this.log('error', `Falha ao excluir mensagens de "${conv.path}": ${failure}`);
      }
    }
    if (this.cancelled) return;
    for (const recordId of [...byId.keys()]) record(recordId, { ok: false, error: failure || 'O servidor não confirmou a exclusão.' });
    this.progress(true);
  }
}
