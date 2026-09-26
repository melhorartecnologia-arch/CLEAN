// Servidores IMAP: Exchange local, Zimbra, Dovecot, provedores de hospedagem, Gmail com senha de app
// etc. Cada caixa tem o seu login; a senha pode ser individual ou uma senha padrão da conexão (útil
// com contas de serviço, ex.: "DOMINIO\servico\caixa" no Exchange ou "caixa*mestre" no Dovecot).
import { ApiError } from './http.js';
import { folderMatcher, addressMatcher, normalizeAddress, deletionItems, normalizeMessageId } from './common.js';

// Marcadores do Gmail exibidos como pasta (quando o servidor é o Gmail).
const GMAIL_LABELS = { '\\Inbox': 'Caixa de entrada', '\\Sent': 'Enviados', '\\Draft': 'Rascunhos', '\\Spam': 'Spam', '\\Trash': 'Lixeira' };
const HIDDEN_GMAIL_LABELS = new Set(['\\Important', '\\Starred', '\\Muted', '\\Chat']);

const BATCH_MESSAGES = 50;
const BATCH_BYTES = 32 * 1024 * 1024;
// Mensagens por comando na exclusão: conjuntos de UIDs muito longos passam do tamanho máximo de
// comando de alguns servidores (10 KB no Exchange).
const DELETE_CHUNK = 200;

/** A biblioteca IMAP é carregada só quando usada: sem ela (npm install não executado após uma atualização) o resto do CLEAN funciona. */
let imapFlowClass = null;
async function loadImapFlow() {
  if (!imapFlowClass) {
    try {
      imapFlowClass = (await import('imapflow')).ImapFlow;
    } catch {
      throw new ApiError('O componente IMAP não está instalado. Na pasta do CLEAN, execute "npm install --omit=dev" e reinicie o servidor.');
    }
  }
  return imapFlowClass;
}

const TLS_ERRORS = new Set([
  'DEPTH_ZERO_SELF_SIGNED_CERT',
  'SELF_SIGNED_CERT_IN_CHAIN',
  'UNABLE_TO_VERIFY_LEAF_SIGNATURE',
  'UNABLE_TO_GET_ISSUER_CERT_LOCALLY',
  'CERT_HAS_EXPIRED',
  'ERR_TLS_CERT_ALTNAME_INVALID',
]);

/** Traduz erros de conexão e autenticação IMAP. */
export function imapError(err, host = '') {
  if (err instanceof ApiError) return err;
  const code = err?.code || '';
  const where = host ? ` (${host})` : '';
  if (err?.authenticationFailed) {
    const detail = err.responseText ? `: ${err.responseText}` : '';
    return new ApiError(`Usuário ou senha recusados pelo servidor IMAP${detail}.`, { code: 'AUTH' });
  }
  if (TLS_ERRORS.has(code)) {
    return new ApiError(
      `Certificado do servidor${where} não é confiável (${code}). Para um servidor interno com certificado próprio, marque "Aceitar certificado não confiável".`,
      { code },
    );
  }
  const known = {
    ENOTFOUND: `Servidor${where} não encontrado (DNS).`,
    EAI_AGAIN: `Servidor${where} não encontrado (DNS).`,
    ECONNREFUSED: `Conexão recusada${where}: confira o servidor, a porta e o firewall.`,
    ECONNRESET: `A conexão${where} foi interrompida pelo servidor.`,
    ETIMEDOUT: `Tempo esgotado ao conectar${where}.`,
    CONNECT_TIMEOUT: `Tempo esgotado ao conectar${where}: confira o servidor, a porta e o firewall.`,
    GREETING_TIMEOUT: `O servidor${where} não respondeu como IMAP: confira a porta e o tipo de segurança (SSL/TLS ou STARTTLS).`,
    ETIMEOUT: `Tempo esgotado aguardando o servidor IMAP${where}.`,
    NoConnection: `Sem conexão com o servidor IMAP${where}.`,
  };
  if (known[code]) return new ApiError(known[code], { code });
  if (/wrong version number|packet length too long|unknown protocol/i.test(err?.message || '')) {
    return new ApiError(`Falha no TLS${where}: confira a porta e o tipo de segurança (993 = SSL/TLS; 143 = STARTTLS).`, { code: 'TLS' });
  }
  const text = err?.responseText || err?.message || String(err);
  return new ApiError(`Erro no servidor IMAP${where}: ${text}`, { code });
}

/** Conjunto de UIDs no formato IMAP ("1:5,8,10:12"), montado em tempo linear. */
export function packUids(uids) {
  const sorted = [...new Set(uids)].sort((a, b) => a - b);
  const parts = [];
  for (let i = 0; i < sorted.length; ) {
    let j = i;
    while (j + 1 < sorted.length && sorted[j + 1] === sorted[j] + 1) j++;
    parts.push(j > i ? `${sorted[i]}:${sorted[j]}` : String(sorted[i]));
    i = j + 1;
  }
  return parts.join(',');
}

function batches(list) {
  const out = [];
  let current = [];
  let bytes = 0;
  for (const item of list) {
    if (current.length && (current.length >= BATCH_MESSAGES || bytes + item.size > BATCH_BYTES)) {
      out.push(current);
      current = [];
      bytes = 0;
    }
    current.push(item);
    bytes += item.size;
  }
  if (current.length) out.push(current);
  return out;
}

function labelsText(labels) {
  if (!labels || !labels.size) return '';
  return [...labels]
    .filter((l) => !HIDDEN_GMAIL_LABELS.has(l))
    .map((l) => GMAIL_LABELS[l] || l)
    .join('; ');
}

export class ImapConnector {
  constructor(source, { signal, log = () => {} } = {}) {
    this.source = source;
    this.imap = source.imap || {};
    this.signal = signal;
    this.log = log;
  }

  async mailboxes() {
    const excluded = addressMatcher(this.source.excludeMailboxes);
    return (this.source.mailboxes || []).filter((m) => !excluded(m.address)).map((m) => ({ address: m.address, login: m.login || m.address, name: '' }));
  }

  password(mailbox) {
    const secrets = this.source.secrets || {};
    const value = secrets.passwords?.[normalizeAddress(mailbox.address)] || secrets.defaultPassword;
    if (!value) throw new ApiError(`Senha não informada para ${mailbox.address}.`);
    return value;
  }

  /** Abre uma sessão na caixa. signal: interrompe a conexão (padrão: o da análise). */
  async connect(mailbox, { signal = this.signal } = {}) {
    const { host, port, security, allowSelfSigned } = this.imap;
    const ImapFlow = await loadImapFlow();
    const client = new ImapFlow({
      host,
      port: Number(port) || (security === 'tls' ? 993 : 143),
      secure: security === 'tls',
      doSTARTTLS: security === 'starttls' ? true : security === 'none' ? false : undefined,
      auth: { user: mailbox.login || mailbox.address, pass: this.password(mailbox) },
      tls: { rejectUnauthorized: !allowSelfSigned, minVersion: 'TLSv1.2' },
      logger: false,
      disableAutoIdle: true,
      connectionTimeout: 30000,
      greetingTimeout: 20000,
      socketTimeout: 10 * 60 * 1000,
      clientInfo: { name: 'CLEAN' },
    });
    client.on('error', () => {}); // erros de conexão também chegam pelas promessas
    const onAbort = () => client.close();
    signal?.addEventListener('abort', onAbort, { once: true });
    const dispose = async () => {
      signal?.removeEventListener('abort', onAbort);
      try {
        if (client.usable) await client.logout();
      } catch {
        // ignora
      }
      client.close();
    };
    try {
      await client.connect();
    } catch (err) {
      await dispose();
      if (signal?.aborted) throw signal.reason;
      throw imapError(err, host);
    }
    return { client, dispose };
  }

  /** Pastas selecionáveis, sem as excluídas. No Gmail, "Todos os e-mails" já contém os marcadores. */
  async folders(client, { includeTrash = true, includeJunk = false } = {}) {
    const list = (await client.list()).filter((f) => !f.flags.has('\\Noselect') && !f.flags.has('\\NonExistent'));
    const all = list.find((f) => f.specialUse === '\\All');
    let selected = all ? list.filter((f) => f === all || f.specialUse === '\\Junk' || f.specialUse === '\\Trash') : list;
    const excluded = folderMatcher(this.source.excludeFolders);
    const display = (f) => (f.delimiter ? f.path.split(f.delimiter).join('/') : f.path);
    selected = selected.filter((f) => {
      if (f.specialUse === '\\Trash' && !includeTrash) return false;
      if (f.specialUse === '\\Junk' && !includeJunk) return false;
      return !excluded(display(f));
    });
    return selected
      .map((f) => ({ path: f.path, display: display(f), inbox: f.specialUse === '\\Inbox' || f.path.toUpperCase() === 'INBOX', all: f === all }))
      .sort((a, b) => Number(b.inbox) - Number(a.inbox) || a.display.localeCompare(b.display));
  }

  /**
   * Mensagens da caixa, pasta por pasta. As pequenas são baixadas em lotes; as maiores que maxBytes,
   * uma a uma e apenas até o limite. Entrega { folder, id, raw, truncated, size, receivedAt } ou
   * { folder, id, error } (pasta que não pôde ser aberta).
   */
  async *messages(mailbox, { since = null, before = null, headersOnly = false, includeTrash = true, includeJunk = false, maxBytes = 50 * 1048576, onFolder } = {}) {
    const { client, dispose } = await this.connect(mailbox);
    try {
      const gmail = client.capabilities?.has?.('X-GM-EXT-1');
      for (const folder of await this.folders(client, { includeTrash, includeJunk })) {
        if (this.signal?.aborted) return;
        onFolder?.(folder.display);
        let lock;
        try {
          lock = await client.getMailboxLock(folder.path, { readOnly: true });
        } catch (err) {
          if (this.signal?.aborted) throw this.signal.reason;
          yield { folder: folder.display, id: null, raw: null, error: imapError(err, this.imap.host) };
          continue;
        }
        try {
          if (!client.mailbox?.exists) continue;
          const validity = String(client.mailbox.uidValidity ?? '');
          const meta = [];
          for await (const m of client.fetch('1:*', { uid: true, size: true, internalDate: true, labels: Boolean(gmail && folder.all), envelope: headersOnly }, { uid: true })) {
            // O filtro por data é feito aqui, e não com SEARCH SINCE: a lista de UIDs de uma busca
            // pode passar do tamanho máximo de comando do servidor (10 KB no Exchange).
            if (since && m.internalDate instanceof Date && m.internalDate < since) continue;
            // "Antes de" (retenção): sem data conhecida, a mensagem não entra.
            if (before && !(m.internalDate instanceof Date && m.internalDate < before)) continue;
            meta.push({ uid: m.uid, size: Number(m.size) || 0, date: m.internalDate, labels: m.labels, envelope: m.envelope });
          }
          if (headersOnly) {
            // Retenção: só os dados do envelope (sem baixar a mensagem).
            for (const info of meta) {
              const env = info.envelope || {};
              const from = env.from?.[0];
              yield {
                folder: (folder.all && labelsText(info.labels)) || folder.display,
                id: `${folder.path}:${validity}:${info.uid}`,
                size: info.size,
                receivedAt: info.date instanceof Date && !Number.isNaN(info.date.getTime()) ? info.date.toISOString() : null,
                subject: env.subject || '',
                from: from ? { name: from.name || '', address: from.address || '' } : null,
                internetMessageId: env.messageId || null,
                headersOnly: true,
              };
            }
            continue;
          }
          const small = meta.filter((m) => m.size <= maxBytes);
          const large = meta.filter((m) => m.size > maxBytes);
          const make = (info, source, truncated) => ({
            folder: (folder.all && labelsText(info.labels)) || folder.display,
            id: `${folder.path}:${validity}:${info.uid}`,
            raw: source || Buffer.alloc(0),
            truncated,
            size: info.size,
            receivedAt: info.date instanceof Date && !Number.isNaN(info.date.getTime()) ? info.date.toISOString() : null,
          });
          for (const batch of batches(small)) {
            if (this.signal?.aborted) return;
            const byUid = new Map(batch.map((m) => [m.uid, m]));
            const items = [];
            for await (const m of client.fetch(packUids(batch.map((b) => b.uid)), { uid: true, source: true }, { uid: true })) items.push(m);
            for (const m of items) if (byUid.has(m.uid)) yield make(byUid.get(m.uid), m.source, false);
          }
          for (const info of large) {
            if (this.signal?.aborted) return;
            const m = await client.fetchOne(String(info.uid), { uid: true, source: { start: 0, maxLength: maxBytes } }, { uid: true });
            // O Exchange informa um tamanho estimado: só está cortada se veio até o limite.
            if (m) yield make(info, m.source, (m.source?.length || 0) >= maxBytes);
          }
        } catch (err) {
          if (this.signal?.aborted) throw this.signal.reason;
          yield { folder: folder.display, id: null, raw: null, error: imapError(err, this.imap.host) };
          if (!client.usable) return;
        } finally {
          lock.release();
        }
      }
    } finally {
      await dispose();
    }
  }

  async test() {
    const boxes = await this.mailboxes();
    if (boxes.length === 0) return { ok: false, message: 'Informe ao menos uma caixa de e-mail.', details: [] };
    const details = [];
    for (const box of boxes.slice(0, 3)) {
      let session;
      try {
        session = await this.connect(box);
        const folders = await this.folders(session.client, { includeTrash: true, includeJunk: true });
        const status = await session.client.status('INBOX', { messages: true }).catch(() => null);
        details.push(`${box.address}: login OK, ${folders.length} pasta(s)${status ? `, ${status.messages} mensagem(ns) na caixa de entrada` : ''}.`);
      } catch (err) {
        return { ok: false, message: `${box.address}: ${imapError(err, this.imap.host).message}`, details };
      } finally {
        await session?.dispose();
      }
    }
    return { ok: true, message: 'Conexão IMAP funcionando.', details };
  }

  /**
   * Exclui mensagens (ids no formato "pasta:uidvalidity:uid" gerado na análise).
   * - 'permanent': marca como excluída e expurga só essas mensagens (UID EXPUNGE, extensão UIDPLUS;
   *   sem ela a exclusão é recusada, porque um EXPUNGE comum apagaria também as outras mensagens
   *   marcadas como excluídas na pasta);
   * - 'trash': move para a pasta Lixeira do servidor (MOVE; sem ele, cópia conferida + UID EXPUNGE).
   *   Mensagens que já estão na Lixeira ficam lá.
   * No Gmail, a exclusão definitiva move para a Lixeira e expurga de lá (expurgar de outra pasta só
   * tiraria o marcador, conforme as configurações de IMAP da conta).
   * Antes, confere se a mensagem ainda é a mesma da análise (Message-ID); depois, confere se ela
   * saiu da pasta. items: ids ou { id, messageId }. options: signal (padrão: o da análise; a
   * análise usa null para que as exclusões em andamento terminem mesmo ao cancelar),
   * onResult(id, resultado) e shouldStop() (não começa outros lotes).
   * Retorna Map(id → { ok, missing?, error?, note? }).
   */
  async deleteMessages(mailbox, items, mode = 'permanent', { signal = this.signal, onResult, shouldStop } = {}) {
    const results = new Map();
    const done = (id, result) => {
      if (results.has(id)) return;
      results.set(id, result);
      onResult?.(id, result);
    };
    const groups = new Map();
    for (const item of deletionItems(items)) {
      const m = /^(.*):([^:]*):(\d+)$/.exec(String(item.id));
      if (!m) {
        done(item.id, { ok: false, error: 'Identificador de mensagem inválido.' });
        continue;
      }
      const group = groups.get(m[1]) || { validity: m[2], items: [] };
      group.items.push({ ...item, uid: Number(m[3]) });
      groups.set(m[1], group);
    }
    if (groups.size === 0) return results;
    const failRest = (error) => {
      for (const group of groups.values()) for (const item of group.items) done(item.id, { ok: false, error });
    };
    const { client, dispose } = await this.connect(mailbox, { signal });
    try {
      const has = (cap) => Boolean(client.capabilities?.has?.(cap));
      const ctx = { mode, gmail: has('X-GM-EXT-1'), canMove: has('MOVE'), uidplus: has('UIDPLUS'), trash: null, done, shouldStop, toPurge: [] };
      if (!ctx.uidplus && (mode === 'permanent' || !ctx.canMove)) {
        failRest(
          'O servidor IMAP não oferece a exclusão seletiva (extensão UIDPLUS): a exclusão foi recusada para não apagar outras mensagens marcadas como excluídas na pasta. Exclua pelo programa de e-mail ou use a opção "Mover para a Lixeira", se o servidor oferecer MOVE.',
        );
        return results;
      }
      if (mode === 'trash' || ctx.gmail) {
        ctx.trash = (await client.list()).find((f) => f.specialUse === '\\Trash')?.path || null;
        if (!ctx.trash) {
          failRest('O servidor não tem uma pasta Lixeira identificada.');
          return results;
        }
      }
      for (const [folder, group] of groups) {
        if (shouldStop?.()) break;
        await this.deleteInFolder(client, folder, group, ctx);
      }
      // Gmail, exclusão definitiva: expurga da Lixeira as mensagens movidas para lá.
      if (ctx.toPurge.length) await this.purgeTrash(client, ctx);
    } finally {
      await dispose();
    }
    return results;
  }

  async deleteInFolder(client, folder, group, ctx) {
    let lock = null;
    try {
      lock = await client.getMailboxLock(folder);
      if (String(client.mailbox?.uidValidity ?? '') !== group.validity) {
        for (const item of group.items) {
          ctx.done(item.id, { ok: false, error: 'A pasta foi recriada no servidor depois da análise (UIDVALIDITY diferente): a mensagem não pode ser localizada com segurança e não foi excluída.' });
        }
        return;
      }
      for (let i = 0; i < group.items.length; i += DELETE_CHUNK) {
        if (ctx.shouldStop?.()) return;
        await this.deleteChunk(client, folder, group.items.slice(i, i + DELETE_CHUNK), ctx);
      }
    } catch (err) {
      const message = imapError(err, this.imap.host).message;
      for (const item of group.items) ctx.done(item.id, { ok: false, error: message });
    } finally {
      lock?.release();
    }
  }

  /** Um lote de mensagens de uma pasta (já aberta para escrita). */
  async deleteChunk(client, folder, items, ctx) {
    // 1. Quais ainda existem e se são as mesmas da análise.
    const found = new Map();
    for await (const msg of client.fetch(packUids(items.map((i) => i.uid)), { uid: true, flags: true, envelope: true }, { uid: true })) {
      found.set(msg.uid, { messageId: normalizeMessageId(msg.envelope?.messageId), deleted: Boolean(msg.flags?.has('\\Deleted')) });
    }
    const targets = [];
    for (const item of items) {
      const info = found.get(item.uid);
      const expected = normalizeMessageId(item.messageId);
      if (!info) ctx.done(item.id, { ok: false, missing: true, error: 'Mensagem não encontrada (já excluída ou movida).' });
      else if (expected && info.messageId && expected !== info.messageId) {
        ctx.done(item.id, { ok: false, error: 'A mensagem nesta posição da pasta não é a mesma da análise (Message-ID diferente): nada foi excluído.' });
      } else targets.push({ ...item, wasDeleted: info.deleted });
    }
    if (targets.length === 0) return;
    const range = packUids(targets.map((t) => t.uid));
    const inTrash = ctx.trash && folder === ctx.trash;
    if (ctx.mode === 'trash' && inTrash) {
      for (const t of targets) ctx.done(t.id, { ok: true, note: 'já estava na lixeira' });
      return;
    }

    // 2. A exclusão.
    let error = null;
    let moved = null;
    if ((ctx.mode === 'trash' || ctx.gmail) && !inTrash) {
      if (ctx.canMove) {
        moved = await client.messageMove(range, ctx.trash, { uid: true });
        if (!moved) error = 'O servidor recusou mover a mensagem para a Lixeira.';
      } else {
        // Sem MOVE: copia, confere a cópia e só então expurga da pasta original.
        moved = await client.messageCopy(range, ctx.trash, { uid: true });
        if (!moved) error = 'O servidor recusou copiar a mensagem para a Lixeira: nada foi excluído.';
        else error = await this.expunge(client, range);
      }
    } else {
      error = await this.expunge(client, range);
    }

    // 3. Confere o que saiu da pasta.
    const still = new Set();
    for await (const msg of client.fetch(range, { uid: true }, { uid: true })) still.add(msg.uid);
    const restore = targets.filter((t) => still.has(t.uid) && !t.wasDeleted).map((t) => t.uid);
    if (restore.length) await client.messageFlagsRemove(packUids(restore), ['\\Deleted'], { uid: true }).catch(() => false);
    for (const t of targets) {
      if (still.has(t.uid)) {
        ctx.done(t.id, { ok: false, error: error || 'O servidor não excluiu a mensagem (ela continua na pasta).' });
      } else if (ctx.gmail && ctx.mode === 'permanent' && !inTrash) {
        const trashUid = moved?.uidMap?.get?.(t.uid);
        if (trashUid) ctx.toPurge.push({ ...t, trashUid });
        else ctx.done(t.id, { ok: true, note: 'movida para a lixeira; o Gmail não informou a posição dela lá para a exclusão definitiva' });
      } else {
        ctx.done(t.id, { ok: true });
      }
    }
  }

  /** Marca como excluídas e expurga só estas mensagens (UID EXPUNGE). Retorna o erro, se houver. */
  async expunge(client, range) {
    if (!(await client.messageFlagsAdd(range, ['\\Deleted'], { uid: true }))) {
      return 'O servidor recusou marcar a mensagem como excluída (sem permissão de escrita na pasta?).';
    }
    if (!(await client.messageDelete(range, { uid: true }))) return 'O servidor recusou expurgar a mensagem.';
    return null;
  }

  /** Gmail: expurga da Lixeira as mensagens que a exclusão definitiva moveu para lá. */
  async purgeTrash(client, ctx) {
    const note = (reason) => `movida para a lixeira; a exclusão definitiva falhou (${reason}) e ela será apagada pelo Gmail em 30 dias`;
    let lock = null;
    try {
      lock = await client.getMailboxLock(ctx.trash);
      for (let i = 0; i < ctx.toPurge.length; i += DELETE_CHUNK) {
        const chunk = ctx.toPurge.slice(i, i + DELETE_CHUNK);
        const range = packUids(chunk.map((t) => t.trashUid));
        const error = await this.expunge(client, range);
        const still = new Set();
        for await (const msg of client.fetch(range, { uid: true }, { uid: true })) still.add(msg.uid);
        for (const t of chunk) ctx.done(t.id, still.has(t.trashUid) ? { ok: true, note: note(error || 'a mensagem continua na lixeira') } : { ok: true });
      }
    } catch (err) {
      const reason = imapError(err, this.imap.host).message;
      for (const t of ctx.toPurge) ctx.done(t.id, { ok: true, note: note(reason) });
    } finally {
      lock?.release();
    }
  }

  async close() {}
}
