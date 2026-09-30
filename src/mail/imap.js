// Servidores IMAP: Exchange local, Zimbra, Dovecot, provedores de hospedagem, Gmail com senha de app
// etc. Cada caixa tem o seu login; a senha pode ser individual ou uma senha padrão da conexão (útil
// com contas de serviço, ex.: "DOMINIO\servico\caixa" no Exchange ou "caixa*mestre" no Dovecot).
// No Exchange Online e no Outlook.com (que não aceitam mais senha), o login é por OAuth 2.0 da
// Microsoft (XOAUTH2): com a conta conectada ou com um aplicativo (segredo do cliente ou certificado).
import { ApiError, sleep } from './http.js';
import { folderMatcher, addressMatcher, normalizeAddress, deletionItems, normalizeMessageId } from './common.js';
import { MicrosoftAuth, MICROSOFT_IMAP_HOSTS } from '../cloud/microsoft-auth.js';

// Marcadores do Gmail exibidos como pasta (quando o servidor é o Gmail).
const GMAIL_LABELS = { '\\Inbox': 'Caixa de entrada', '\\Sent': 'Enviados', '\\Draft': 'Rascunhos', '\\Spam': 'Spam', '\\Trash': 'Lixeira' };
const HIDDEN_GMAIL_LABELS = new Set(['\\Important', '\\Starred', '\\Muted', '\\Chat']);

const BATCH_MESSAGES = 50;
const BATCH_BYTES = 32 * 1024 * 1024;
// Mensagens por comando na exclusão: conjuntos de UIDs muito longos passam do tamanho máximo de
// comando de alguns servidores (10 KB no Exchange).
const DELETE_CHUNK = 200;
// Login OAuth: o Exchange Online encerra a sessão IMAP quando o token vence. Cada sessão começa com um
// token válido por ao menos 30 minutos e é trocada por outra (aberta antes de fechar a atual) quando
// faltam menos de 5 — entre as pastas e entre os lotes de mensagens de uma pasta.
const OAUTH_SESSION_MIN_MS = 30 * 60 * 1000;
const OAUTH_RENEW_MS = 5 * 60 * 1000;
// Sessão que cai (token vencido, rede): a leitura continua numa nova sessão, de onde parou. Só desiste
// da caixa depois de tantas falhas seguidas sem ler nenhuma mensagem (quedas ou reconexões recusadas),
// com uma pausa crescente entre as tentativas.
const MAX_FAILED_SESSIONS = 5;
const reconnectPause = (failures) => Math.min(1000 * 2 ** (failures - 1), 15000);

/** Problema da própria pasta (e não da sessão): não adianta tentar de novo em outra sessão. */
const folderProblem = (message) => Object.assign(new ApiError(message), { folderProblem: true });

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

/** Servidor IMAP da Microsoft (Exchange Online ou Outlook.com). */
export const isMicrosoftImapHost = (host) => MICROSOFT_IMAP_HOSTS.has(String(host || '').trim().toLowerCase());

/**
 * Onde o CLEAN aceita enviar o token OAuth da Microsoft: só aos servidores IMAP da Microsoft, com
 * TLS e certificado verificado (endpoints.microsoftImap troca a lista nos testes). Devolve o motivo
 * para recusar, ou null.
 */
export function imapOAuthProblem(imap = {}, endpoints = {}) {
  const test = endpoints.microsoftImap || null;
  const host = String(imap.host || '').toLowerCase();
  if (!isMicrosoftImapHost(host) && !test?.hosts?.includes(host)) {
    return `O login OAuth da Microsoft só é usado nos servidores IMAP da Microsoft (${[...MICROSOFT_IMAP_HOSTS].join(', ')}): em outro servidor, o token daria acesso à caixa a quem o recebesse.`;
  }
  if ((imap.security === 'none' && !test?.insecure) || imap.allowSelfSigned) return 'Com o login OAuth da Microsoft, use SSL/TLS (ou STARTTLS) com o certificado do servidor verificado.';
  return null;
}

/** Traduz erros de conexão e autenticação IMAP. oauth: login OAuth da Microsoft (XOAUTH2). */
export function imapError(err, host = '', { oauth = false } = {}) {
  if (err instanceof ApiError) return err;
  const code = err?.code || '';
  const where = host ? ` (${host})` : '';
  if (err?.authenticationFailed) {
    const detail = err.responseText ? `: ${err.responseText}` : '';
    if (oauth) {
      return new ApiError(
        `O servidor IMAP da Microsoft recusou o login OAuth${detail}. Confira se o IMAP está habilitado na caixa (Microsoft 365: centro de administração › Usuários › a pessoa › Email › Gerenciar aplicativos de email) e se a conta conectada tem acesso a ela (em caixas de outras pessoas, a permissão de Acesso Total). Com um aplicativo, a entidade de serviço dele precisa estar registrada no Exchange Online e ter Acesso Total à caixa (veja o LEIA-ME).`,
        { code: 'AUTH' },
      );
    }
    if (isMicrosoftImapHost(host)) {
      return new ApiError(
        `A Microsoft recusou o login com senha${detail}. O Exchange Online e o Outlook.com não aceitam mais senha no IMAP (a autenticação básica foi desativada): edite a conexão e, em "Autenticação", escolha "OAuth 2.0 da Microsoft" — ou use o tipo de conexão Microsoft 365.`,
        { code: 'AUTH' },
      );
    }
    return new ApiError(`Usuário ou senha recusados pelo servidor IMAP${detail}.`, { code: 'AUTH' });
  }
  if (oauth && /Unsupported authentication mechanism/i.test(err?.message || '')) {
    return new ApiError(`O servidor IMAP${where} não oferece o login por OAuth (XOAUTH2).`, { code: 'AUTH' });
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

/** Endereço do envelope IMAP ({ name, address, ... }) no formato { name, address }. */
const envAddress = (a) => ({ name: a?.name || '', address: a?.address || '' });

/** Se a estrutura do corpo (BODYSTRUCTURE) tem alguma parte anexada (com disposição ou nome de arquivo). */
function attachmentInStructure(node) {
  if (!node) return false;
  const disp = String(node.disposition || '').toLowerCase();
  if (disp === 'attachment' || (disp !== 'inline' && node.dispositionParameters?.filename)) return true;
  return (node.childNodes || []).some(attachmentInStructure);
}

export class ImapConnector {
  /**
   * options: signal, log(level, message), endpoints (troca os endereços, nos testes) e
   * onRefreshToken(token) (login OAuth com a conta conectada: novo token de atualização a gravar).
   */
  constructor(source, { signal, log = () => {}, endpoints = {}, onRefreshToken } = {}) {
    this.source = source;
    this.imap = source.imap || {};
    this.signal = signal;
    this.log = log;
    this.endpoints = endpoints;
    this.oauth = this.imap.auth === 'oauth';
    if (this.oauth) this.auth = new MicrosoftAuth(source.graph || {}, source.secrets || {}, { login: endpoints.graphLogin, signal, onRefreshToken });
  }

  error(err) {
    return imapError(err, this.imap.host, { oauth: this.oauth });
  }

  async mailboxes() {
    const excluded = addressMatcher(this.source.excludeMailboxes);
    return (this.source.mailboxes || []).filter((m) => !excluded(m.address)).map((m) => ({ address: m.address, login: m.login || m.address, name: '' }));
  }

  /**
   * Catálogo de contas (listagem): o IMAP não expõe a lista de contas do servidor, então traz apenas
   * as caixas cadastradas nesta conexão, com um aviso.
   */
  async *directory() {
    const excluded = addressMatcher(this.source.excludeMailboxes);
    const note = 'O IMAP não tem um catálogo de contas do servidor: a listagem traz apenas as caixas cadastradas nesta conexão.';
    for (const m of this.source.mailboxes || []) {
      if (excluded(m.address)) continue;
      yield {
        address: m.address,
        name: '',
        login: m.login || m.address,
        aliases: [],
        enabled: null,
        type: 'Caixa cadastrada (IMAP)',
        licensed: null,
        created: null,
        lastActivity: null,
        department: '',
        title: '',
        location: '',
        phone: '',
        orgUnit: '',
        admin: null,
        note,
      };
    }
  }

  /**
   * Login da caixa: usuário e senha, ou o e-mail da caixa e um token de acesso da Microsoft (OAuth,
   * XOAUTH2) válido por ao menos 30 minutos.
   */
  async credentials(mailbox) {
    if (!this.oauth) return { user: mailbox.login || mailbox.address, pass: this.password(mailbox) };
    const problem = imapOAuthProblem(this.imap, this.endpoints);
    if (problem) throw new ApiError(problem);
    return { user: mailbox.address, accessToken: await this.auth.token('imap', { minValidityMs: OAUTH_SESSION_MIN_MS }) };
  }

  password(mailbox) {
    const secrets = this.source.secrets || {};
    const value = secrets.passwords?.[normalizeAddress(mailbox.address)] || secrets.defaultPassword;
    if (!value) throw new ApiError(`Senha não informada para ${mailbox.address}.`);
    return value;
  }

  /**
   * Abre uma sessão na caixa: { client, dispose, expiresAt } (expiresAt: quando vence o token da
   * sessão OAuth; Infinity com senha). signal: interrompe a conexão (padrão: o da análise).
   */
  async connect(mailbox, { signal = this.signal } = {}) {
    const { host, port, security, allowSelfSigned } = this.imap;
    const ImapFlow = await loadImapFlow();
    const auth = await this.credentials(mailbox);
    const client = new ImapFlow({
      host,
      port: Number(port) || (security === 'tls' ? 993 : 143),
      secure: security === 'tls',
      doSTARTTLS: security === 'starttls' ? true : security === 'none' ? false : undefined,
      auth,
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
      throw this.error(err);
    }
    return { client, dispose, expiresAt: auth.accessToken ? this.auth.expiresAt('imap') : Infinity };
  }

  /** Pastas selecionáveis, sem as excluídas. No Gmail, "Todos os e-mails" já contém os marcadores. */
  async folders(client, { includeTrash = true, includeJunk = false } = {}) {
    const list = (await client.list()).filter((f) => !f.flags.has('\\Noselect') && !f.flags.has('\\NonExistent'));
    const all = list.find((f) => f.specialUse === '\\All');
    let selected = all ? list.filter((f) => f === all || f.specialUse === '\\Junk' || f.specialUse === '\\Trash') : list;
    const excluded = folderMatcher(this.source.excludeFolders);
    const display = (f) => (f.delimiter ? f.path.split(f.delimiter).join('/') : f.path);
    // inTrash: a Lixeira e as subpastas dela (sem a Lixeira, as subpastas também ficam de fora, como no
    // Microsoft 365).
    const trash = list.find((f) => f.specialUse === '\\Trash');
    const inTrash = (f) => Boolean(trash) && (f === trash || Boolean(trash.delimiter && f.path.startsWith(`${trash.path}${trash.delimiter}`)));
    selected = selected.filter((f) => {
      if (!includeTrash && inTrash(f)) return false;
      if (f.specialUse === '\\Junk' && !includeJunk) return false;
      return !excluded(display(f));
    });
    return selected
      .map((f) => ({ path: f.path, display: display(f), inbox: f.specialUse === '\\Inbox' || f.path.toUpperCase() === 'INBOX', all: f === all, inTrash: inTrash(f) }))
      .sort((a, b) => Number(b.inbox) - Number(a.inbox) || a.display.localeCompare(b.display));
  }

  /**
   * Mensagens da caixa, pasta por pasta. As pequenas são baixadas em lotes; as maiores que maxBytes,
   * uma a uma e apenas até o limite. Entrega { folder, id, raw, truncated, size, receivedAt } ou
   * { folder, id, error } (pasta que não pôde ser lida). Se a sessão cair (ex.: o Exchange Online a
   * encerra quando o token OAuth vence), a leitura continua numa nova sessão, de onde parou.
   */
  async *messages(mailbox, { since = null, before = null, headersOnly = false, fullHeaders = false, includeTrash = true, includeJunk = false, maxBytes = 50 * 1048576, onFolder } = {}) {
    let session = await this.connect(mailbox);
    let read = 0; // mensagens entregues pela sessão atual
    let failures = 0; // falhas seguidas sem ler nenhuma mensagem (quedas e reconexões recusadas)
    // Troca planejada (token perto de vencer): a nova sessão é aberta antes de fechar a atual; se não
    // abrir, a atual continua até cair (a queda é tratada como as outras).
    const renewDue = () => session.expiresAt - Date.now() < OAUTH_RENEW_MS && !session.keep;
    const renew = async () => {
      try {
        const next = await this.connect(mailbox);
        await session.dispose();
        session = next;
        read = 0;
        // Tokens que valem menos que a margem (política do locatário): trocar de novo não adianta.
        if (renewDue()) session.keep = true;
      } catch (err) {
        if (this.signal?.aborted) throw this.signal.reason;
        session.keep = true;
      }
    };
    // A sessão caiu: abre outra (com pausas crescentes entre as tentativas). Devolve o erro que faz
    // desistir da caixa, ou null.
    const reconnect = async () => {
      // A sessão que caiu depois de ler mensagens é reaberta na hora; quedas sem nenhuma leitura contam
      // como falhas.
      failures = read > 0 ? 0 : failures + 1;
      await session.dispose();
      for (;;) {
        if (failures > MAX_FAILED_SESSIONS) {
          return new ApiError(`A sessão com o servidor IMAP caiu (ou não pôde ser reaberta) ${MAX_FAILED_SESSIONS} vezes seguidas sem ler nenhuma mensagem: o restante da caixa não foi lido.`);
        }
        if (failures > 0) await sleep(reconnectPause(failures), this.signal);
        try {
          session = await this.connect(mailbox);
          read = 0;
          if (renewDue()) session.keep = true;
          return null;
        } catch (err) {
          if (this.signal?.aborted) throw this.signal.reason;
          failures++;
        }
      }
    };
    try {
      const gmail = session.client.capabilities?.has?.('X-GM-EXT-1');
      for (const folder of await this.folders(session.client, { includeTrash, includeJunk })) {
        if (this.signal?.aborted) return;
        onFolder?.(folder.display);
        // Mensagens já entregues desta pasta (uma nova sessão continua sem repeti-las).
        const state = { validity: null, delivered: new Set() };
        for (;;) {
          if (session.client.usable && renewDue()) await renew();
          if (!session.client.usable) {
            const stop = await reconnect();
            if (stop) {
              yield { folder: folder.display, id: null, raw: null, error: stop };
              return;
            }
          }
          let failure = null;
          try {
            for await (const item of this.folderMessages(session.client, folder, state, { gmail, since, before, headersOnly, fullHeaders, maxBytes, renewDue })) {
              read++;
              failures = 0;
              yield item;
            }
          } catch (err) {
            if (this.signal?.aborted) throw this.signal.reason;
            failure = err;
          }
          if (!failure) break;
          // Token perto de vencer no meio da pasta, ou a sessão caiu: continua a mesma pasta em outra.
          if (failure.renewSession || (!session.client.usable && !failure.folderProblem)) continue;
          yield { folder: folder.display, id: null, raw: null, error: this.error(failure) };
          break;
        }
      }
    } finally {
      await session.dispose();
    }
  }

  /**
   * As mensagens de uma pasta ainda não entregues (state.delivered), na sessão dada. Um erro da
   * sessão (conexão perdida) é lançado para messages() continuar em outra; um problema da própria
   * pasta vem com folderProblem. renewDue(): o token da sessão está para vencer — entre um lote e
   * outro, a leitura para (renewSession) e continua numa nova sessão.
   */
  async *folderMessages(client, folder, state, { gmail, since, before, headersOnly, fullHeaders = false, maxBytes, renewDue = () => false }) {
    const pause = () => Object.assign(new Error('Renovar a sessão IMAP.'), { renewSession: true });
    let lock;
    try {
      lock = await client.getMailboxLock(folder.path, { readOnly: true });
    } catch (err) {
      if (this.signal?.aborted) throw this.signal.reason;
      if (!client.usable) throw err;
      throw Object.assign(this.error(err), { folderProblem: true });
    }
    try {
      if (!client.mailbox?.exists) return;
      const validity = String(client.mailbox.uidValidity ?? '');
      // Numa nova sessão, os números (UIDs) só continuam valendo se a pasta não foi recriada.
      if (state.validity !== null && state.validity !== validity) {
        throw folderProblem('A pasta foi recriada no servidor durante a análise (UIDVALIDITY diferente): o restante dela não foi lido.');
      }
      state.validity = validity;
      const meta = [];
      for await (const m of client.fetch('1:*', { uid: true, size: true, internalDate: true, labels: Boolean(gmail && folder.all), envelope: headersOnly, bodyStructure: fullHeaders }, { uid: true })) {
        if (state.delivered.has(m.uid)) continue;
        // O filtro por data é feito aqui, e não com SEARCH SINCE: a lista de UIDs de uma busca
        // pode passar do tamanho máximo de comando do servidor (10 KB no Exchange).
        if (since && m.internalDate instanceof Date && m.internalDate < since) continue;
        // "Antes de" (retenção): sem data conhecida, a mensagem não entra.
        if (before && !(m.internalDate instanceof Date && m.internalDate < before)) continue;
        meta.push({ uid: m.uid, size: Number(m.size) || 0, date: m.internalDate, labels: m.labels, envelope: m.envelope, bodyStructure: m.bodyStructure });
      }
      const deliver = (uid, item) => {
        state.delivered.add(uid);
        return item;
      };
      if (headersOnly) {
        // Retenção e listagem: só os dados do envelope (sem baixar a mensagem). A listagem
        // (fullHeaders) usa também os destinatários e a data de envio do envelope.
        for (const info of meta) {
          const env = info.envelope || {};
          const from = env.from?.[0];
          yield deliver(info.uid, {
            folder: (folder.all && labelsText(info.labels)) || folder.display,
            id: `${folder.path}:${validity}:${info.uid}`,
            size: info.size,
            receivedAt: info.date instanceof Date && !Number.isNaN(info.date.getTime()) ? info.date.toISOString() : null,
            subject: env.subject || '',
            from: from ? { name: from.name || '', address: from.address || '' } : null,
            internetMessageId: env.messageId || null,
            inTrash: folder.inTrash,
            headersOnly: true,
            ...(fullHeaders
              ? {
                  to: (env.to || []).map(envAddress),
                  cc: (env.cc || []).map(envAddress),
                  sent: env.date instanceof Date && !Number.isNaN(env.date.getTime()) ? env.date.toISOString() : null,
                  hasAttachments: info.bodyStructure ? attachmentInStructure(info.bodyStructure) : null,
                }
              : {}),
          });
        }
        return;
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
        if (renewDue()) throw pause();
        const byUid = new Map(batch.map((m) => [m.uid, m]));
        const items = [];
        for await (const m of client.fetch(packUids(batch.map((b) => b.uid)), { uid: true, source: true }, { uid: true })) items.push(m);
        for (const m of items) if (byUid.has(m.uid)) yield deliver(m.uid, make(byUid.get(m.uid), m.source, false));
      }
      for (const info of large) {
        if (this.signal?.aborted) return;
        if (renewDue()) throw pause();
        const m = await client.fetchOne(String(info.uid), { uid: true, source: { start: 0, maxLength: maxBytes } }, { uid: true });
        // O Exchange informa um tamanho estimado: só está cortada se veio até o limite.
        if (m) yield deliver(info.uid, make(info, m.source, (m.source?.length || 0) >= maxBytes));
      }
    } finally {
      lock.release();
    }
  }

  async test() {
    const boxes = await this.mailboxes();
    if (boxes.length === 0) return { ok: false, message: 'Informe ao menos uma caixa de e-mail.', details: [] };
    const details = [];
    if (this.oauth) {
      try {
        await this.auth.token('imap');
      } catch (err) {
        return { ok: false, message: this.error(err).message, details };
      }
      const account = this.source.graph?.account;
      details.push(`Login OAuth da Microsoft: token obtido (${account ? `conta conectada ${account.username || account.address}` : this.auth.mode === 'certificate' ? 'aplicativo com certificado' : 'aplicativo com segredo do cliente'}).`);
    }
    for (const box of boxes.slice(0, 3)) {
      let session;
      try {
        session = await this.connect(box);
        const folders = await this.folders(session.client, { includeTrash: true, includeJunk: true });
        const status = await session.client.status('INBOX', { messages: true }).catch(() => null);
        details.push(`${box.address}: login OK, ${folders.length} pasta(s)${status ? `, ${status.messages} mensagem(ns) na caixa de entrada` : ''}.`);
      } catch (err) {
        return { ok: false, message: `${box.address}: ${this.error(err).message}`, details };
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
