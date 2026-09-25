// Google Workspace (Gmail) pela API do Gmail, com uma conta de serviço autorizada por "delegação em
// todo o domínio": acessa as caixas dos usuários sem as senhas deles. A lista de usuários vem da
// Admin SDK (Directory API), consultada em nome de um administrador.
import crypto from 'node:crypto';
import { request, pool, ApiError } from './http.js';
import { SkipMailboxError, folderMatcher, addressMatcher, deletionItems } from './common.js';

export const GOOGLE_ENDPOINTS = {
  token: 'https://oauth2.googleapis.com/token',
  gmail: 'https://gmail.googleapis.com/gmail/v1',
  directory: 'https://admin.googleapis.com/admin/directory/v1',
};

export const GMAIL_SCOPE = 'https://www.googleapis.com/auth/gmail.readonly';
// Exclusão: definitiva exige o escopo completo; mover para a lixeira, o de modificação.
export const GMAIL_DELETE_SCOPE = 'https://mail.google.com/';
export const GMAIL_TRASH_SCOPE = 'https://www.googleapis.com/auth/gmail.modify';
export const DIRECTORY_SCOPE = 'https://www.googleapis.com/auth/admin.directory.user.readonly';

// Downloads simultâneos por caixa: cada mensagem vem inteira no JSON (base64), então o limite
// também controla o uso de memória.
const MAX_CONCURRENCY = 4;

// Marcadores do sistema exibidos como "pasta"; os demais (não lido, estrela, categorias) não.
const SYSTEM_LABELS = { INBOX: 'Caixa de entrada', SENT: 'Enviados', DRAFT: 'Rascunhos', SPAM: 'Spam', TRASH: 'Lixeira' };

const enc = encodeURIComponent;
const b64url = (value) => Buffer.from(value).toString('base64url');

/** Assina um JWT RS256 com a chave privada da conta de serviço. */
export function signJwt(claims, privateKey) {
  const head = b64url(JSON.stringify({ alg: 'RS256', typ: 'JWT' }));
  const body = b64url(JSON.stringify(claims));
  const signature = crypto.sign('RSA-SHA256', Buffer.from(`${head}.${body}`), privateKey);
  return `${head}.${body}.${b64url(signature)}`;
}

/** Traduz erros do Google para mensagens com a providência a tomar. */
export function googleError(err, subject = '') {
  if (!(err instanceof ApiError)) return err;
  const text = `${err.code} ${err.message}`;
  if (/unauthorized_client/i.test(text)) {
    return new ApiError(
      'A conta de serviço não está autorizada para os escopos necessários. No Admin Console do Google (Segurança > Acesso e controle de dados > Controles de API > Delegação em todo o domínio), autorize o ID do cliente da conta de serviço com os escopos indicados no LEIA-ME.',
      err,
    );
  }
  if (/invalid_grant/i.test(text)) {
    if (/invalid email|user id/i.test(text)) return new ApiError(`Usuário ${subject} não encontrado no Google Workspace.`, err);
    return new ApiError(`Autorização recusada pelo Google (${err.message}). Confira a chave da conta de serviço e o relógio deste servidor.`, err);
  }
  if (/accessNotConfigured|has not been used|is disabled/i.test(text)) {
    return new ApiError('Ative a Gmail API e a Admin SDK API no projeto do Google Cloud da conta de serviço.', err);
  }
  if (err.status === 403 && /directory|admin/i.test(text)) {
    return new ApiError(`O administrador informado não tem permissão para listar os usuários (Admin SDK): ${err.message}`, err);
  }
  const detail = err.code ? ` (${err.code})` : err.status ? ` (HTTP ${err.status})` : '';
  return new ApiError(`${err.message}${detail}`, err);
}

/**
 * Monta uma mensagem MIME a partir da estrutura do formato "full" do Gmail: os textos vêm com o
 * conteúdo; os anexos, só com os cabeçalhos e marcados como não baixados (X-Clean-Omitted com o
 * código `token`, que só o conector conhece).
 */
export function payloadToMime(part, token) {
  const header = (name) => (part.headers || []).find((h) => String(h.name).toLowerCase() === name)?.value;
  const clean = (value) => String(value ?? '').replace(/[\r\n]+/g, ' ');
  const lines = (part.headers || [])
    .filter((h) => !/^(content-type|content-transfer-encoding|content-length)$/i.test(h.name))
    .map((h) => `${clean(h.name)}: ${clean(h.value)}`);
  const type = String(part.mimeType || 'text/plain').toLowerCase();
  if (type.startsWith('multipart/') && part.parts?.length) {
    const boundary = `clean-${crypto.randomUUID()}`;
    lines.push(`Content-Type: ${type}; boundary="${boundary}"`);
    const children = part.parts.map((child) => `--${boundary}\r\n${payloadToMime(child, token)}`);
    return `${lines.join('\r\n')}\r\n\r\n${children.join('\r\n')}\r\n--${boundary}--\r\n`;
  }
  lines.push(`Content-Type: ${clean(header('content-type') || type)}`);
  if (part.body?.data && !part.body?.attachmentId) {
    const b64 = Buffer.from(part.body.data, 'base64url').toString('base64').replace(/.{76}/g, '$&\r\n');
    lines.push('Content-Transfer-Encoding: base64');
    return `${lines.join('\r\n')}\r\n\r\n${b64}\r\n`;
  }
  lines.push(`X-Clean-Omitted: ${token}`, `X-Clean-Size: ${Number(part.body?.size) || 0}`);
  return `${lines.join('\r\n')}\r\n\r\n`;
}

function mailDisabled(err) {
  return err instanceof ApiError && (err.status === 400 || err.status === 412) && /failedPrecondition|mail service not enabled/i.test(`${err.code} ${err.message}`);
}

export class GmailConnector {
  constructor(source, { endpoints = {}, signal, log = () => {} } = {}) {
    this.source = source;
    this.endpoints = {
      token: endpoints.googleToken || GOOGLE_ENDPOINTS.token,
      gmail: endpoints.gmail || GOOGLE_ENDPOINTS.gmail,
      directory: endpoints.googleDirectory || GOOGLE_ENDPOINTS.directory,
    };
    this.signal = signal;
    this.log = log;
    this.tokens = new Map();
    this.lastThrottleLog = 0;
  }

  /** Token de acesso em nome de `subject` (usuário da caixa ou administrador). */
  async token(subject, scope, force = false, signal = this.signal) {
    const key = `${subject}|${scope}`;
    const cached = this.tokens.get(key);
    if (!force && cached && !cached.promise && Date.now() < cached.expires - 120000) return cached.value;
    if (cached?.promise) return cached.promise;
    const { clientEmail } = this.source.gmail || {};
    const privateKey = this.source.secrets?.privateKey;
    if (!clientEmail || !privateKey) throw new ApiError('Informe a chave da conta de serviço do Google.');
    const now = Math.floor(Date.now() / 1000);
    let assertion;
    try {
      assertion = signJwt({ iss: clientEmail, scope, aud: GOOGLE_ENDPOINTS.token, sub: subject, iat: now, exp: now + 3600 }, privateKey);
    } catch (err) {
      throw new ApiError(`Chave privada da conta de serviço inválida: ${err.message}`);
    }
    const promise = request(this.endpoints.token, {
      method: 'POST',
      form: { grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer', assertion },
      signal,
      retries: 3,
    })
      .then((res) => {
        if (!res?.access_token) throw new ApiError('O Google não devolveu um token de acesso.');
        this.tokens.set(key, { value: res.access_token, expires: Date.now() + (Number(res.expires_in) || 3600) * 1000 });
        return res.access_token;
      })
      .catch((err) => {
        this.tokens.delete(key);
        throw googleError(err, subject);
      });
    this.tokens.set(key, { promise });
    return promise;
  }

  async api(subject, scope, url, options = {}) {
    const signal = options.signal !== undefined ? options.signal : this.signal;
    for (let attempt = 0; ; attempt++) {
      const token = await this.token(subject, scope, attempt > 0, signal);
      try {
        return await request(url, {
          ...options,
          signal,
          headers: { Authorization: `Bearer ${token}`, ...(options.headers || {}) },
          onRetry: (info) => {
            this.throttled(info.wait, info.error);
            options.onRetry?.(info);
          },
        });
      } catch (err) {
        if (err instanceof ApiError && err.status === 401 && attempt === 0) continue;
        if (mailDisabled(err)) throw new SkipMailboxError('Gmail não habilitado para este usuário.');
        throw googleError(err, subject);
      }
    }
  }

  throttled(wait, error) {
    if (Date.now() - this.lastThrottleLog < 60000) return;
    this.lastThrottleLog = Date.now();
    const reason = error.status === 429 || error.status === 403 ? 'limite de requisições do Google atingido' : `falha temporária (${error.message})`;
    this.log('warn', `Google Workspace: ${reason}; nova tentativa em ${Math.round(wait / 1000)}s.`);
  }

  async listUsers(maxResults = 500, onePage = false) {
    const admin = this.source.gmail?.adminEmail;
    if (!admin) throw new ApiError('Informe o e-mail de um administrador para listar os usuários do domínio.');
    const out = [];
    let pageToken = '';
    do {
      const url = `${this.endpoints.directory}/users?customer=my_customer&maxResults=${maxResults}&projection=basic&orderBy=email${pageToken ? `&pageToken=${enc(pageToken)}` : ''}`;
      const page = await this.api(admin, DIRECTORY_SCOPE, url);
      for (const u of page?.users || []) if (u.primaryEmail) out.push({ address: u.primaryEmail, name: u.name?.fullName || '', suspended: Boolean(u.suspended) });
      pageToken = onePage ? '' : page?.nextPageToken;
    } while (pageToken);
    return out;
  }

  /** Caixas a analisar: todos os usuários do domínio (Admin SDK) ou os da lista. */
  async mailboxes() {
    const excluded = addressMatcher(this.source.excludeMailboxes);
    if (this.source.scope !== 'all') return (this.source.mailboxes || []).map((m) => ({ address: m.address, name: '' })).filter((m) => !excluded(m.address));
    return (await this.listUsers()).filter((u) => !excluded(u.address)).sort((a, b) => a.address.localeCompare(b.address));
  }

  /** Nomes de exibição dos marcadores (as "pastas" do Gmail). */
  async labels(address) {
    const res = await this.api(address, GMAIL_SCOPE, `${this.endpoints.gmail}/users/${enc(address)}/labels`);
    const map = new Map();
    for (const l of res?.labels || []) {
      if (l.type === 'system') {
        if (SYSTEM_LABELS[l.id]) map.set(l.id, SYSTEM_LABELS[l.id]);
      } else if (l.name) {
        map.set(l.id, l.name);
      }
    }
    return map;
  }

  /**
   * Mensagens da caixa (formato MIME original), com até `concurrency` downloads simultâneos.
   * A "pasta" de cada mensagem são os seus marcadores. Mensagens com marcador em uma pasta
   * ignorada são descartadas. Mensagens maiores que maxBytes não são baixadas inteiras: vêm só o
   * corpo e a lista de anexos (formato "full"), sem o conteúdo dos anexos.
   */
  async *messages(mailbox, { since = null, includeTrash = true, includeJunk = false, maxBytes = 50 * 1048576, concurrency = 4, onFolder } = {}) {
    const { address } = mailbox;
    const base = `${this.endpoints.gmail}/users/${enc(address)}`;
    const labels = await this.labels(address);
    onFolder?.('Todos os e-mails');
    const excluded = folderMatcher(this.source.excludeFolders);
    const terms = [];
    if (since) terms.push(`after:${Math.floor(since.getTime() / 1000)}`);
    const spamTrash = includeTrash || includeJunk;
    if (spamTrash && !includeJunk) terms.push('-in:spam');
    if (spamTrash && !includeTrash) terms.push('-in:trash');
    const limit = Math.max(2, Math.floor(maxBytes));
    const self = this;
    async function* list(sizeTerm, big) {
      const query = `&includeSpamTrash=${spamTrash}&q=${enc([...terms, sizeTerm].join(' '))}`;
      let pageToken = '';
      do {
        const page = await self.api(address, GMAIL_SCOPE, `${base}/messages?maxResults=500${query}${pageToken ? `&pageToken=${enc(pageToken)}` : ''}`);
        for (const m of page?.messages || []) yield { id: m.id, big };
        pageToken = page?.nextPageToken;
      } while (pageToken);
    }
    async function* all() {
      yield* list(`smaller:${limit}`, false);
      yield* list(`larger:${limit - 1}`, true);
    }
    yield* pool(all(), Math.max(1, Math.min(concurrency, MAX_CONCURRENCY)), async ({ id, big }) => {
      try {
        const msg = await this.api(address, GMAIL_SCOPE, `${base}/messages/${enc(id)}?format=${big ? 'full' : 'raw'}`, { retries: 4 });
        const names = (msg?.labelIds || []).map((l) => labels.get(l)).filter(Boolean);
        if (names.some((n) => excluded(n))) return undefined;
        const received = Number(msg?.internalDate);
        const item = {
          folder: names.join('; ') || 'Todos os e-mails',
          id,
          size: Number(msg?.sizeEstimate) || 0,
          receivedAt: Number.isFinite(received) && received > 0 ? new Date(received).toISOString() : null,
          webLink: null,
        };
        if (big) {
          const token = crypto.randomUUID();
          const mb = (n) => `${Math.round((n / 1048576) * 10) / 10} MB`;
          return {
            ...item,
            raw: Buffer.from(payloadToMime(msg?.payload || {}, token), 'utf8'),
            truncated: false,
            partial: true,
            omittedToken: token,
            note: `Mensagem com ${mb(item.size)}, acima do limite de ${mb(maxBytes)}: o corpo e os nomes dos anexos foram verificados, mas o conteúdo dos anexos não foi baixado.`,
          };
        }
        let raw = Buffer.from(msg?.raw || '', 'base64url');
        item.size ||= raw.length;
        const truncated = raw.length > maxBytes;
        if (truncated) raw = raw.subarray(0, maxBytes);
        return { ...item, raw, truncated };
      } catch (err) {
        if (this.signal?.aborted) throw err;
        return { folder: '', id, raw: null, error: err };
      }
    });
  }

  async test() {
    const details = [];
    let boxes;
    if (this.source.scope === 'all') {
      boxes = await this.listUsers(20, true);
      details.push(`Listagem de usuários (Admin SDK) como ${this.source.gmail?.adminEmail}: OK.`);
      if (boxes.length === 0) return { ok: false, message: 'Nenhum usuário encontrado no domínio.', details };
    } else {
      boxes = (this.source.mailboxes || []).map((m) => ({ address: m.address }));
      if (boxes.length === 0) return { ok: false, message: 'Informe ao menos uma caixa de e-mail.', details };
    }
    let checked = 0;
    for (const box of boxes) {
      if (checked >= 3) break;
      try {
        const profile = await this.api(box.address, GMAIL_SCOPE, `${this.endpoints.gmail}/users/${enc(box.address)}/profile`);
        details.push(`${box.address}: ${Number(profile?.messagesTotal) || 0} mensagem(ns).`);
        checked++;
      } catch (err) {
        if (err.skipMailbox && this.source.scope === 'all') continue;
        return { ok: false, message: `${box.address}: ${err.message}`, details };
      }
    }
    if (checked === 0) return { ok: false, message: 'Nenhuma das caixas testadas tem o Gmail habilitado.', details };
    return { ok: true, message: 'Conexão com o Google Workspace funcionando.', details };
  }

  /**
   * Exclui mensagens: 'permanent' = exclusão definitiva (escopo https://mail.google.com/),
   * 'trash' = move para a Lixeira (escopo gmail.modify). items: ids ou { id }. options: signal
   * (padrão: o da conexão; a análise usa null para que as exclusões em andamento terminem mesmo ao
   * cancelar), onResult(id, resultado) a cada mensagem e shouldStop() (não começa outras).
   * Retorna Map(id → { ok, missing?, error? }).
   */
  async deleteMessages(mailbox, items, mode = 'permanent', { signal = this.signal, onResult, shouldStop } = {}) {
    const scope = mode === 'trash' ? GMAIL_TRASH_SCOPE : GMAIL_DELETE_SCOPE;
    const base = `${this.endpoints.gmail}/users/${enc(mailbox.address)}`;
    const results = new Map();
    const run = pool(deletionItems(items), MAX_CONCURRENCY, async ({ id }) => {
      if (shouldStop?.()) return undefined;
      const result = await this.deleteOne(mailbox.address, scope, base, id, mode, signal);
      results.set(id, result);
      onResult?.(id, result);
      return undefined;
    });
    for await (const _ of run); // eslint-disable-line no-unused-vars
    return results;
  }

  async deleteOne(subject, scope, base, id, mode, signal) {
    // Uma tentativa interrompida (falha de rede, tempo esgotado, erro 5xx) pode ter sido feita pelo
    // servidor: nesse caso, "não encontrada" na repetição quer dizer que a exclusão funcionou.
    let uncertain = false;
    const onRetry = ({ error }) => {
      if (error.status !== 429 && error.status !== 403) uncertain = true;
    };
    try {
      if (mode === 'trash') await this.api(subject, scope, `${base}/messages/${enc(id)}/trash`, { method: 'POST', retries: 4, signal, onRetry });
      else await this.api(subject, scope, `${base}/messages/${enc(id)}`, { method: 'DELETE', retries: 4, signal, onRetry });
      return { ok: true };
    } catch (err) {
      if (err.status === 404) return uncertain ? { ok: true } : { ok: false, missing: true, error: 'Mensagem não encontrada (já excluída).' };
      // Escopo não autorizado: recusado na emissão do token (unauthorized_client) ou pela API (403
      // "insufficient authentication scopes"); os 403 de limite de uso não entram aqui.
      if (/não está autorizada/.test(err.message) || (err.status === 403 && /insufficient|scope|permission/i.test(err.message) && !/rate|quota|limit/i.test(err.message))) {
        return { ok: false, error: `Sem permissão para excluir: autorize o escopo ${scope} na delegação em todo o domínio da conta de serviço.` };
      }
      return { ok: false, error: err.message };
    }
  }

  async close() {}
}
