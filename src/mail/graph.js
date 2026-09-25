// Microsoft 365 / Exchange Online pela API Microsoft Graph, com permissões de aplicativo (sem usuário
// conectado): um registro de aplicativo no Microsoft Entra ID com Mail.Read e User.Read.All e um
// segredo do cliente. Cada mensagem é baixada no formato MIME original (com os anexos).
import { request, pool, ApiError } from './http.js';
import { SkipMailboxError, folderMatcher, addressMatcher } from './common.js';

export const GRAPH_ENDPOINTS = {
  login: 'https://login.microsoftonline.com',
  graph: 'https://graph.microsoft.com/v1.0',
};

// O Exchange Online aceita até 4 requisições simultâneas por caixa para cada aplicativo.
const MAX_CONCURRENCY = 4;

// Usuário sem caixa no Exchange Online (sem licença, caixa desativada ou hospedada localmente).
const NO_MAILBOX = new Set(['MailboxNotEnabledForRESTAPI', 'MailboxNotHostedInExchangeOnline', 'ErrorNonExistentMailbox', 'ErrorInvalidUser']);

const AAD_ERRORS = {
  700016: 'O aplicativo (ID do cliente) não foi encontrado neste locatário. Confira o ID do cliente e o ID do locatário.',
  7000215: 'Segredo do cliente inválido. Copie o VALOR do segredo (e não o "ID do segredo").',
  7000222: 'O segredo do cliente expirou. Gere um novo segredo no registro do aplicativo.',
  90002: 'Locatário não encontrado. Confira o ID do locatário (ou o domínio, ex.: empresa.onmicrosoft.com).',
  900023: 'ID do locatário inválido.',
  700023: 'O ID do locatário informado não corresponde ao aplicativo.',
  50034: 'Conta não encontrada no diretório.',
  53003: 'Acesso bloqueado por uma política de Acesso Condicional do Microsoft Entra ID.',
};

const enc = encodeURIComponent;

/** Traduz erros do Graph/Entra ID para mensagens com a providência a tomar. */
export function graphError(err) {
  if (!(err instanceof ApiError)) return err;
  const aad = /AADSTS(\d+)/.exec(err.message);
  if (aad && AAD_ERRORS[aad[1]]) return new ApiError(`${AAD_ERRORS[aad[1]]} (AADSTS${aad[1]})`, err);
  if (err.status === 403 && /Authorization_RequestDenied/i.test(err.code)) {
    return new ApiError('Permissão insuficiente para listar os usuários: conceda ao aplicativo a permissão User.Read.All (tipo Aplicativo) com consentimento do administrador.', err);
  }
  if (err.status === 403 && /AccessDenied/i.test(err.code)) {
    return new ApiError(
      'Acesso negado à caixa de correio: conceda ao aplicativo a permissão Mail.Read (tipo Aplicativo) com consentimento do administrador e confira se uma política de acesso do Exchange Online (RBAC para aplicativos ou Application Access Policy) libera esta caixa.',
      err,
    );
  }
  if (err.status === 401) return new ApiError(`Credenciais recusadas pelo Microsoft 365: ${err.message}`, err);
  const detail = err.code ? ` (${err.code})` : err.status ? ` (HTTP ${err.status})` : '';
  return new ApiError(`${err.message}${detail}`, err);
}

export class GraphConnector {
  constructor(source, { endpoints = {}, signal, log = () => {} } = {}) {
    this.source = source;
    this.endpoints = { login: endpoints.graphLogin || GRAPH_ENDPOINTS.login, graph: endpoints.graph || GRAPH_ENDPOINTS.graph };
    this.signal = signal;
    this.log = log;
    this.token = null;
    this.tokenExpires = 0;
    this.tokenPromise = null;
    this.lastThrottleLog = 0;
  }

  async accessToken(force = false) {
    if (!force && this.token && Date.now() < this.tokenExpires - 120000) return this.token;
    const { tenantId, clientId } = this.source.graph || {};
    const secret = this.source.secrets?.clientSecret;
    if (!tenantId || !clientId || !secret) throw new ApiError('Informe o ID do locatário, o ID do cliente e o segredo do cliente.');
    this.tokenPromise ||= request(`${this.endpoints.login}/${enc(tenantId)}/oauth2/v2.0/token`, {
      method: 'POST',
      form: { client_id: clientId, client_secret: secret, scope: 'https://graph.microsoft.com/.default', grant_type: 'client_credentials' },
      signal: this.signal,
      retries: 3,
    })
      .then((res) => {
        if (!res?.access_token) throw new ApiError('O Microsoft Entra ID não devolveu um token de acesso.');
        this.token = res.access_token;
        this.tokenExpires = Date.now() + (Number(res.expires_in) || 3600) * 1000;
        return this.token;
      })
      .catch((err) => {
        throw graphError(err);
      })
      .finally(() => {
        this.tokenPromise = null;
      });
    return this.tokenPromise;
  }

  /** Chamada autenticada ao Graph (renova o token uma vez se ele for recusado). */
  async api(pathOrUrl, options = {}) {
    const url = pathOrUrl.startsWith('/') ? `${this.endpoints.graph}${pathOrUrl}` : pathOrUrl;
    for (let attempt = 0; ; attempt++) {
      const token = await this.accessToken(attempt > 0);
      try {
        return await request(url, {
          ...options,
          signal: this.signal,
          headers: { Authorization: `Bearer ${token}`, ...(options.headers || {}) },
          onRetry: ({ wait, error }) => this.throttled(wait, error),
        });
      } catch (err) {
        if (err instanceof ApiError && err.status === 401 && attempt === 0) continue;
        throw graphError(err);
      }
    }
  }

  throttled(wait, error) {
    if (Date.now() - this.lastThrottleLog < 60000) return;
    this.lastThrottleLog = Date.now();
    const reason = error.status === 429 ? 'limite de requisições do Microsoft 365 atingido' : `falha temporária (${error.message})`;
    this.log('warn', `Microsoft 365: ${reason}; nova tentativa em ${Math.round(wait / 1000)}s.`);
  }

  /** Próxima página. O link precisa apontar para o próprio Graph (o token não vai para outro endereço). */
  next(page) {
    const link = page?.['@odata.nextLink'];
    if (!link) return null;
    if (!link.startsWith(`${this.endpoints.graph}/`)) throw new ApiError('Resposta inesperada do Microsoft Graph (paginação para outro endereço).');
    return link;
  }

  /** Caixas a analisar: todas as do locatário (usuários com e-mail) ou as da lista. */
  async mailboxes() {
    const excluded = addressMatcher(this.source.excludeMailboxes);
    if (this.source.scope !== 'all') {
      return (this.source.mailboxes || []).map((m) => ({ address: m.address, name: '' })).filter((m) => !excluded(m.address));
    }
    const out = [];
    let url = '/users?$select=id,displayName,mail,userPrincipalName&$top=999';
    while (url) {
      const page = await this.api(url);
      for (const u of page?.value || []) if (u.mail && !excluded(u.mail)) out.push({ address: u.mail, name: u.displayName || '', id: u.id });
      url = this.next(page);
    }
    return out.sort((a, b) => a.address.localeCompare(b.address));
  }

  /** Identificador do usuário dono da caixa (o endereço pode ser diferente do nome de logon). */
  async resolveUser(mailbox) {
    if (mailbox.id) return { id: mailbox.id, name: mailbox.name };
    const select = '$select=id,displayName,mail';
    try {
      const u = await this.api(`/users/${enc(mailbox.address)}?${select}`);
      return { id: u.id, name: u.displayName || '' };
    } catch (err) {
      if (err.status !== 404) throw err;
    }
    const quoted = mailbox.address.replace(/'/g, "''");
    const byMail = await this.api(`/users?$filter=${enc(`mail eq '${quoted}'`)}&${select}`);
    let u = byMail?.value?.[0];
    if (!u) {
      const byAlias = await this.api(`/users?$filter=${enc(`proxyAddresses/any(x:x eq 'smtp:${quoted}')`)}&$count=true&${select}`, { headers: { ConsistencyLevel: 'eventual' } });
      u = byAlias?.value?.[0];
    }
    if (!u) throw new ApiError(`A caixa ${mailbox.address} não foi encontrada no Microsoft 365.`, { status: 404 });
    return { id: u.id, name: u.displayName || '' };
  }

  async wellKnownFolder(userId, name) {
    try {
      return (await this.api(`/users/${enc(userId)}/mailFolders/${name}?$select=id`))?.id || null;
    } catch (err) {
      if (err.status === 404 && !NO_MAILBOX.has(err.code)) return null;
      throw err;
    }
  }

  /** Pastas da caixa (com subpastas), sem as excluídas pelas opções e pelos padrões da conexão. */
  async folders(userId, { includeTrash = true, includeJunk = false } = {}) {
    const skip = new Set();
    try {
      const trash = await this.wellKnownFolder(userId, 'deleteditems');
      const junk = await this.wellKnownFolder(userId, 'junkemail');
      const sync = await this.wellKnownFolder(userId, 'syncissues');
      if (!includeTrash && trash) skip.add(trash);
      if (!includeJunk && junk) skip.add(junk);
      if (sync) skip.add(sync); // registros de sincronização do Outlook
    } catch (err) {
      if (NO_MAILBOX.has(err.code) || err.status === 404) throw new SkipMailboxError('usuário sem caixa de correio no Exchange Online (sem licença, desativada ou local).');
      throw err;
    }
    const excluded = folderMatcher(this.source.excludeFolders);
    const fields = '$select=id,displayName,childFolderCount,totalItemCount&$top=250';
    const out = [];
    const walk = async (url, parent) => {
      for (let next = url; next; ) {
        const page = await this.api(next);
        for (const f of page?.value || []) {
          const path = parent ? `${parent}/${f.displayName}` : f.displayName;
          if (skip.has(f.id) || excluded(path)) continue;
          out.push({ id: f.id, path, total: Number(f.totalItemCount) || 0 });
          if (f.childFolderCount > 0) await walk(`/users/${enc(userId)}/mailFolders/${enc(f.id)}/childFolders?${fields}`, path);
        }
        next = this.next(page);
      }
    };
    await walk(`/users/${enc(userId)}/mailFolders?${fields}`, '');
    return out;
  }

  /**
   * Mensagens da caixa, já baixadas (MIME), com até `concurrency` downloads simultâneos.
   * Entrega { folder, id, raw, truncated, size, receivedAt, webLink } ou { folder, id, error }.
   */
  async *messages(mailbox, { since = null, includeTrash = true, includeJunk = false, maxBytes = 50 * 1048576, concurrency = 4, onFolder } = {}) {
    const user = await this.resolveUser(mailbox);
    mailbox.name ||= user.name;
    const folders = await this.folders(user.id, { includeTrash, includeJunk });
    const userPath = `/users/${enc(user.id)}`;
    const filter = since ? `&$filter=${enc(`receivedDateTime ge ${since.toISOString()}`)}` : '';
    const size = `&$expand=${enc("singleValueExtendedProperties($filter=id eq 'Integer 0x0E08')")}`;
    const self = this;
    async function* list() {
      for (const folder of folders) {
        onFolder?.(folder.path);
        if (!folder.total) continue;
        let url = `${userPath}/mailFolders/${enc(folder.id)}/messages?$select=id,receivedDateTime,webLink&$top=100${size}${filter}`;
        while (url) {
          const page = await self.api(url);
          for (const m of page?.value || []) {
            yield {
              folder: folder.path,
              id: m.id,
              receivedAt: m.receivedDateTime || null,
              webLink: m.webLink || null,
              size: Number(m.singleValueExtendedProperties?.[0]?.value) || 0,
            };
          }
          url = self.next(page);
        }
      }
    }
    yield* pool(list(), Math.max(1, Math.min(concurrency, MAX_CONCURRENCY)), async (item) => {
      try {
        const res = await this.api(`${userPath}/messages/${enc(item.id)}/$value`, { type: 'buffer', maxBytes, retries: 4, headers: { Accept: '*/*' } });
        return { ...item, raw: res.data, truncated: res.truncated, size: item.size || res.size };
      } catch (err) {
        if (this.signal?.aborted) throw err;
        return { ...item, raw: null, error: err };
      }
    });
  }

  /** Teste da conexão: autenticação, listagem de usuários (se for o caso) e acesso a até 3 caixas. */
  async test() {
    const details = [];
    await this.accessToken();
    details.push('Autenticação no Microsoft Entra ID: OK.');
    let boxes;
    if (this.source.scope === 'all') {
      const page = await this.api('/users?$select=id,displayName,mail&$top=50');
      boxes = (page?.value || []).filter((u) => u.mail).map((u) => ({ address: u.mail, id: u.id, name: u.displayName }));
      details.push(`Listagem de usuários (User.Read.All): OK — ${boxes.length}${page?.['@odata.nextLink'] ? '+' : ''} com e-mail.`);
      if (boxes.length === 0) return { ok: false, message: 'Nenhum usuário com e-mail encontrado no locatário.', details };
    } else {
      boxes = (this.source.mailboxes || []).map((m) => ({ address: m.address }));
      if (boxes.length === 0) return { ok: false, message: 'Informe ao menos uma caixa de e-mail.', details };
    }
    let checked = 0;
    for (const box of boxes) {
      if (checked >= 3) break;
      try {
        const user = await this.resolveUser(box);
        const inbox = await this.api(`/users/${enc(user.id)}/mailFolders/inbox?$select=displayName,totalItemCount`);
        details.push(`${box.address}: ${Number(inbox?.totalItemCount) || 0} mensagem(ns) em "${inbox?.displayName || 'Caixa de Entrada'}".`);
        checked++;
      } catch (err) {
        if (this.source.scope === 'all' && (NO_MAILBOX.has(err.code) || err.status === 404)) continue; // usuário sem caixa
        return { ok: false, message: `${box.address}: ${err.message}`, details };
      }
    }
    if (checked === 0) return { ok: false, message: 'Nenhuma das caixas testadas está disponível no Exchange Online.', details };
    return { ok: true, message: 'Conexão com o Microsoft 365 funcionando.', details };
  }

  async close() {}
}
