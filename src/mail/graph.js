// Microsoft 365 / Exchange Online pela API Microsoft Graph, com permissões de aplicativo (sem usuário
// conectado): um registro de aplicativo no Microsoft Entra ID com Mail.Read e User.Read.All e um
// segredo do cliente. Cada mensagem é baixada no formato MIME original (com os anexos).
import { request, pool, ApiError } from './http.js';
import { SkipMailboxError, folderMatcher, addressMatcher, deletionItems } from './common.js';

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

// Identificadores imutáveis: continuam válidos se a mensagem for movida de pasta depois da análise
// (necessário para a exclusão manual pelo relatório).
const IMMUTABLE_IDS = { Prefer: 'IdType="ImmutableId"' };

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
          signal: options.signal !== undefined ? options.signal : this.signal,
          headers: { Authorization: `Bearer ${token}`, ...(options.headers || {}) },
          onRetry: (info) => {
            this.throttled(info.wait, info.error);
            options.onRetry?.(info);
          },
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
  async resolveUser(mailbox, { signal } = {}) {
    if (mailbox.id) return { id: mailbox.id, name: mailbox.name };
    const select = '$select=id,displayName,mail';
    try {
      const u = await this.api(`/users/${enc(mailbox.address)}?${select}`, { signal });
      return { id: u.id, name: u.displayName || '' };
    } catch (err) {
      if (err.status !== 404) throw err;
    }
    const quoted = mailbox.address.replace(/'/g, "''");
    const byMail = await this.api(`/users?$filter=${enc(`mail eq '${quoted}'`)}&${select}`, { signal });
    let u = byMail?.value?.[0];
    if (!u) {
      const byAlias = await this.api(`/users?$filter=${enc(`proxyAddresses/any(x:x eq 'smtp:${quoted}')`)}&$count=true&${select}`, { headers: { ConsistencyLevel: 'eventual' }, signal });
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
      // Com "todas as caixas" e o acesso limitado pelo RBAC para aplicativos, as caixas fora do
      // escopo respondem 403: são ignoradas (com aviso), não contadas como erro.
      if (err.status === 403 && this.source.scope === 'all') throw new SkipMailboxError('acesso à caixa não liberado para o aplicativo (permissão Mail.Read ou escopo do RBAC para aplicativos).');
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
        try {
          while (url) {
            const page = await self.api(url, { headers: IMMUTABLE_IDS });
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
        } catch (err) {
          // Uma pasta que não pôde ser lida (removida durante a análise, erro do serviço...) vira
          // erro dela: as demais pastas da caixa continuam.
          if (self.signal?.aborted) throw err;
          yield { folder: folder.path, id: null, error: err };
        }
      }
    }
    yield* pool(list(), Math.max(1, Math.min(concurrency, MAX_CONCURRENCY)), async (item) => {
      if (item.error) return item;
      try {
        const res = await this.api(`${userPath}/messages/${enc(item.id)}/$value`, { type: 'buffer', maxBytes, retries: 4, headers: { Accept: '*/*', ...IMMUTABLE_IDS } });
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
    let skipped = null;
    for (const box of boxes) {
      if (checked >= 3) break;
      try {
        const user = await this.resolveUser(box);
        const inbox = await this.api(`/users/${enc(user.id)}/mailFolders/inbox?$select=displayName,totalItemCount`);
        details.push(`${box.address}: ${Number(inbox?.totalItemCount) || 0} mensagem(ns) em "${inbox?.displayName || 'Caixa de Entrada'}".`);
        checked++;
      } catch (err) {
        // Em "todas as caixas", usuários sem caixa ou fora do escopo do RBAC são pulados.
        if (this.source.scope === 'all' && (NO_MAILBOX.has(err.code) || err.status === 404 || err.status === 403)) {
          skipped = err;
          continue;
        }
        return { ok: false, message: `${box.address}: ${err.message}`, details };
      }
    }
    if (checked === 0) {
      const reason = skipped ? ` Último erro: ${skipped.message}` : '';
      return { ok: false, message: `Nenhuma das caixas testadas está disponível para o aplicativo.${reason}`, details };
    }
    return { ok: true, message: 'Conexão com o Microsoft 365 funcionando.', details };
  }

  /**
   * Exclui mensagens da caixa: 'permanent' = exclusão definitiva (permanentDelete: a mensagem vai
   * para a área de expurgo e some para o usuário; retenções e bloqueios de litígio continuam
   * valendo), 'trash' = move para Itens Excluídos. Exige a permissão Mail.ReadWrite.
   * items: ids ou { id }. options: signal (padrão: o da conexão; a análise usa null para que as
   * exclusões em andamento terminem mesmo ao cancelar), onResult(id, resultado) a cada mensagem e
   * shouldStop() (não começa outras). Retorna Map(id → { ok, missing?, error? }).
   */
  async deleteMessages(mailbox, items, mode = 'permanent', { signal = this.signal, onResult, shouldStop } = {}) {
    const results = new Map();
    const list = deletionItems(items);
    if (list.length === 0) return results;
    const user = await this.resolveUser(mailbox, { signal });
    const userPath = `/users/${enc(user.id)}`;
    const run = pool(list, MAX_CONCURRENCY, async ({ id }) => {
      if (shouldStop?.()) return undefined;
      const result = await this.deleteOne(userPath, id, mode, signal);
      results.set(id, result);
      onResult?.(id, result);
      return undefined;
    });
    for await (const _ of run); // eslint-disable-line no-unused-vars
    return results;
  }

  async deleteOne(userPath, id, mode, signal) {
    // Uma tentativa interrompida (falha de rede, tempo esgotado, erro 5xx) pode ter sido feita pelo
    // servidor: nesse caso, "não encontrada" na repetição quer dizer que a exclusão funcionou.
    let uncertain = false;
    const onRetry = ({ error }) => {
      if (error.status !== 429) uncertain = true;
    };
    const options = { method: 'POST', headers: IMMUTABLE_IDS, retries: 4, signal, onRetry };
    try {
      if (mode === 'trash') await this.api(`${userPath}/messages/${enc(id)}/move`, { ...options, json: { destinationId: 'deleteditems' } });
      else await this.api(`${userPath}/messages/${enc(id)}/permanentDelete`, options);
      return { ok: true };
    } catch (err) {
      if (err.status === 404 && /ErrorItemNotFound/i.test(err.code)) {
        return uncertain ? { ok: true } : { ok: false, missing: true, error: 'Mensagem não encontrada (já excluída ou movida).' };
      }
      if (err.status === 403) {
        return { ok: false, error: 'Sem permissão para excluir: conceda ao aplicativo a permissão Mail.ReadWrite (tipo Aplicativo) ou a função "Application Mail.ReadWrite" do RBAC para aplicativos.' };
      }
      return { ok: false, error: err.message };
    }
  }

  async close() {}
}
