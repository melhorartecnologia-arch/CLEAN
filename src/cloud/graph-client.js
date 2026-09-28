// Cliente do Microsoft Graph: token OAuth 2.0 do Microsoft Entra ID (aplicativo com segredo do
// cliente ou com certificado, ou conta conectada — veja microsoft-auth.js), novas tentativas em
// limites de requisição, paginação e localização de usuários. Usado pelas caixas de e-mail do
// Microsoft 365 e pelos repositórios do OneDrive e do SharePoint.
import { request, ApiError } from '../mail/http.js';
import { MicrosoftAuth, aadError } from './microsoft-auth.js';

export const GRAPH_ENDPOINTS = {
  login: 'https://login.microsoftonline.com',
  graph: 'https://graph.microsoft.com/v1.0',
};

export const enc = encodeURIComponent;

/**
 * Erros comuns do Entra ID e do Graph (credenciais, locatário, listagem de usuários). Retorna null
 * quando não há uma tradução específica.
 */
export function commonGraphError(err) {
  if (!(err instanceof ApiError)) return err;
  const aad = aadError(err);
  if (aad) return aad;
  if (err.status === 403 && /Authorization_RequestDenied/i.test(err.code)) {
    return new ApiError('Permissão insuficiente para listar os usuários: conceda ao aplicativo a permissão User.Read.All (tipo Aplicativo) com consentimento do administrador.', err);
  }
  if (err.status === 401) return new ApiError(`Credenciais recusadas pelo Microsoft 365: ${err.message}`, err);
  return null;
}

/** Mensagem do Graph com o código do erro (ex.: "Item not found (itemNotFound)"). */
export function detailedGraphError(err) {
  if (!(err instanceof ApiError)) return err;
  const detail = err.code ? ` (${err.code})` : err.status ? ` (HTTP ${err.status})` : '';
  return new ApiError(`${err.message}${detail}`, err);
}

export class GraphClient {
  /**
   * source: cadastro com graph { tenantId, clientId, auth, certificate?, account? } e secrets
   * { clientSecret | certificateKey | refreshToken } (já decifrados).
   * options: { endpoints (troca os endereços, nos testes), signal, log(level, message),
   * onRefreshToken(token) (conta conectada: novo token de atualização a gravar) }.
   */
  constructor(source, { endpoints = {}, signal, log = () => {}, onRefreshToken } = {}) {
    this.source = source;
    this.endpoints = { login: endpoints.graphLogin || GRAPH_ENDPOINTS.login, graph: endpoints.graph || GRAPH_ENDPOINTS.graph };
    this.signal = signal;
    this.log = log;
    this.auth = new MicrosoftAuth(source.graph || {}, source.secrets || {}, { login: this.endpoints.login, signal, onRefreshToken });
    this.lastThrottleLog = 0;
  }

  /** Conta conectada (permissões delegadas), e não um aplicativo. */
  get delegated() {
    return this.auth.delegated;
  }

  /** Tradução dos erros para mensagens com a providência a tomar (cada conector completa a sua). */
  translate(err) {
    return commonGraphError(err) || detailedGraphError(err);
  }

  async accessToken(force = false) {
    try {
      return await this.auth.token('graph', { force });
    } catch (err) {
      throw this.translate(err);
    }
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
        throw this.translate(err);
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

  /**
   * Todos os usuários do locatário: { id, address, mail, upn, name } (address = e-mail ou, sem ele,
   * o nome de logon).
   */
  async *allUsers() {
    let url = '/users?$select=id,displayName,mail,userPrincipalName&$top=999';
    while (url) {
      const page = await this.api(url);
      for (const u of page?.value || []) {
        const address = u.mail || u.userPrincipalName;
        if (address) yield { id: u.id, address, mail: u.mail || '', upn: u.userPrincipalName || '', name: u.displayName || '' };
      }
      url = this.next(page);
    }
  }

  /**
   * Identificador de um usuário pelo e-mail (o endereço pode ser diferente do nome de logon).
   * notFound: mensagem quando não existe (ex.: "A caixa ... não foi encontrada").
   */
  async resolveUser(mailbox, { signal, notFound = '' } = {}) {
    if (mailbox.id) return { id: mailbox.id, name: mailbox.name };
    const select = '$select=id,displayName,mail,userPrincipalName';
    const found = (u) => ({ id: u.id, name: u.displayName || '', mail: u.mail || '', upn: u.userPrincipalName || '' });
    try {
      return found(await this.api(`/users/${enc(mailbox.address)}?${select}`, { signal }));
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
    if (!u) throw new ApiError(notFound || `${mailbox.address} não foi encontrado no Microsoft 365.`, { status: 404 });
    return found(u);
  }
}
