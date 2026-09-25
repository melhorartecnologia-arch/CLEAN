// Cliente do Microsoft Graph com permissões de aplicativo (sem usuário conectado): token pelo fluxo
// "client credentials" de um registro de aplicativo no Microsoft Entra ID, novas tentativas em
// limites de requisição, paginação e localização de usuários. Usado pelas caixas de e-mail do
// Microsoft 365 e pelos repositórios do OneDrive e do SharePoint.
import { request, ApiError } from '../mail/http.js';

export const GRAPH_ENDPOINTS = {
  login: 'https://login.microsoftonline.com',
  graph: 'https://graph.microsoft.com/v1.0',
};

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

export const enc = encodeURIComponent;

/**
 * Erros comuns do Entra ID e do Graph (credenciais, locatário, listagem de usuários). Retorna null
 * quando não há uma tradução específica.
 */
export function commonGraphError(err) {
  if (!(err instanceof ApiError)) return err;
  const aad = /AADSTS(\d+)/.exec(err.message);
  if (aad && AAD_ERRORS[aad[1]]) return new ApiError(`${AAD_ERRORS[aad[1]]} (AADSTS${aad[1]})`, err);
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
   * source: cadastro com graph { tenantId, clientId } e secrets { clientSecret } (já decifrado).
   * options: { endpoints (troca os endereços, nos testes), signal, log(level, message) }.
   */
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

  /** Tradução dos erros para mensagens com a providência a tomar (cada conector completa a sua). */
  translate(err) {
    return commonGraphError(err) || detailedGraphError(err);
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
        throw this.translate(err);
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
