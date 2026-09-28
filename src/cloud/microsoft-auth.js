// Autenticação OAuth 2.0 na plataforma de identidade da Microsoft (Microsoft Entra ID), usada pelas
// conexões de e-mail do Microsoft 365 (API Microsoft Graph), pelo IMAP da Microsoft (Exchange Online
// e Outlook.com) e pelos repositórios do OneDrive e do SharePoint. Três formas:
// - 'secret': aplicativo com segredo do cliente (permissões de aplicativo, fluxo "client credentials");
// - 'certificate': aplicativo com certificado (o mesmo fluxo, com uma asserção assinada pela chave
//   privada do certificado no lugar do segredo — a forma recomendada pela Microsoft);
// - 'delegated': conta conectada (permissões delegadas): a pessoa entra com a conta Microsoft pelo
//   código de dispositivo e o CLEAN guarda o token de atualização (refresh token), que a Microsoft
//   renova a cada uso (o novo é gravado de novo, cifrado).
import crypto from 'node:crypto';
import { request, ApiError } from '../mail/http.js';

export const MICROSOFT_LOGIN = 'https://login.microsoftonline.com';

/** Formas de autenticação na Microsoft. */
export const MS_AUTH = {
  secret: 'Aplicativo com segredo do cliente',
  certificate: 'Aplicativo com certificado',
  delegated: 'Conta Microsoft conectada',
};

/** Locatários genéricos, só para a conta conectada: contas de trabalho ou escola, qualquer conta ou contas pessoais. */
export const TENANT_KEYWORDS = new Set(['organizations', 'common', 'consumers']);

/**
 * Servidores IMAP da Microsoft (Exchange Online e Outlook.com): os únicos para os quais o CLEAN envia
 * um token OAuth da Microsoft — em outro servidor, o token poderia ser usado para ler a caixa.
 */
export const MICROSOFT_IMAP_HOSTS = new Set(['outlook.office365.com', 'outlook.office.com', 'imap-mail.outlook.com']);

// Permissões de aplicativo: tudo o que foi concedido ao aplicativo naquele recurso.
const APP_SCOPES = { graph: 'https://graph.microsoft.com/.default', imap: 'https://outlook.office365.com/.default' };
const GRAPH = 'https://graph.microsoft.com/';
const OIDC = ['openid', 'profile', 'offline_access'];

/**
 * Permissões pedidas na entrada da conta (fluxo delegado). Microsoft 365: ler (ou ler e alterar, para
 * a exclusão) as mensagens da conta e das caixas compartilhadas com ela — as compartilhadas não são
 * pedidas com "consumers" e "common" (contas pessoais não têm). IMAP: acesso ao IMAP da conta.
 */
export function delegatedScopes(purpose, { write = false, tenantId = '' } = {}) {
  if (purpose === 'imap') return ['https://outlook.office.com/IMAP.AccessAsUser.All', ...OIDC];
  const personal = ['consumers', 'common'].includes(String(tenantId).toLowerCase());
  const mail = write ? 'Mail.ReadWrite' : 'Mail.Read';
  return [`${GRAPH}User.Read`, `${GRAPH}${mail}`, ...(personal ? [] : [`${GRAPH}${mail}.Shared`]), ...OIDC];
}

/** A entrada autorizou alterar as mensagens (Mail.ReadWrite): necessário para a exclusão pelo Microsoft Graph. */
export const canWriteMail = (scopes = []) => scopes.some((s) => /(^|\/)Mail\.ReadWrite$/i.test(String(s)));

const RECONNECT = 'Em Caixas de e-mail, edite a conexão e clique em "Conectar conta" para entrar de novo.';
const CERTIFICATE_EXPIRED = 'O certificado do aplicativo venceu: na conexão, gere um novo certificado, envie o arquivo .cer ao registro do aplicativo e salve.';

const AAD_ERRORS = {
  700016: 'O aplicativo (ID do cliente) não foi encontrado neste locatário. Confira o ID do cliente e o ID do locatário.',
  7000215: 'Segredo do cliente inválido. Copie o VALOR do segredo (e não o "ID do segredo").',
  7000222: 'O segredo do cliente expirou. Gere um novo segredo no registro do aplicativo.',
  90002: 'Locatário não encontrado. Confira o ID do locatário (ou o domínio, ex.: empresa.onmicrosoft.com).',
  900023: 'ID do locatário inválido.',
  700023: 'O ID do locatário informado não corresponde ao aplicativo.',
  50034: 'Conta não encontrada no diretório.',
  53003: 'Acesso bloqueado por uma política de Acesso Condicional do Microsoft Entra ID.',
  530003: 'Acesso bloqueado por uma política de Acesso Condicional do Microsoft Entra ID (dispositivo não gerenciado ou fora da conformidade).',
  700027: 'O certificado não foi reconhecido pelo aplicativo: envie o arquivo do certificado (.cer) em "Certificados e segredos › Certificados" do registro do aplicativo — o mesmo desta conexão (confira a impressão digital).',
  700024: 'A asserção assinada com o certificado ficou fora do prazo aceito: confira a data, a hora e o fuso horário do servidor do CLEAN.',
  7000218: 'O registro do aplicativo não permite a entrada de contas pelo código de dispositivo: em "Autenticação", ative "Permitir fluxos de clientes públicos" (Sim) e salve.',
  50194: 'O aplicativo aceita somente contas do próprio locatário: informe o ID do locatário (ou o domínio) no lugar de "organizations" ou "common".',
  50020: 'A conta usada na entrada não pertence ao locatário informado (nem é convidada nele).',
  65001: `A conta (ou o administrador) ainda não autorizou as permissões do aplicativo: conecte a conta de novo e aceite as permissões — ou peça ao administrador para conceder o consentimento ao aplicativo (Permissões de API). ${RECONNECT}`,
  65004: 'As permissões do aplicativo foram recusadas na entrada da conta.',
  90094: 'As permissões pedidas precisam da aprovação de um administrador: peça a ele para conceder o consentimento do administrador ao aplicativo (Permissões de API).',
  700082: `A autorização da conta conectada expirou por falta de uso (90 dias). ${RECONNECT}`,
  700084: `A autorização da conta conectada expirou. ${RECONNECT}`,
  50089: `A autorização da conta conectada expirou. ${RECONNECT}`,
  50173: `A autorização da conta conectada foi revogada (troca de senha ou encerramento das sessões). ${RECONNECT}`,
  50133: `A autorização da conta conectada foi revogada (troca de senha ou encerramento das sessões). ${RECONNECT}`,
  50076: `A política de acesso exige uma nova verificação da conta (autenticação multifator). ${RECONNECT}`,
  50079: `A conta precisa concluir o cadastro da autenticação multifator: entre uma vez pelo navegador e depois conecte a conta de novo. ${RECONNECT}`,
  50105: 'A conta não está liberada para usar o aplicativo: peça ao administrador para atribuí-la ao aplicativo (Aplicativos empresariais › Usuários e grupos).',
  50055: `A senha da conta expirou: troque a senha e conecte a conta de novo. ${RECONNECT}`,
  50057: 'A conta está desativada no Microsoft Entra ID.',
  7000112: 'O aplicativo está desativado no locatário.',
  70011: 'Permissão (escopo) não aceita para esta conta ou aplicativo.',
  500011: 'O recurso pedido não existe no locatário (o Exchange Online está disponível nele?).',
};

/** Identificadores do erro no Entra ID, para o administrador localizar a entrada nos logs de entrada. */
function traceIds(message) {
  const trace = /Trace ID:\s*([0-9a-f-]{36})/i.exec(message)?.[1];
  const correlation = /Correlation ID:\s*([0-9a-f-]{36})/i.exec(message)?.[1];
  return `${trace ? ` · Trace ID ${trace}` : ''}${correlation ? ` · Correlation ID ${correlation}` : ''}`;
}

/**
 * Traduz os erros do Microsoft Entra ID (códigos AADSTS) para mensagens com a providência a tomar.
 * delegated: erro na renovação do token da conta conectada (um "invalid_grant" sem código conhecido
 * também pede para conectar a conta de novo). Devolve null quando não há tradução.
 */
export function aadError(err, { delegated = false } = {}) {
  if (!(err instanceof ApiError)) return null;
  const aad = /AADSTS(\d+)/.exec(err.message);
  const ids = traceIds(err.message);
  if (aad?.[1] === '700027' && /expired/i.test(err.message)) return new ApiError(`${CERTIFICATE_EXPIRED} (AADSTS700027${ids})`, err);
  if (aad && AAD_ERRORS[aad[1]]) return new ApiError(`${AAD_ERRORS[aad[1]]} (AADSTS${aad[1]}${ids})`, err);
  if (delegated && err.code === 'invalid_grant') {
    const detail = String(err.message || '').split(/\r?\n/)[0].slice(0, 300);
    return new ApiError(`A autorização da conta conectada não vale mais (${detail}). ${RECONNECT}`, err);
  }
  return null;
}

/**
 * Asserção do cliente para a autenticação com certificado (JWT assinado com PS256): substitui o
 * segredo do cliente. thumbprint256: impressão digital SHA-256 do certificado (base64url).
 */
export function clientAssertion({ clientId, audience, privateKey, thumbprint256, now = Date.now() }) {
  const part = (value) => Buffer.from(JSON.stringify(value)).toString('base64url');
  const iat = Math.floor(now / 1000);
  const header = { alg: 'PS256', typ: 'JWT', 'x5t#S256': thumbprint256 };
  const claims = { aud: audience, iss: clientId, sub: clientId, jti: crypto.randomUUID(), iat, nbf: iat, exp: iat + 600 };
  const input = `${part(header)}.${part(claims)}`;
  const signature = crypto.sign('sha256', Buffer.from(input), {
    key: privateKey,
    padding: crypto.constants.RSA_PKCS1_PSS_PADDING,
    saltLength: crypto.constants.RSA_PSS_SALTLEN_DIGEST,
  });
  return `${input}.${signature.toString('base64url')}`;
}

const tokenUrl = (login, tenantId) => `${login}/${encodeURIComponent(tenantId)}/oauth2/v2.0/token`;
const scopeText = (scopes) => [...new Set(scopes)].join(' ');

/** Conta que entrou, pelos dados do id_token: { id, tenantId, username, name }. */
export function accountFromTokens(tokens) {
  let claims = {};
  try {
    claims = JSON.parse(Buffer.from(String(tokens?.id_token || '').split('.')[1] || '', 'base64url').toString('utf8')) || {};
  } catch {
    // sem id_token legível: a conta fica sem nome
  }
  return {
    id: String(claims.oid || claims.sub || ''),
    tenantId: String(claims.tid || ''),
    username: String(claims.preferred_username || claims.email || claims.upn || ''),
    name: String(claims.name || ''),
  };
}

/** Permissões concedidas, informadas pela Microsoft junto com o token. */
export const grantedScopes = (tokens) => String(tokens?.scope || '').split(/\s+/).filter(Boolean);

/**
 * Início da entrada de uma conta pelo código de dispositivo. Devolve { deviceCode, userCode,
 * verificationUri, expiresIn (s), interval (s) }: a pessoa abre o endereço, digita o código e entra
 * com a conta; enquanto isso, o CLEAN consulta o resultado (pollDeviceCode).
 */
export async function startDeviceCode({ login = MICROSOFT_LOGIN, tenantId, clientId, scopes, signal }) {
  let res;
  try {
    res = await request(`${login}/${encodeURIComponent(tenantId)}/oauth2/v2.0/devicecode`, {
      method: 'POST',
      form: { client_id: clientId, scope: scopeText(scopes) },
      signal,
      retries: 1,
      timeoutMs: 20000,
    });
  } catch (err) {
    throw aadError(err) || err;
  }
  if (!res?.device_code || !res.user_code) throw new ApiError('O Microsoft Entra ID não devolveu o código para a entrada da conta.');
  return {
    deviceCode: String(res.device_code),
    userCode: String(res.user_code),
    verificationUri: String(res.verification_uri || 'https://microsoft.com/devicelogin'),
    expiresIn: Math.min(Math.max(Number(res.expires_in) || 900, 60), 3600),
    interval: Math.min(Math.max(Number(res.interval) || 5, 1), 60),
  };
}

/**
 * Uma consulta da entrada pelo código: { status: 'pending' } enquanto a pessoa não termina (ou
 * 'slow_down': consultar com menos frequência), { status: 'connected', tokens } quando termina, ou
 * um erro (entrada recusada, código expirado, aplicativo sem fluxos de clientes públicos...). Um erro
 * com retryable (rede, tempo esgotado, falha temporária da Microsoft) não encerra a entrada: o
 * chamador consulta de novo.
 */
export async function pollDeviceCode({ login = MICROSOFT_LOGIN, tenantId, clientId, deviceCode, signal }) {
  try {
    const tokens = await request(tokenUrl(login, tenantId), {
      method: 'POST',
      form: { grant_type: 'urn:ietf:params:oauth:grant-type:device_code', client_id: clientId, device_code: deviceCode },
      signal,
      retries: 1,
      timeoutMs: 20000,
    });
    if (!tokens?.access_token) throw new ApiError('O Microsoft Entra ID não devolveu um token de acesso.');
    if (!tokens.refresh_token) throw new ApiError('O Microsoft Entra ID não devolveu o token de atualização (permissão offline_access): confira as permissões do aplicativo e conecte a conta de novo.');
    return { status: 'connected', tokens };
  } catch (err) {
    if (err instanceof ApiError) {
      if (err.code === 'authorization_pending') return { status: 'pending' };
      if (err.code === 'slow_down') return { status: 'slow_down' };
      if (err.code === 'authorization_declined') throw new ApiError('A entrada foi cancelada ou recusada na página da Microsoft.', { code: err.code });
      if (err.code === 'expired_token' || err.code === 'code_expired') throw new ApiError('O código expirou antes de a entrada terminar: clique em "Conectar conta" de novo.', { code: 'expired_token' });
      if (err.code === 'bad_verification_code') throw new ApiError('Código de entrada não reconhecido pela Microsoft: clique em "Conectar conta" de novo.', { code: err.code });
    }
    throw aadError(err) || err;
  }
}

export class MicrosoftAuth {
  /**
   * creds: graph do cadastro — { tenantId, clientId, auth, certificate: { thumbprint256 },
   *   account: { scopes } }. secrets: { clientSecret, certificateKey, refreshToken } (já decifrados).
   * options: login (endereço do Entra ID; trocado nos testes), signal, onRefreshToken(token):
   *   chamado quando a Microsoft devolve um novo token de atualização (conta conectada), para gravá-lo.
   */
  constructor(creds = {}, secrets = {}, { login = MICROSOFT_LOGIN, signal, onRefreshToken } = {}) {
    this.creds = creds || {};
    this.secrets = secrets || {};
    this.login = login || MICROSOFT_LOGIN;
    this.signal = signal;
    this.onRefreshToken = onRefreshToken;
    this.mode = MS_AUTH[this.creds.auth] ? this.creds.auth : 'secret';
    this.refreshToken = this.secrets.refreshToken || '';
    this.cache = new Map(); // recurso → { token, expires, promise }
    // Recusa definitiva (credencial inválida, autorização revogada...): as próximas chamadas recebem o
    // mesmo erro, sem repetir o pedido à Microsoft a cada caixa.
    this.failure = null;
  }

  get delegated() {
    return this.mode === 'delegated';
  }

  /** Quando vence o token de acesso ao recurso que está guardado (0 se não há). */
  expiresAt(resource = 'graph') {
    return this.cache.get(resource)?.expires || 0;
  }

  /**
   * Token de acesso ao recurso ('graph' ou 'imap'). É renovado quando faltam menos de minValidityMs
   * para vencer (padrão: 2 minutos; uma sessão IMAP pede mais, porque o Exchange Online a encerra
   * quando o token vence). force: pede um novo.
   */
  async token(resource = 'graph', { force = false, minValidityMs = 120000 } = {}) {
    if (this.failure) throw this.failure;
    let entry = this.cache.get(resource);
    if (!entry) this.cache.set(resource, (entry = { token: null, expires: 0, promise: null }));
    if (!force && entry.token && Date.now() < entry.expires - minValidityMs) return entry.token;
    entry.promise ||= this.#request(resource)
      .then((res) => {
        entry.token = res.access_token;
        entry.expires = Date.now() + (Number(res.expires_in) || 3600) * 1000;
        return entry.token;
      })
      .finally(() => {
        entry.promise = null;
      });
    return entry.promise;
  }

  async #request(resource) {
    const { tenantId, clientId } = this.creds;
    if (!tenantId || !clientId) throw new ApiError('Informe o ID do locatário e o ID do aplicativo (cliente).');
    const url = tokenUrl(this.login, tenantId);
    const form = { client_id: clientId };
    if (this.delegated) {
      if (!this.refreshToken) throw new ApiError(`Nenhuma conta Microsoft conectada a esta conexão. ${RECONNECT}`);
      const granted = (this.creds.account?.scopes || []).filter((s) => s !== 'offline_access');
      const scopes = granted.length ? granted : delegatedScopes(resource === 'imap' ? 'imap' : 'graph', { tenantId });
      Object.assign(form, { grant_type: 'refresh_token', refresh_token: this.refreshToken, scope: scopeText([...scopes, 'offline_access']) });
    } else {
      Object.assign(form, { grant_type: 'client_credentials', scope: APP_SCOPES[resource] || APP_SCOPES.graph });
      if (this.mode === 'certificate') {
        const key = this.secrets.certificateKey;
        const thumbprint256 = this.creds.certificate?.thumbprint256;
        if (!key || !thumbprint256) throw new ApiError('A conexão não tem o certificado do aplicativo: gere ou importe o certificado.');
        const notAfter = Date.parse(this.creds.certificate.notAfter);
        if (notAfter <= Date.now()) {
          throw (this.failure = new ApiError(`${CERTIFICATE_EXPIRED.replace('venceu:', `venceu em ${new Date(notAfter).toLocaleDateString('pt-BR')}:`)}`));
        }
        form.client_assertion_type = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer';
        form.client_assertion = clientAssertion({ clientId, audience: url, privateKey: key, thumbprint256 });
      } else {
        if (!this.secrets.clientSecret) throw new ApiError('Informe o segredo do cliente (valor do segredo criado no registro do aplicativo).');
        form.client_secret = this.secrets.clientSecret;
      }
    }
    let res;
    try {
      res = await request(url, { method: 'POST', form, signal: this.signal, retries: 3 });
    } catch (err) {
      const translated = aadError(err, { delegated: this.delegated }) || err;
      if (err instanceof ApiError && !err.retryable && err.status >= 400 && err.status < 500) this.failure = translated;
      throw translated;
    }
    if (!res?.access_token) throw new ApiError('O Microsoft Entra ID não devolveu um token de acesso.');
    if (this.delegated && res.refresh_token && res.refresh_token !== this.refreshToken) {
      this.refreshToken = res.refresh_token;
      try {
        this.onRefreshToken?.(res.refresh_token);
      } catch {
        // a gravação é tentada de novo na próxima renovação
      }
    }
    return res;
  }
}
