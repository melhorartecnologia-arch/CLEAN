// Validação dos dados recebidos pela API.
import path from 'node:path';
import crypto from 'node:crypto';
import { validateTerm, foldText } from '../scan/matcher.js';
import { VALIDATORS } from '../scan/presets.js';
import { MS_AUTH, TENANT_KEYWORDS } from '../cloud/microsoft-auth.js';

export class HttpError extends Error {
  /** code: identificador opcional devolvido à interface junto com a mensagem (ex.: 'method-changed'). */
  constructor(status, message, code = '') {
    super(message);
    this.status = status;
    this.code = code;
  }
}

export const bad = (message) => new HttpError(400, message);

/**
 * Impede excluir um repositório ('files'), uma conexão de e-mail ('mail') ou uma lista ('list')
 * usado por agendamentos (eles passariam a falhar). what: ex.: 'O repositório'.
 */
export function assertUnused(store, kind, id, what) {
  const using = store.schedulesUsing(kind, id);
  if (!using.length) return;
  const names = (list) => list.map((s) => `"${s.name}"`).join(', ');
  const schedules = using.filter((s) => s.purpose !== 'retention');
  const policies = using.filter((s) => s.purpose === 'retention');
  const pick = (list, one, many) => (list.length ? (list.length === 1 ? one : many) : '');
  const where = [
    schedules.length && `${pick(schedules, 'no agendamento', 'nos agendamentos')} ${names(schedules)}`,
    policies.length && `${pick(policies, 'na política de retenção', 'nas políticas de retenção')} ${names(policies)}`,
  ].filter(Boolean);
  const from = [pick(schedules, 'do agendamento', 'dos agendamentos'), pick(policies, 'da política', 'das políticas')].filter(Boolean);
  const remove = [pick(schedules, 'o agendamento', 'os agendamentos'), pick(policies, 'a política', 'as políticas')].filter(Boolean);
  const it = what.startsWith('A ') ? 'a' : 'o';
  throw new HttpError(409, `${what} está em uso ${where.join(' e ')}. Retire-${it} ${from.join(' e ')} (ou exclua ${remove.join(' e ')}) antes de excluir.`, 'in-use');
}

export const EMAIL_RE = /^[^\s@<>()",;:]+@[^\s@<>()",;:]+$/;
export const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
export const DOMAIN_RE = /^(?=.{3,253}$)[a-z0-9-]+(\.[a-z0-9-]+)+$/i;

export function text(value, field, { required = false, max = 500 } = {}) {
  const v = typeof value === 'string' ? value.trim() : value === undefined || value === null ? '' : String(value).trim();
  if (required && !v) throw bad(`Informe ${field}.`);
  if (v.length > max) throw bad(`${field[0].toUpperCase()}${field.slice(1)} deve ter no máximo ${max} caracteres.`);
  return v;
}

export function lines(value, max = 500) {
  const items = Array.isArray(value) ? value : String(value || '').split(/\r?\n/);
  return [...new Set(items.map((v) => String(v).trim()).filter(Boolean))].slice(0, max);
}

/** Um endereço de e-mail (vazio é aceito; o chamador decide se é obrigatório). */
export function email(value, field) {
  const v = text(value, field, { max: 320 });
  if (v && !EMAIL_RE.test(v)) throw bad(`${field[0].toUpperCase()}${field.slice(1)} inválido: "${v.slice(0, 80)}".`);
  return v;
}

/** Lista de e-mails (texto com um por linha, vírgula ou ponto e vírgula, ou lista), sem repetições. */
export function emailList(value, field, { max = 5000, noun = 'endereços' } = {}) {
  const items = Array.isArray(value) ? value.map((v) => (typeof v === 'object' && v ? v.address : v)) : String(value || '').split(/[\r\n,;]+/);
  const seen = new Set();
  const out = [];
  for (const raw of items) {
    const address = email(raw, field);
    const key = address.toLowerCase();
    if (!address || seen.has(key)) continue;
    seen.add(key);
    out.push(address);
  }
  if (out.length > max) throw bad(`Informe no máximo ${max} ${noun}.`);
  return out;
}

/**
 * Credenciais do Microsoft Graph (registro de aplicativo no Microsoft Entra ID). Campo de segredo
 * vazio mantém o segredo salvo (previousSecret), desde que o locatário e o aplicativo continuem os
 * mesmos (previousGraph). Retorna { graph: { tenantId, clientId }, clientSecret } com o segredo cifrado.
 */
export function graphCredentials(g = {}, { previousSecret = null, previousGraph = null, box }) {
  const tenantId = text(g.tenantId, 'o ID do locatário', { required: true, max: 255 });
  if (!GUID_RE.test(tenantId) && !DOMAIN_RE.test(tenantId)) throw bad('ID do locatário inválido: use o GUID (ID do diretório) ou o domínio, ex.: empresa.onmicrosoft.com.');
  const clientId = text(g.clientId, 'o ID do cliente (aplicativo)', { required: true, max: 64 });
  if (!GUID_RE.test(clientId)) throw bad('ID do cliente inválido: use o "ID do aplicativo (cliente)" do registro do aplicativo.');
  const secret = text(g.clientSecret, 'o segredo do cliente', { max: 2000 });
  const sameApp = previousGraph && previousGraph.tenantId?.toLowerCase() === tenantId.toLowerCase() && previousGraph.clientId?.toLowerCase() === clientId.toLowerCase();
  if (!secret && previousSecret && previousGraph && !sameApp) throw bad('Ao trocar o locatário ou o aplicativo, informe o segredo do cliente novamente.');
  const clientSecret = secret ? box.seal(secret) : previousSecret;
  if (!clientSecret) throw bad('Informe o segredo do cliente (valor do segredo criado no registro do aplicativo).');
  return { graph: { tenantId, clientId }, clientSecret };
}

/**
 * Locatário e aplicativo (cliente) do Microsoft Entra ID. delegated: conta conectada, que aceita
 * também os locatários genéricos ("organizations", "common" e "consumers"). Retorna { tenantId, clientId }.
 */
export function microsoftApp(g = {}, { delegated = false } = {}) {
  let tenantId = text(g.tenantId, 'o ID do locatário', { required: true, max: 255 });
  const keyword = TENANT_KEYWORDS.has(tenantId.toLowerCase());
  if (keyword && !delegated) throw bad(`"${tenantId}" vale só para a conta conectada: para um aplicativo, informe o ID do locatário (GUID) ou o domínio, ex.: empresa.onmicrosoft.com.`);
  if (!keyword && !GUID_RE.test(tenantId) && !DOMAIN_RE.test(tenantId)) {
    throw bad(
      delegated
        ? 'Locatário inválido: use o GUID (ID do diretório), o domínio (ex.: empresa.onmicrosoft.com), "organizations" (contas de trabalho ou escola), "consumers" (contas pessoais: Outlook.com, Hotmail) ou "common" (as duas).'
        : 'ID do locatário inválido: use o GUID (ID do diretório) ou o domínio, ex.: empresa.onmicrosoft.com.',
    );
  }
  if (keyword) tenantId = tenantId.toLowerCase();
  const clientId = text(g.clientId, 'o ID do cliente (aplicativo)', { required: true, max: 64 });
  if (!GUID_RE.test(clientId)) throw bad('ID do cliente inválido: use o "ID do aplicativo (cliente)" do registro do aplicativo.');
  return { tenantId, clientId };
}

/**
 * Credenciais da Microsoft (Microsoft Entra ID) de uma conexão de e-mail — do Microsoft 365 ou do
 * IMAP da Microsoft com OAuth — numa das três formas (auth): 'secret' (aplicativo com segredo do
 * cliente), 'certificate' (aplicativo com certificado) ou 'delegated' (conta Microsoft conectada).
 * Campos vazios mantêm o que está salvo (previous: { graph, secrets } do cadastro atual): o segredo e
 * a conta conectada só para o mesmo locatário e aplicativo; o certificado também em outro aplicativo
 * (a chave privada não sai do CLEAN — o arquivo do certificado é que precisa ser enviado a ele).
 * pending: credenciais ainda não salvas — signIn(id) (entrada da conta) e certificate(id) (certificado
 * gerado ou importado). purpose: 'graph' ou 'imap' (a conta conectada autoriza um dos dois). sourceId:
 * a conexão em edição (null numa nova): a entrada da conta só vale para a conexão em que foi começada.
 * Retorna { graph, secrets } com os segredos cifrados.
 */
export function microsoftCredentials(g = {}, { previous = null, box, pending = {}, purpose = 'graph', sourceId = null }) {
  const auth = Object.hasOwn(MS_AUTH, g.auth) ? g.auth : 'secret';
  const { tenantId, clientId } = microsoftApp(g, { delegated: auth === 'delegated' });
  const before = previous?.graph || null;
  const saved = previous?.secrets || {};
  const beforeAuth = before ? before.auth || 'secret' : null;
  const sameApp = Boolean(before) && String(before.tenantId).toLowerCase() === tenantId.toLowerCase() && String(before.clientId).toLowerCase() === clientId.toLowerCase();
  const graph = { tenantId, clientId, auth };
  const secrets = {};
  if (auth === 'secret') {
    const secret = text(g.clientSecret, 'o segredo do cliente', { max: 2000 });
    if (!secret && saved.clientSecret && beforeAuth === 'secret' && !sameApp) throw bad('Ao trocar o locatário ou o aplicativo, informe o segredo do cliente novamente.');
    secrets.clientSecret = secret ? box.seal(secret) : sameApp && beforeAuth === 'secret' ? saved.clientSecret : null;
    if (!secrets.clientSecret) throw bad('Informe o segredo do cliente (valor do segredo criado no registro do aplicativo).');
  } else if (auth === 'certificate') {
    const id = typeof g.certificateId === 'string' ? g.certificateId : '';
    if (id) {
      const item = pending.certificate?.(id);
      if (!item) throw bad('O certificado gerado (ou importado) não está mais disponível para salvar (mais de 24 horas ou o CLEAN foi reiniciado): gere ou importe o certificado de novo.');
      graph.certificate = item.certificate;
      secrets.certificateKey = box.seal(item.privateKeyPem);
    } else if (beforeAuth === 'certificate' && before.certificate && saved.certificateKey) {
      graph.certificate = before.certificate;
      secrets.certificateKey = saved.certificateKey;
    } else {
      throw bad('Gere o certificado (ou importe um existente) e envie o arquivo do certificado ao registro do aplicativo.');
    }
  } else {
    const id = typeof g.signIn === 'string' ? g.signIn : '';
    if (id) {
      const flow = pending.signIn?.(id);
      if (!flow) throw bad('A entrada da conta não está mais disponível (cancelada, mais de 2 horas ou o CLEAN foi reiniciado): clique em "Conectar conta" de novo.');
      if (flow.status === 'pending') throw bad('A entrada da conta ainda não terminou: conclua a entrada na página da Microsoft com o código mostrado.');
      if (flow.status !== 'connected') throw bad(`A entrada da conta não foi concluída${flow.error ? ` (${flow.error})` : ''}: clique em "Conectar conta" de novo.`);
      if ((flow.sourceId || null) !== (sourceId || null)) throw bad('A entrada da conta foi feita em outra conexão: clique em "Conectar conta" de novo.');
      if (String(flow.tenantId).toLowerCase() !== tenantId.toLowerCase() || String(flow.clientId).toLowerCase() !== clientId.toLowerCase()) {
        throw bad('O locatário ou o aplicativo foi alterado depois da entrada da conta: clique em "Conectar conta" de novo.');
      }
      if (flow.purpose !== purpose) throw bad('A conta foi conectada para outro tipo de conexão: clique em "Conectar conta" de novo.');
      graph.account = flow.account;
      secrets.refreshToken = box.seal(flow.refreshToken);
    } else if (beforeAuth === 'delegated' && sameApp && before.account && saved.refreshToken && (before.account.purpose || 'graph') === purpose) {
      graph.account = before.account;
      secrets.refreshToken = saved.refreshToken;
    } else {
      throw bad(beforeAuth === 'delegated' && !sameApp ? 'Ao trocar o locatário ou o aplicativo, conecte a conta novamente ("Conectar conta").' : 'Conecte a conta Microsoft: clique em "Conectar conta" e entre com a conta cujas caixas serão analisadas.');
    }
  }
  return { graph, secrets };
}

/**
 * Endereço de um site do SharePoint (https), sem parâmetros e sem barra final. Links de páginas e
 * bibliotecas do site também são aceitos (o site é localizado na análise).
 */
export function siteUrl(value) {
  const v = text(value, 'o endereço do site', { max: 2000 });
  if (!v) return '';
  let u;
  try {
    u = new URL(v);
  } catch {
    throw bad(`Endereço de site inválido: "${v.slice(0, 120)}". Use o endereço completo, ex.: https://empresa.sharepoint.com/sites/Financeiro.`);
  }
  if (u.protocol !== 'https:' || u.username || u.password) throw bad(`Endereço de site inválido: "${v.slice(0, 120)}". Use um endereço https://.`);
  if (/-my\.sharepoint\.[a-z.]+$/i.test(u.hostname) || /^\/personal\//i.test(u.pathname)) {
    throw bad(`"${v.slice(0, 120)}" é um OneDrive pessoal: cadastre um repositório do tipo OneDrive com a conta do usuário.`);
  }
  return `https://${u.hostname.toLowerCase()}${u.pathname.replace(/\/+$/, '')}`;
}

/** Caminho absoluto de pasta: C:\..., \\servidor\compartilhamento\... (ou /... fora do Windows). */
export function normalizeRepoPath(value) {
  let p = text(value, 'o caminho da pasta', { required: true, max: 1000 });
  const isWindowsStyle = /^[a-z]:[\\/]/i.test(p) || /^\\\\[^\\]+\\[^\\]+/.test(p);
  if (process.platform === 'win32' || isWindowsStyle) {
    if (!isWindowsStyle) throw bad('Informe um caminho completo, como D:\\Dados ou \\\\servidor\\compartilhamento.');
    p = p.replace(/\//g, '\\');
    if (!/^[a-z]:\\$/i.test(p)) p = p.replace(/\\+$/, '');
    return p;
  }
  if (!path.isAbsolute(p)) throw bad('Informe um caminho absoluto para a pasta.');
  return p.length > 1 ? p.replace(/\/+$/, '') : p;
}

const CLOUD_REPO_TYPES = new Set(['onedrive', 'sharepoint']);
const NO_AUDIT = { enabled: false, computer: '', localPath: '', days: 30, maxEvents: 200000, ignoreUsers: [] };

/** Descrição de onde estão os arquivos de um repositório na nuvem (mostrada como "caminho"). */
export function describeCloud(type, cloud) {
  if (type === 'onedrive') {
    if (cloud.scope === 'all') return 'OneDrive: todas as contas';
    return `OneDrive: ${cloud.accounts.slice(0, 3).join(', ')}${cloud.accounts.length > 3 ? ` e mais ${cloud.accounts.length - 3}` : ''}`;
  }
  if (cloud.scope === 'all') return 'SharePoint: todos os sites';
  return `SharePoint: ${cloud.sites.slice(0, 2).join(', ')}${cloud.sites.length > 2 ? ` e mais ${cloud.sites.length - 2}` : ''}`;
}

/**
 * Valida um repositório: pasta do Windows (type 'local', padrão) ou OneDrive/SharePoint.
 * options: existing (cadastro atual, para manter o segredo salvo), box (cifra os segredos),
 * mailSource (conexão de e-mail do Microsoft 365 da qual copiar as credenciais), forTest (teste da
 * conexão: o nome não é obrigatório).
 */
export function parseRepository(body = {}, { existing = null, box = null, mailSource = null, forTest = false } = {}) {
  const type = CLOUD_REPO_TYPES.has(body.type) ? body.type : 'local';
  const common = {
    type,
    name: text(body.name, 'o nome do repositório', { required: !forTest, max: 200 }),
    description: text(body.description, 'a descrição', { max: 1000 }),
    exclude: lines(body.exclude),
    // Permite excluir os arquivos encontrados (automaticamente na análise ou pelo relatório).
    allowDelete: body.allowDelete === true,
  };
  if (type !== 'local') return { ...common, ...parseCloud(body, type, { existing, box, mailSource }) };
  const audit = body.audit || {};
  const computer = text(audit.computer, 'o computador da auditoria', { max: 255 });
  if (computer && !/^[A-Za-z0-9._-]+$/.test(computer)) throw bad('Nome de computador inválido para a auditoria.');
  const localPath = text(audit.localPath, 'o caminho local da auditoria', { max: 1000 });
  if (localPath && !/^[a-z]:\\/i.test(localPath.replace(/\//g, '\\'))) throw bad('O caminho local no servidor deve começar com a letra da unidade (ex.: E:\\Compartilhamentos).');
  const days = Number(audit.days) || 30;
  const maxEvents = Number(audit.maxEvents) || 200000;
  return {
    ...common,
    path: normalizeRepoPath(body.path),
    // Campos dos repositórios na nuvem ficam nulos (ao trocar o tipo, os dados antigos são descartados).
    deleteMode: null,
    graph: null,
    secrets: null,
    credentialsFrom: null,
    cloud: null,
    audit: {
      enabled: Boolean(audit.enabled),
      computer,
      localPath: localPath.replace(/\//g, '\\'),
      days: Math.min(Math.max(Math.round(days), 1), 365),
      maxEvents: Math.min(Math.max(Math.round(maxEvents), 100), 5_000_000),
      ignoreUsers: lines(audit.ignoreUsers, 100).map((u) => u.replace(/;/g, '')),
    },
  };
}

/**
 * Credenciais de aplicativo de uma conexão de e-mail do Microsoft 365 (segredo do cliente ou
 * certificado) para um repositório do OneDrive/SharePoint ligado a ela: { graph, secrets }, ou null se
 * a conexão não tiver credenciais de aplicativo (conta conectada: as permissões são só de e-mail).
 */
export function linkedCredentials(source) {
  if (source?.type !== 'graph' || !source.graph) return null;
  const auth = source.graph.auth || 'secret';
  const { tenantId, clientId } = source.graph;
  if (auth === 'secret' && source.secrets?.clientSecret) return { graph: { tenantId, clientId }, secrets: { clientSecret: source.secrets.clientSecret } };
  if (auth === 'certificate' && source.secrets?.certificateKey && source.graph.certificate) {
    return { graph: { tenantId, clientId, auth, certificate: source.graph.certificate }, secrets: { certificateKey: source.secrets.certificateKey } };
  }
  return null;
}

function parseCloud(body, type, { existing, box, mailSource }) {
  let graph;
  let secrets;
  let credentialsFrom = null;
  if (body.credentialsFrom) {
    // Mesmo registro de aplicativo de uma conexão de e-mail do Microsoft 365: as credenciais ficam
    // ligadas a ela (um novo segredo ou certificado salvo na conexão também passa a valer para o repositório).
    if (!mailSource || mailSource.type !== 'graph' || !mailSource.graph) throw bad('Escolha uma conexão de e-mail do Microsoft 365 para usar as mesmas credenciais.');
    if (mailSource.graph.auth === 'delegated') throw bad(`A conexão "${mailSource.name}" usa uma conta conectada (permissões só de e-mail): o OneDrive e o SharePoint precisam das credenciais de um aplicativo (segredo do cliente ou certificado).`);
    const linked = linkedCredentials(mailSource);
    if (!linked) throw bad(`A conexão "${mailSource.name}" não tem o segredo do cliente (ou o certificado) salvo.`);
    ({ graph, secrets } = linked);
    credentialsFrom = mailSource.id;
  } else {
    const cloudBefore = CLOUD_REPO_TYPES.has(existing?.type);
    const g = body.graph || {};
    const before = cloudBefore ? existing.graph || {} : {};
    const same = (a, b) => String(a || '').trim().toLowerCase() === String(b || '').trim().toLowerCase();
    // Certificado copiado de uma conexão de e-mail que deixou de estar ligada: continua valendo para o
    // mesmo aplicativo enquanto nenhum segredo do cliente for informado.
    const keepCertificate =
      before.auth === 'certificate' && before.certificate && existing.secrets?.certificateKey && !text(g.clientSecret, 'o segredo do cliente', { max: 2000 }) && same(before.tenantId, g.tenantId) && same(before.clientId, g.clientId);
    if (keepCertificate) {
      const app = microsoftApp(g);
      graph = { ...app, auth: 'certificate', certificate: before.certificate };
      secrets = { certificateKey: existing.secrets.certificateKey };
    } else {
      const previousSecret = cloudBefore ? existing.secrets?.clientSecret || null : null;
      let clientSecret;
      ({ graph, clientSecret } = graphCredentials(g, { previousSecret, previousGraph: cloudBefore ? existing.graph : null, box }));
      secrets = { clientSecret };
    }
  }
  const scope = body.scope === 'all' ? 'all' : 'list';
  const cloud = { scope, accounts: [], sites: [], exclude: lines(body.excludeTargets, 500) };
  if (scope === 'list' && type === 'onedrive') {
    cloud.accounts = emailList(body.accounts, 'o e-mail da conta', { noun: 'contas' });
    if (cloud.accounts.length === 0) throw bad('Informe ao menos uma conta de OneDrive (e-mail do usuário) ou escolha "Todas as contas".');
  }
  if (scope === 'list' && type === 'sharepoint') {
    const items = Array.isArray(body.sites) ? body.sites : String(body.sites || '').split(/[\r\n]+/);
    cloud.sites = [...new Set(items.map(siteUrl).filter(Boolean))];
    if (cloud.sites.length === 0) throw bad('Informe ao menos o endereço de um site do SharePoint ou escolha "Todos os sites".');
    if (cloud.sites.length > 2000) throw bad('Informe no máximo 2000 sites.');
  }
  return {
    path: describeCloud(type, cloud),
    // Lixeira do site/OneDrive (recuperável) ou exclusão definitiva.
    deleteMode: body.deleteMode === 'permanent' ? 'permanent' : 'trash',
    graph,
    secrets,
    credentialsFrom,
    cloud,
    audit: { ...NO_AUDIT },
  };
}

const MAX_TERMS = 50000;

/** Valida os termos de uma lista, preservando os identificadores existentes e removendo duplicados. */
export function parseTerms(input) {
  if (!Array.isArray(input)) throw bad('A lista de termos é inválida.');
  if (input.length > MAX_TERMS) throw bad(`Cada lista pode ter no máximo ${MAX_TERMS} termos.`);
  const seen = new Set();
  const terms = [];
  input.forEach((raw, index) => {
    const type = raw?.type === 'regex' ? 'regex' : 'text';
    const term = {
      id: typeof raw?.id === 'string' && /^[\w-]{1,64}$/.test(raw.id) ? raw.id : crypto.randomUUID(),
      type,
      value: text(raw?.value, `o termo ${index + 1}`, { max: 2000 }),
      wholeWord: type === 'text' ? Boolean(raw?.wholeWord) : false,
      validator: type === 'regex' && raw?.validator && VALIDATORS[raw.validator] ? raw.validator : null,
      label: text(raw?.label, `o rótulo do termo ${index + 1}`, { max: 200 }),
    };
    if (!term.value) return;
    const error = validateTerm(term);
    if (error) throw bad(`Termo ${index + 1} ("${term.value.slice(0, 60)}"): ${error}`);
    const key = `${type}|${type === 'text' ? foldText(term.value) : term.value}|${term.wholeWord}|${term.validator}`;
    if (seen.has(key)) return;
    seen.add(key);
    terms.push(term);
  });
  return terms;
}

export function parseList(body = {}) {
  return {
    name: text(body.name, 'o nome da lista', { required: true, max: 200 }),
    description: text(body.description, 'a descrição', { max: 1000 }),
    terms: parseTerms(body.terms || []),
  };
}
