// Validação dos dados recebidos pela API.
import path from 'node:path';
import crypto from 'node:crypto';
import { validateTerm, foldText } from '../scan/matcher.js';
import { VALIDATORS } from '../scan/presets.js';

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
  const names = using.map((s) => `"${s.name}"`).join(', ');
  const one = using.length === 1;
  const it = what.startsWith('A ') ? 'a' : 'o';
  throw new HttpError(409, `${what} está em uso ${one ? 'no agendamento' : 'nos agendamentos'} ${names}. Retire-${it} ${one ? 'do agendamento' : 'dos agendamentos'} (ou exclua ${one ? 'o agendamento' : 'os agendamentos'}) antes de excluir.`, 'in-use');
}

export const EMAIL_RE = /^[^\s@<>()",;:]+@[^\s@<>()",;:]+$/;
export const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
export const DOMAIN_RE = /^(?=.{3,253}$)[a-z0-9-]+(\.[a-z0-9-]+)+$/i;

export function text(value, field, { required = false, max = 500 } = {}) {
  const v = typeof value === 'string' ? value.trim() : value === undefined || value === null ? '' : String(value).trim();
  if (required && !v) throw bad(`Informe ${field}.`);
  if (v.length > max) throw bad(`${field} deve ter no máximo ${max} caracteres.`);
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

function parseCloud(body, type, { existing, box, mailSource }) {
  let graph;
  let clientSecret;
  let credentialsFrom = null;
  if (body.credentialsFrom) {
    // Mesmo registro de aplicativo de uma conexão de e-mail do Microsoft 365: as credenciais ficam
    // ligadas a ela (um novo segredo salvo na conexão também passa a valer para o repositório).
    if (!mailSource || mailSource.type !== 'graph' || !mailSource.graph) throw bad('Escolha uma conexão de e-mail do Microsoft 365 para usar as mesmas credenciais.');
    if (!mailSource.secrets?.clientSecret) throw bad(`A conexão "${mailSource.name}" não tem o segredo do cliente salvo.`);
    graph = { tenantId: mailSource.graph.tenantId, clientId: mailSource.graph.clientId };
    clientSecret = mailSource.secrets.clientSecret;
    credentialsFrom = mailSource.id;
  } else {
    const cloudBefore = CLOUD_REPO_TYPES.has(existing?.type);
    const previousSecret = cloudBefore ? existing.secrets?.clientSecret || null : null;
    ({ graph, clientSecret } = graphCredentials(body.graph || {}, { previousSecret, previousGraph: cloudBefore ? existing.graph : null, box }));
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
    secrets: { clientSecret },
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
