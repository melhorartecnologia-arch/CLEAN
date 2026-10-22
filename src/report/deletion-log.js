// Registro geral de exclusões (data/exclusoes.ndjson): todas as exclusões feitas pelo CLEAN —
// automáticas (políticas de retenção, buscas por tipo e análises com exclusão automática), manuais
// pelo relatório e em lote —, que continua valendo depois que o relatório da análise é excluído.
//
// Aqui ficam os filtros, a forma de exibição de cada evento e o resumo (totais, por dia, política,
// repositório, situação e tipo de local), calculados numa passada só pelo arquivo, lido em fluxo
// (ele pode ter centenas de MB). As linhas antigas, gravadas antes de o registro levar o contexto
// da análise e o local do item, são completadas com o cadastro atual (análises, agendamentos,
// repositórios e conexões); se a análise já foi excluída, com o que a própria linha permite deduzir.
import crypto from 'node:crypto';
import fs from 'node:fs/promises';
import { foldText } from '../scan/matcher.js';
import { DELETION_LABELS, MAIL_DELETION_LABELS, deletionHow } from './model.js';
import { FILE_CRITERIA, MAIL_CRITERIA, amountText } from '../retention/policy.js';

export const STATUSES = ['deleted', 'missing', 'changed', 'failed'];
export const MODES = ['auto', 'manual', 'retention'];
export const PURPOSES = ['retention', 'types', 'date', 'terms'];
/** Tipos de local do item ('cloud': OneDrive ou SharePoint, quando a linha não permite distinguir). */
export const KINDS = ['local', 'onedrive', 'sharepoint', 'cloud', 'mail', 'teams'];
const TARGET_TYPES = new Set(['local', 'onedrive', 'sharepoint', 'mail', 'teams']);
const SCAN_KINDS = new Set(['files', 'mail', 'teams']);

export const KIND_LABELS = { local: 'Pasta do Windows', onedrive: 'OneDrive', sharepoint: 'SharePoint', cloud: 'OneDrive/SharePoint', mail: 'E-mail', teams: 'Teams', '': '—' };
export const PURPOSE_LABELS = { retention: 'Política de retenção', types: 'Busca por tipo', date: 'Busca por data de alteração', terms: 'Análise por termos', '': '—' };
export const ORIGIN_LABELS = { retention: 'Automática (política)', auto: 'Automática', manual: 'Manual', '': '—' };
export const METHOD_LABELS = { file: 'Definitiva', permanent: 'Definitiva', trash: 'Para a lixeira', teams: 'Removida do Teams (recuperável)' };
export const NO_POLICY = 'Sem política (análise avulsa ou manual)';
export const UNKNOWN_REPOSITORY = '(não identificado)';

/** Parâmetros de filtro aceitos pela API (os demais são ignorados). */
export const DELETION_FILTER_KEYS = ['from', 'to', 'status', 'mode', 'purpose', 'kind', 'scheduleId', 'repository', 'scanId', 'q'];

const SEP = ' › ';
const CLOUD_KINDS = new Set(['onedrive', 'sharepoint', 'cloud']);
// "política de retenção "Nome"" ou "agendamento "Nome"" no campo "por" (linhas antigas sem o agendamento).
const BY_SCHEDULE_RE = /(política de retenção|agendamento) "([^"]+)"/;
const PERSONAL_URL_RE = /-my\.sharepoint\.[a-z.]+\/personal\//i;

/** Filtro inválido: a API responde 400 com a mensagem. */
export class DeletionFilterError extends Error {
  constructor(message) {
    super(message);
    this.status = 400;
  }
}

const bad = (message) => new DeletionFilterError(message);
const str = (value) => (typeof value === 'string' ? value : value === null || value === undefined ? '' : String(value));

// -- filtros -------------------------------------------------------------------------------------

const DAY_RE = /^(\d{4})-(\d{2})-(\d{2})$/;

/** "AAAA-MM-DD" -> [ano, mês, dia], recusando datas impossíveis (ex.: 2026-02-30). */
function parseDay(value, label) {
  const m = DAY_RE.exec(value);
  if (!m) throw bad(`${label} inválida: use o formato AAAA-MM-DD.`);
  const [y, mo, d] = [Number(m[1]), Number(m[2]), Number(m[3])];
  const date = new Date(y, mo - 1, d);
  if (date.getFullYear() !== y || date.getMonth() !== mo - 1 || date.getDate() !== d) throw bad(`${label} inválida: o dia ${value} não existe.`);
  if (y < 1970) throw bad(`${label} inválida: use uma data a partir de 01/01/1970.`);
  return [y, mo, d];
}

/** Início do dia (no fuso do servidor) em milissegundos; plus: dias a somar. */
function dayStart(value, plus = 0) {
  const [y, m, d] = value.split('-').map(Number);
  return new Date(y, m - 1, d + plus).getTime();
}

/** Dia local de um instante: "AAAA-MM-DD" (no fuso do servidor). */
export function localDay(ms) {
  const d = new Date(ms);
  const p = (n) => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
}

/**
 * Valida os filtros recebidos pela API e devolve só os informados, como texto canônico (usado também
 * na chave do cache). Chaves desconhecidas são ignoradas; valores inválidos lançam DeletionFilterError.
 *   from, to: "AAAA-MM-DD" no fuso do servidor (from inclusive; to inclui o dia inteiro)
 *   status: lista separada por vírgulas (deleted, missing, changed, failed)
 *   mode: auto | manual | retention · purpose: retention | types | date | terms
 *   kind: local | onedrive | sharepoint | cloud (OneDrive ou SharePoint) | mail | teams
 *   scheduleId: agendamento ou política ("none": sem agendamento) · repository: repositório ou conexão
 *   scanId: análise · q: busca livre (sem diferenciar acentos e maiúsculas)
 */
export function parseDeletionFilters(query = {}) {
  const value = (key) => {
    const raw = query[key];
    if (Array.isArray(raw)) return key === 'status' ? raw.filter((v) => typeof v === 'string').join(',') : typeof raw[0] === 'string' ? raw[0].trim() : '';
    return typeof raw === 'string' ? raw.trim() : '';
  };
  const filters = {};
  const from = value('from');
  const to = value('to');
  if (from) {
    parseDay(from, 'Data inicial');
    filters.from = from;
  }
  if (to) {
    parseDay(to, 'Data final');
    filters.to = to;
  }
  if (from && to && from > to) throw bad('A data inicial é posterior à data final.');
  const status = value('status');
  if (status) {
    const list = status.split(',').map((s) => s.trim()).filter(Boolean);
    const wrong = list.find((s) => !STATUSES.includes(s));
    if (wrong) throw bad(`Situação desconhecida: "${wrong.slice(0, 40)}". Use deleted, missing, changed ou failed.`);
    if (list.length) filters.status = STATUSES.filter((s) => list.includes(s)).join(',');
  }
  const oneOf = (key, allowed, what) => {
    const v = value(key);
    if (!v) return;
    if (!allowed.includes(v)) throw bad(`${what}: "${v.slice(0, 40)}". Use ${allowed.join(', ')}.`);
    filters[key] = v;
  };
  oneOf('mode', MODES, 'Origem desconhecida');
  oneOf('purpose', PURPOSES, 'Finalidade desconhecida');
  oneOf('kind', KINDS, 'Tipo de local desconhecido');
  for (const [key, max, what] of [
    ['scheduleId', 300, 'O agendamento'],
    ['repository', 200, 'O repositório'],
    ['scanId', 100, 'A análise'],
  ]) {
    const v = value(key);
    if (v.length > max) throw bad(`${what} informad${what.startsWith('A ') ? 'a' : 'o'} no filtro é inválid${what.startsWith('A ') ? 'a' : 'o'}.`);
    if (v) filters[key] = v;
  }
  const q = value('q');
  if (q.length > 500) throw bad('A busca deve ter no máximo 500 caracteres.');
  if (q) filters.q = q;
  return filters;
}

/** Filtros prontos para a comparação de cada evento. */
function compileFilters(filters) {
  return {
    fromMs: filters.from ? dayStart(filters.from) : null,
    toMs: filters.to ? dayStart(filters.to, 1) : null, // dia seguinte, 00:00 (exclusive)
    statuses: filters.status ? new Set(filters.status.split(',')) : null,
    mode: filters.mode || '',
    purpose: filters.purpose || '',
    kinds: filters.kind === 'cloud' ? CLOUD_KINDS : filters.kind ? new Set([filters.kind]) : null,
    policy: filters.scheduleId || '',
    repository: filters.repository || '',
    scanId: filters.scanId || '',
    q: filters.q ? foldText(filters.q).trim() : '',
  };
}

/** Texto da busca livre de um evento (em minúsculas e sem acentos). */
const haystack = (n) => foldText([n.item, n.scanName, n.scheduleName, n.by, n.error, n.note, n.repositoryName, n.relativePath, n.folder, n.mailbox, n.formerNames].filter(Boolean).join(' | '));

function matches(n, c) {
  if (c.fromMs !== null || c.toMs !== null) {
    if (!Number.isFinite(n.time)) return false;
    if (c.fromMs !== null && n.time < c.fromMs) return false;
    if (c.toMs !== null && n.time >= c.toMs) return false;
  }
  if (c.statuses && !c.statuses.has(n.status)) return false;
  if (c.mode && n.mode !== c.mode) return false;
  if (c.purpose && n.purpose !== c.purpose) return false;
  if (c.kinds && !c.kinds.has(n.kind)) return false;
  if (c.policy && (c.policy === 'none' ? n.policyKey !== null : n.policyKey !== c.policy)) return false;
  if (c.repository && n.repositoryId !== c.repository) return false;
  if (c.scanId && n.scanId !== c.scanId) return false;
  if (c.q && !haystack(n).includes(c.q)) return false;
  return true;
}

// -- cadastro atual (linhas antigas) -------------------------------------------------------------

/**
 * Consulta ao cadastro atual para completar as linhas antigas: análises (que podem ter sido
 * excluídas), agendamentos (nome pelo id), repositórios e conexões. Guarda o que já consultou.
 */
export class DeletionResolver {
  constructor(store) {
    this.store = store;
    this.scans = new Map();
    this.schedules = new Map(store.listSchedules().map((s) => [s.id, s]));
    this.schedulesByName = new Map(); // só os nomes sem repetição
    for (const s of store.listSchedules()) this.schedulesByName.set(s.name, this.schedulesByName.has(s.name) ? null : s);
    this.repositories = new Map(store.listRepositories().map((r) => [r.id, r]));
    this.sources = new Map(store.listMailSources().map((s) => [s.id, s]));
  }

  /** A análise como está no cadastro (null se foi excluída). */
  scan(id) {
    if (!this.scans.has(id)) {
      const scan = id ? this.store.getScan(id) : null;
      this.scans.set(
        id,
        scan
          ? {
              name: scan.name || '',
              kind: SCAN_KINDS.has(scan.kind) ? scan.kind : 'files',
              purpose: scan.retention ? 'retention' : scan.fileTypes?.all ? 'date' : scan.fileTypes ? 'types' : 'terms',
              scheduleId: scan.scheduleId || null,
              scheduleName: scan.scheduleName || null,
              retention: scan.retention || null,
              repositories: scan.summary?.repositories || [],
              sources: scan.summary?.sources || [],
            }
          : null,
      );
    }
    return this.scans.get(id);
  }

  /** Nome atual de um repositório ou conexão (ou o do resumo da análise). */
  targetName(id, scan) {
    return this.repositories.get(id)?.name || this.sources.get(id)?.name || scan?.repositories.find((r) => r.id === id)?.name || scan?.sources.find((s) => s.id === id)?.name || '';
  }
}

// -- evento para exibição ------------------------------------------------------------------------

/** Caminho no formato do Windows ou do Linux (pastas locais). */
const looksLikePath = (item) => /^([a-z]:[\\/]|\\\\|\/)/i.test(item);

/** OneDrive ou SharePoint pelo endereço (ou pelos repositórios da análise). */
function cloudKind(item, scan) {
  const types = new Set((scan?.repositories || []).map((r) => r.type).filter((t) => t === 'onedrive' || t === 'sharepoint'));
  if (types.size === 1) return [...types][0];
  if (PERSONAL_URL_RE.test(item)) return 'onedrive';
  if (/^https?:\/\/[^/]*sharepoint\./i.test(item)) return 'sharepoint';
  return 'cloud';
}

/** Tipo de local do item: o gravado no evento ou, nas linhas antigas, o deduzido da análise e do item. */
function kindOf(e, target, scanKind, scan, item) {
  if (TARGET_TYPES.has(target.type)) return target.type;
  if (target.type === 'cloud') return cloudKind(item, scan);
  if (scanKind === 'mail') return 'mail';
  if (scanKind === 'teams' || e.method === 'teams') return 'teams';
  if (/^https?:\/\//i.test(item)) return cloudKind(item, scan);
  if (e.method === 'file' || looksLikePath(item)) return 'local';
  // Análise excluída: "caixa › pasta › assunto" (a caixa é só um endereço) é uma mensagem; outro
  // "a › b" é um arquivo na nuvem ("OneDrive de ana@... › pasta/arquivo": OneDrive).
  if (!scanKind && item.includes(SEP)) {
    const parts = item.split(SEP);
    if (parts.length >= 3 && /^[^\s@]+@[^\s@]+$/.test(parts[0])) return 'mail';
    return /^OneDrive de /i.test(parts[0]) ? 'onedrive' : 'cloud';
  }
  if (scanKind === 'files') return (scan?.repositories || []).every((r) => !r.type || r.type === 'local') ? 'local' : cloudKind(item, scan);
  return '';
}

const sameType = (kind, type = 'local') => (kind === 'cloud' ? type === 'onedrive' || type === 'sharepoint' : kind === type);
const normPath = (p) => str(p).replace(/\\/g, '/').replace(/\/+$/, '').toLowerCase();

/**
 * Repositório ou conexão do item: o gravado no evento ou, nas linhas antigas, o da análise (o único
 * do tipo, a pasta que contém o arquivo ou a conexão que tem a caixa). relativePath: nas pastas do
 * Windows, o caminho a partir da pasta do repositório.
 */
function repositoryOf(target, kind, scan, resolver, item, mailbox) {
  const id = str(target.repositoryId) || str(target.sourceId);
  if (id) {
    // O nome atual (o cadastro pode ter sido renomeado); se não existe mais, o gravado na linha.
    const recorded = str(target.repositoryName) || str(target.sourceName);
    const current = resolver.repositories.get(id)?.name || resolver.sources.get(id)?.name || '';
    return { id, name: current || recorded || resolver.targetName(id, scan) || UNKNOWN_REPOSITORY, former: current && recorded && recorded !== current ? recorded : '' };
  }
  const unknown = { id: null, name: UNKNOWN_REPOSITORY };
  if (!scan || !kind) return unknown;
  if (kind === 'mail' || kind === 'teams') {
    if (scan.sources.length === 1) return { id: scan.sources[0].id, name: resolver.targetName(scan.sources[0].id, scan) || scan.sources[0].name };
    if (kind === 'mail' && mailbox) {
      const address = mailbox.toLowerCase();
      const found = scan.sources.filter((s) => (resolver.sources.get(s.id)?.mailboxes || []).some((m) => str(m?.address || m).toLowerCase() === address));
      if (found.length === 1) return { id: found[0].id, name: resolver.targetName(found[0].id, scan) || found[0].name };
    }
    return unknown;
  }
  const repos = scan.repositories.filter((r) => sameType(kind, r.type));
  let repo = repos.length === 1 ? repos[0] : null;
  if (kind === 'local') {
    // A pasta mais longa que contém o arquivo.
    const file = normPath(item);
    const inside = repos.filter((r) => r.path && (file === normPath(r.path) || file.startsWith(`${normPath(r.path)}/`))).sort((a, b) => b.path.length - a.path.length);
    if (inside.length) repo = inside[0];
    if (repo?.path && normPath(item).startsWith(`${normPath(repo.path)}/`)) {
      return { id: repo.id, name: resolver.targetName(repo.id, scan) || repo.name, relativePath: item.slice(repo.path.replace(/[\\/]+$/, '').length + 1) };
    }
  }
  return repo ? { id: repo.id, name: resolver.targetName(repo.id, scan) || repo.name } : unknown;
}

/** Pasta de um caminho relativo ("a/b/c.txt" -> "a/b"). */
function dirOf(rel) {
  const idx = Math.max(rel.lastIndexOf('\\'), rel.lastIndexOf('/'));
  return idx === -1 ? '' : rel.slice(0, idx);
}

/** Data "AAAA-MM-DD" ou ISO, por extenso (dd/mm/aaaa). */
function dateText(value) {
  const m = DAY_RE.exec(str(value));
  if (m) return `${m[3]}/${m[2]}/${m[1]}`;
  const ms = Date.parse(value);
  return Number.isFinite(ms) ? new Date(ms).toLocaleDateString('pt-BR') : str(value);
}

/** Regra da política de retenção, curta: "Última modificação: mais de 5 anos". */
function policyRule(policy, scanKind) {
  if (!policy) return '';
  const criteria = scanKind === 'mail' ? MAIL_CRITERIA : FILE_CRITERIA;
  const label = criteria[policy.criterion]?.label || '';
  const period = policy.from && policy.to ? `entre ${dateText(policy.from)} e ${dateText(policy.to)}` : policy.from ? `a partir de ${dateText(policy.from)}` : policy.to ? `até ${dateText(policy.to)}` : '';
  const age = period || (policy.since ? `antes de ${dateText(policy.since)}` : Number.isFinite(policy.amount) && policy.unit ? `mais de ${amountText(policy)}` : '');
  const rule = [label, age].filter(Boolean).join(': ');
  // Pastas escolhidas na política: [{ repositoryId, path }] (ou só os caminhos).
  const names = Array.isArray(policy.folders) ? policy.folders.map((f) => (typeof f === 'string' ? f : str(f?.path))).filter(Boolean) : [];
  const folders = names.length ? `pastas: ${names.slice(0, 5).join(', ')}${names.length > 5 ? '…' : ''}` : '';
  return [rule, folders].filter(Boolean).join(' · ');
}

/**
 * Evento do registro pronto para a tela e as exportações. Campos:
 *   at, time (ms), status, statusLabel, kind, kindLabel, item, relativePath, folder, size,
 *   criterionDate, repositoryId, repositoryName, mailbox, scheduleId, scheduleName, policyKey
 *   (agendamento: o id, ou "nome:..." quando só o nome é conhecido; null sem agendamento), purpose,
 *   purposeLabel, policyRule, criterion, cutoff, since, periodFrom, periodTo, mode, originLabel, method,
 *   methodLabel, how,
 *   by, error, note, detail, scanId, scanName, scanKind, scanExists, recordId
 */
export function normalizeEvent(e, resolver) {
  const target = e.target && typeof e.target === 'object' && !Array.isArray(e.target) ? e.target : {};
  const scanId = str(e.scanId);
  const scan = resolver.scan(scanId);
  const item = str(e.item);
  const by = str(e.by);
  const status = str(e.status);
  const mode = MODES.includes(e.mode) ? e.mode : '';
  const method = str(e.method);

  // Agendamento ou política: o gravado na linha, o da análise ou (análise excluída) o citado em "por".
  const enriched = Object.hasOwn(e, 'scheduleId');
  let scheduleId = str(e.scheduleId) || (!enriched && scan?.scheduleId) || null;
  let scheduleName = str(e.scheduleName) || (!enriched && scan?.scheduleName) || null;
  let byPurpose = '';
  if (!enriched && !scan && !scheduleId) {
    const m = BY_SCHEDULE_RE.exec(by);
    if (m) {
      scheduleName = m[2];
      scheduleId = resolver.schedulesByName.get(m[2])?.id || null;
      byPurpose = m[1] === 'política de retenção' ? 'retention' : resolver.schedulesByName.get(m[2])?.purpose || '';
    }
  }
  const schedule = scheduleId ? resolver.schedules.get(scheduleId) : null;
  // O nome atual do agendamento ou da política (o da linha pode ser anterior a uma renomeação).
  const formerSchedule = schedule && scheduleName && scheduleName !== schedule.name ? scheduleName : '';
  if (schedule) scheduleName = schedule.name;
  const policyKey = scheduleId || (scheduleName ? `nome:${scheduleName}` : null);

  const scanKind = SCAN_KINDS.has(e.scanKind) ? e.scanKind : scan?.kind || (method === 'teams' ? 'teams' : '');
  const purpose = PURPOSES.includes(e.purpose) ? e.purpose : scan?.purpose || byPurpose || (mode === 'retention' ? 'retention' : '');
  const kind = kindOf(e, target, scanKind, scan, item);

  // Caixa e pasta (e-mail: "caixa › pasta › assunto"; Teams: "conversa › assunto").
  const parts = item.split(SEP);
  const mailbox = str(target.mailbox) || (kind === 'mail' && parts.length > 2 ? parts[0] : '');
  const repo = repositoryOf(target, kind, scan, resolver, item, mailbox);
  const relativePath = str(target.relativePath) || repo.relativePath || '';
  let folder = str(target.folder);
  if (!folder) {
    if (kind === 'mail') folder = parts.length > 2 ? parts[1] : '';
    else if (kind === 'teams') folder = parts.length > 1 ? parts.slice(0, -1).join(SEP) : '';
    // OneDrive e SharePoint: com a conta (ou o site) e a biblioteca, como nos relatórios.
    else folder = [str(target.account), str(target.library), dirOf(relativePath)].filter(Boolean).join(SEP);
  }

  const policy = e.policy && typeof e.policy === 'object' ? e.policy : purpose === 'retention' ? scan?.retention || null : null;
  const size = Number.isFinite(target.size) && target.size >= 0 ? target.size : null;
  const at = str(e.at);
  const labels = kind === 'mail' || kind === 'teams' ? MAIL_DELETION_LABELS : DELETION_LABELS;
  const error = str(e.error);
  const note = str(e.note);
  return {
    at,
    time: Date.parse(at),
    status,
    statusLabel: labels[status] || status || '—',
    kind,
    kindLabel: KIND_LABELS[kind] ?? kind,
    item,
    relativePath,
    folder,
    size,
    criterionDate: str(target.criterionDate) || null,
    repositoryId: repo.id,
    repositoryName: repo.name,
    mailbox,
    scheduleId,
    scheduleName,
    policyKey,
    purpose,
    purposeLabel: PURPOSE_LABELS[purpose] ?? purpose,
    policyRule: policyRule(policy, scanKind),
    criterion: str(policy?.criterion) || null,
    cutoff: str(policy?.cutoff) || null,
    since: str(policy?.since) || null,
    // Período da política (a data do critério entre os dois dias), quando ela usa um.
    periodFrom: str(policy?.from) || null,
    periodTo: str(policy?.to) || null,
    mode,
    originLabel: ORIGIN_LABELS[mode],
    method,
    methodLabel: METHOD_LABELS[method] || method,
    how: mode ? deletionHow({ mode, by }) : '',
    by,
    error,
    note,
    detail: error || note,
    scanId,
    scanName: str(e.scanName) || scan?.name || '',
    scanKind,
    scanExists: Boolean(scan),
    recordId: e.recordId ?? null,
    // Nomes anteriores (renomeados depois da exclusão): a busca encontra também por eles.
    formerNames: [formerSchedule, repo.former].filter(Boolean).join(' | '),
  };
}

/** O evento como vai para o navegador (sem campos internos). */
export function publicEvent(n) {
  const { time, formerNames, ...rest } = n;
  return rest;
}

// -- resumo --------------------------------------------------------------------------------------

const newStats = () => ({ attempts: 0, deleted: 0, missing: 0, changed: 0, failed: 0, uniqueDeleted: 0, bytesDeleted: 0 });

function bump(stats, n, unique) {
  stats.attempts++;
  if (Object.hasOwn(stats, n.status) && STATUSES.includes(n.status)) stats[n.status]++;
  if (unique) {
    stats.uniqueDeleted++;
    if (n.size !== null) stats.bytesDeleted += n.size;
  }
}

const byCount = (a, b) => b.uniqueDeleted - a.uniqueDeleted || b.attempts - a.attempts;
const byText = (a, b) => a.localeCompare(b, 'pt-BR');

/** Chaves dos itens únicos (análise + item), as mesmas no registro inteiro e no recorte. */
class ItemKeys {
  constructor() {
    this.scanIndex = new Map();
    this.loose = 0;
  }

  /** Chave do item (número quando possível: ocupa menos memória com milhões de itens). */
  of(n) {
    if (n.recordId === null || n.recordId === undefined || n.recordId === '') return `#${this.loose++}`;
    let idx = this.scanIndex.get(n.scanId);
    if (idx === undefined) {
      idx = this.scanIndex.size;
      this.scanIndex.set(n.scanId, idx);
    }
    const r = n.recordId;
    return Number.isSafeInteger(r) && r >= 0 && r < 2 ** 31 && idx < 2 ** 21 ? idx * 2 ** 31 + r : `${idx}:${r}`;
  }
}

/** Rótulos das situações: no feminino só quando o recorte é de mensagens (e-mail ou Teams). */
export const statusLabelsFor = (filters) => (filters.kind === 'mail' || filters.kind === 'teams' ? MAIL_DELETION_LABELS : DELETION_LABELS);

/**
 * Resumo dos eventos que passaram nos filtros. Itens únicos (análise + item): vale a exclusão, se
 * houve uma, ou a última tentativa (a mesma regra do relatório). O espaço liberado soma o tamanho
 * dos itens excluídos (uma vez cada); os excluídos sem o tamanho registrado são contados à parte.
 */
class Summary {
  constructor(resolver, keys, labels = DELETION_LABELS) {
    this.resolver = resolver;
    this.keys = keys;
    this.labels = labels;
    this.totals = { ...newStats(), uniqueItems: 0, uniqueMissing: 0, uniquePending: 0, bytesUnknown: 0 };
    this.first = null;
    this.last = null;
    this.days = new Map();
    this.policies = new Map();
    this.repositories = new Map();
    this.kinds = new Map();
    this.states = new Map(); // item único -> situação final (dentro do recorte)
  }

  group(map, key, init) {
    let g = map.get(key);
    if (!g) {
      g = { ...init(), ...newStats() };
      map.set(key, g);
    }
    return g;
  }

  add(n, key) {
    const previous = this.states.get(key);
    if (previous === undefined) this.totals.uniqueItems++;
    let unique = false;
    if (previous !== 'deleted') {
      this.states.set(key, n.status);
      unique = n.status === 'deleted';
    }
    if (unique && n.size === null) this.totals.bytesUnknown++;
    bump(this.totals, n, unique);
    if (Number.isFinite(n.time)) {
      if (!this.first || n.time < this.first.time) this.first = n;
      if (!this.last || n.time >= this.last.time) this.last = n;
      const day = localDay(n.time);
      bump(this.group(this.days, day, () => ({ day })), n, unique);
    }
    const policy = this.group(this.policies, n.policyKey, () => ({ key: n.policyKey, scheduleId: n.scheduleId, name: n.scheduleName || NO_POLICY, purpose: n.purpose }));
    if (n.scheduleName) policy.name = n.scheduleName; // o nome mais recente
    if (n.purpose) policy.purpose = n.purpose;
    bump(policy, n, unique);
    const repo = this.group(this.repositories, n.repositoryId, () => ({ key: n.repositoryId, id: n.repositoryId, name: n.repositoryName, kind: n.kind }));
    if (n.repositoryId) repo.name = n.repositoryName;
    if (n.kind && n.kind !== 'cloud') repo.kind = n.kind;
    bump(repo, n, unique);
    bump(this.group(this.kinds, n.kind, () => ({ kind: n.kind })), n, unique);
  }

  /**
   * finals: a situação final de cada item pelo registro inteiro (e não só pelo recorte): um item que
   * falhou e depois foi excluído não "continua no lugar", mesmo filtrando só as falhas.
   */
  result(finals = null) {
    const t = this.totals;
    for (const [key, own] of this.states) {
      const status = finals ? finals.get(key) : own;
      if (status === 'missing') t.uniqueMissing++;
      else if (status === 'failed' || status === 'changed') t.uniquePending++;
    }
    this.states.clear();
    const { schedules } = this.resolver;
    const byStatus = STATUSES.map((status) => ({ status, label: this.labels[status], count: t[status] }));
    return {
      totals: t,
      range: { first: this.first?.at || null, last: this.last?.at || null },
      byDay: [...this.days.values()].sort((a, b) => a.day.localeCompare(b.day)),
      byPolicy: [...this.policies.values()]
        .map((g) => ({ ...g, name: (g.scheduleId && schedules.get(g.scheduleId)?.name) || g.name, purposeLabel: PURPOSE_LABELS[g.purpose] ?? g.purpose, exists: Boolean(g.scheduleId && schedules.has(g.scheduleId)) }))
        .sort((a, b) => (a.key === null) - (b.key === null) || byCount(a, b) || byText(a.name, b.name)),
      byRepository: [...this.repositories.values()]
        .map((g) => ({ ...g, name: (g.id && this.resolver.targetName(g.id)) || g.name, kindLabel: KIND_LABELS[g.kind] ?? g.kind }))
        .sort((a, b) => (a.key === null) - (b.key === null) || byCount(a, b) || byText(a.name, b.name)),
      byStatus,
      byKind: [...this.kinds.values()].map((g) => ({ ...g, label: KIND_LABELS[g.kind] ?? g.kind })).sort(byCount),
    };
  }
}

/**
 * Opções dos filtros, pelo registro inteiro (sem os filtros): agendamentos e políticas citados nas
 * linhas — junto com os do cadastro atual, as políticas de retenção primeiro — e os repositórios e
 * conexões dos itens excluídos.
 */
class Options {
  constructor() {
    this.schedules = new Map();
    this.repositories = new Map();
    this.withoutSchedule = 0;
  }

  add(n) {
    if (n.policyKey === null) this.withoutSchedule++;
    else {
      const s = this.schedules.get(n.policyKey) || { id: n.policyKey, scheduleId: n.scheduleId, name: n.scheduleName, purpose: n.purpose, count: 0 };
      s.count++;
      if (n.scheduleName) s.name = n.scheduleName;
      if (n.purpose) s.purpose = n.purpose;
      this.schedules.set(n.policyKey, s);
    }
    if (n.repositoryId) {
      const r = this.repositories.get(n.repositoryId) || { id: n.repositoryId, name: n.repositoryName, kind: n.kind, count: 0 };
      r.count++;
      r.name = n.repositoryName;
      if (n.kind && n.kind !== 'cloud') r.kind = n.kind;
      this.repositories.set(n.repositoryId, r);
    }
  }

  result(resolver) {
    const schedules = new Map(this.schedules);
    for (const s of resolver.schedules.values()) {
      const seen = schedules.get(s.id);
      schedules.set(s.id, { id: s.id, scheduleId: s.id, name: s.name, purpose: s.purpose || 'terms', count: seen?.count || 0 });
    }
    const rank = (s) => (s.purpose === 'retention' ? 0 : 1);
    return {
      schedules: [...schedules.values()]
        .map((s) => ({ ...s, purposeLabel: PURPOSE_LABELS[s.purpose] ?? s.purpose, exists: Boolean(s.scheduleId && resolver.schedules.has(s.scheduleId)) }))
        .sort((a, b) => rank(a) - rank(b) || byText(a.name || '', b.name || '')),
      repositories: [...this.repositories.values()]
        .map((r) => ({ ...r, name: resolver.targetName(r.id) || r.name, kindLabel: KIND_LABELS[r.kind] ?? r.kind, exists: resolver.repositories.has(r.id) || resolver.sources.has(r.id) }))
        .sort((a, b) => byText(a.name, b.name)),
      withoutSchedule: this.withoutSchedule,
    };
  }
}

// -- consulta com cache --------------------------------------------------------------------------

const READ_BATCH = 1000;

/**
 * Consulta ao registro geral: uma passada pelo arquivo filtra os eventos (guarda só a posição de
 * cada linha) e calcula o resumo e as opções dos filtros. O resultado fica em cache (pelo tamanho e
 * pela data do arquivo, pelo cadastro e pelos filtros): paginar ou exportar o mesmo recorte não
 * relê o arquivo inteiro.
 */
export class DeletionLog {
  constructor(store, { cacheSize = 6, concurrency = 2 } = {}) {
    this.store = store;
    this.cacheSize = cacheSize;
    this.cache = new Map();
    // Passadas pelo arquivo ao mesmo tempo (cada uma lê o registro inteiro): as demais esperam a vez.
    this.concurrency = concurrency;
    this.running = 0;
    this.waiting = [];
  }

  /** Executa uma passada quando houver vaga (no máximo `concurrency` ao mesmo tempo). */
  async #slot(fn) {
    if (this.running >= this.concurrency) await new Promise((resolve) => this.waiting.push(resolve));
    this.running++;
    try {
      return await fn();
    } finally {
      this.running--;
      this.waiting.shift()?.();
    }
  }

  async #stat() {
    await this.store.flushDeletionLog();
    try {
      const st = await fs.stat(this.store.deletionLog);
      return st.isFile() ? { size: st.size, mtimeMs: st.mtimeMs } : { size: 0, mtimeMs: 0 };
    } catch (err) {
      if (err.code === 'ENOENT' || err.code === 'ENOTDIR') return { size: 0, mtimeMs: 0 };
      throw err;
    }
  }

  /** Assinatura do cadastro usado para completar as linhas antigas (nomes, agendamentos, análises). */
  #catalog() {
    const s = this.store;
    const parts = [
      s.listScans().map((x) => `${x.id}:${x.name}:${x.scheduleName || ''}`),
      s.listSchedules().map((x) => `${x.id}:${x.name}:${x.purpose || ''}`),
      s.listRepositories().map((x) => `${x.id}:${x.name}:${x.type || ''}:${x.path}`),
      s.listMailSources().map((x) => `${x.id}:${x.name}:${(x.mailboxes || []).map((m) => str(m?.address || m).toLowerCase()).join(',')}`),
    ];
    return crypto.createHash('sha1').update(parts.map((p) => p.join('|')).join('#')).digest('base64');
  }

  /** Passada pelo registro com os filtros (já validados por parseDeletionFilters). */
  query(filters = {}) {
    return this.#stat().then((st) => {
      // A busca entra na chave como é comparada (sem acentos nem maiúsculas): "Arquivo" e "arquivo"
      // são o mesmo recorte e usam o mesmo cache.
      const canonical = { ...filters, ...(filters.q ? { q: foldText(filters.q).trim() } : {}) };
      const key = `${st.size}|${st.mtimeMs}|${this.#catalog()}|${JSON.stringify(canonical)}`;
      let entry = this.cache.get(key);
      if (entry) {
        this.cache.delete(key);
      } else {
        entry = this.#slot(() => this.#run(filters, st));
        entry.catch(() => {
          if (this.cache.get(key) === entry) this.cache.delete(key);
        });
      }
      this.cache.set(key, entry);
      while (this.cache.size > this.cacheSize) this.cache.delete(this.cache.keys().next().value);
      return entry;
    });
  }

  async #run(filters, st) {
    const resolver = new DeletionResolver(this.store);
    const compiled = compileFilters(filters);
    const keys = new ItemKeys();
    // Situação final de cada item pelo registro inteiro. Sem filtros, o recorte é o registro inteiro e
    // o próprio resumo já a tem (sem um segundo mapa do mesmo tamanho).
    const everything = Object.keys(filters).length === 0;
    const finals = everything ? null : new Map();
    const summary = new Summary(resolver, keys, statusLabelsFor(filters));
    const options = new Options();
    const offsets = [];
    const lengths = [];
    if (st.size > 0) {
      for await (const { offset, length, event } of this.store.readDeletionLog({ end: st.size })) {
        const n = normalizeEvent(event, resolver);
        options.add(n);
        const key = keys.of(n);
        if (finals && finals.get(key) !== 'deleted') finals.set(key, n.status);
        if (!matches(n, compiled)) continue;
        offsets.push(offset);
        lengths.push(length);
        summary.add(n, key);
      }
    }
    return { filters, size: st.size, total: offsets.length, offsets, lengths, resolver, summary: summary.result(finals), options: options.result(resolver) };
  }

  /** Posições das linhas do recorte, das mais recentes para as mais antigas (start: quantas pular). */
  #refs(result, start, count) {
    const refs = [];
    for (let i = result.total - 1 - start; i >= 0 && refs.length < count; i--) refs.push({ offset: result.offsets[i], length: result.lengths[i] });
    return refs;
  }

  /** Uma página do recorte (as mais recentes primeiro), já no formato de exibição. */
  async page(result, page, pageSize) {
    const events = await this.store.readDeletionLines(this.#refs(result, (page - 1) * pageSize, pageSize));
    return events.filter(Boolean).map((e) => normalizeEvent(e, result.resolver));
  }

  /** Todos os eventos do recorte (as mais recentes primeiro), lidos aos poucos (exportações). */
  async *events(result, { limit = Infinity } = {}) {
    const count = Math.min(result.total, limit);
    for (let done = 0; done < count; done += READ_BATCH) {
      const events = await this.store.readDeletionLines(this.#refs(result, done, Math.min(READ_BATCH, count - done)));
      for (const e of events) if (e) yield normalizeEvent(e, result.resolver);
    }
  }
}
