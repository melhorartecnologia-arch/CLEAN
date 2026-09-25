// Rótulos, filtros e agregações dos resultados (usados pela API e pelas exportações).
import { foldText } from '../scan/matcher.js';

export const SOURCE_LABELS = {
  audit: 'Log de auditoria',
  metadata: 'Metadados do documento',
  owner: 'Proprietário do arquivo',
};

export const STATUS_LABELS = {
  ok: 'Analisado',
  partial: 'Analisado parcialmente',
  encrypted: 'Protegido por senha',
  unsupported: 'Formato sem texto',
  'skipped-size': 'Muito grande',
  empty: 'Vazio',
  error: 'Erro de leitura',
  'not-requested': 'Não verificado',
};

export const LOCATION_LABELS = { name: 'Nome', content: 'Conteúdo' };

export const SCAN_STATUS_LABELS = {
  queued: 'Na fila',
  running: 'Em andamento',
  completed: 'Concluída',
  cancelled: 'Cancelada',
  failed: 'Falhou',
  interrupted: 'Interrompida',
};

/** Texto usado na busca livre de um registro (em minúsculas e sem acentos). */
function haystack(record) {
  if (!record._search) {
    record._search = foldText(
      [record.path, record.lastUser, record.owner, record.metadata?.lastModifiedBy, record.metadata?.author, ...record.terms].filter(Boolean).join(' | '),
    );
  }
  return record._search;
}

const SORTERS = {
  path: (a, b) => a.path.localeCompare(b.path, 'pt-BR'),
  name: (a, b) => a.name.localeCompare(b.name, 'pt-BR'),
  modified: (a, b) => String(a.modified).localeCompare(String(b.modified)),
  occurrences: (a, b) => a.occurrences - b.occurrences,
  terms: (a, b) => a.terms.length - b.terms.length,
  size: (a, b) => a.size - b.size,
  lastUser: (a, b) => String(a.lastUser || '').localeCompare(String(b.lastUser || ''), 'pt-BR'),
};

/** Parâmetros de filtro aceitos pela API (os demais são ignorados). */
export const FILTER_KEYS = ['q', 'term', 'user', 'repository', 'location', 'extension', 'status', 'sort', 'dir', 'page'];

/**
 * Filtra e ordena os registros.
 * filters: { q, term, user, repository, location, extension, status, sort, dir }
 */
export function filterRecords(records, filters = {}) {
  const q = filters.q ? foldText(filters.q) : '';
  let out = records.filter((r) => {
    if (filters.term && !r.terms.includes(filters.term)) return false;
    if (filters.user && (r.lastUser || '') !== filters.user) return false;
    if (filters.repository && r.repositoryId !== filters.repository) return false;
    if (filters.location && !r.matches.some((m) => m.location === filters.location)) return false;
    if (filters.extension && r.extension !== filters.extension) return false;
    if (filters.status && r.contentStatus !== filters.status) return false;
    if (q && !haystack(r).includes(q)) return false;
    return true;
  });
  const sorter = Object.hasOwn(SORTERS, filters.sort || '') ? SORTERS[filters.sort] : SORTERS.path;
  out = out.slice().sort(sorter);
  if (filters.dir === 'desc') out.reverse();
  return out;
}

/** Remove campos internos antes de enviar ao navegador. */
export function publicRecord(record) {
  const { _search, ...rest } = record;
  return rest;
}

function countBy(records, keyFn) {
  const map = new Map();
  for (const r of records) {
    for (const key of [].concat(keyFn(r))) {
      if (key === undefined) continue;
      map.set(key, (map.get(key) || 0) + 1);
    }
  }
  return map;
}

/** Agregações para o painel do relatório. */
export function summarize(records) {
  const terms = new Map();
  for (const r of records) {
    for (const m of r.matches) {
      const key = m.termId;
      let t = terms.get(key);
      if (!t) {
        t = { termId: m.termId, term: m.term, list: m.list, kind: m.kind, files: new Set(), occurrences: 0, inName: 0, inContent: 0 };
        terms.set(key, t);
      }
      t.files.add(r.id);
      t.occurrences += m.count;
      if (m.location === 'name') t.inName++;
      else t.inContent++;
    }
  }
  const byTerm = [...terms.values()]
    .map(({ files, ...t }) => ({ ...t, files: files.size }))
    .sort((a, b) => b.files - a.files || b.occurrences - a.occurrences);

  const users = new Map();
  for (const r of records) {
    const key = r.lastUser || '(não identificado)';
    let u = users.get(key);
    if (!u) {
      u = { user: key, identified: Boolean(r.lastUser), files: 0, occurrences: 0, sources: {} };
      users.set(key, u);
    }
    u.files++;
    u.occurrences += r.occurrences;
    if (r.lastUserSource) u.sources[r.lastUserSource] = (u.sources[r.lastUserSource] || 0) + 1;
  }
  const byUser = [...users.values()].sort((a, b) => b.files - a.files);

  const toList = (map, label) => [...map.entries()].map(([key, files]) => ({ [label]: key, files })).sort((a, b) => b.files - a.files);
  return {
    files: records.length,
    occurrences: records.reduce((sum, r) => sum + r.occurrences, 0),
    byTerm,
    byUser,
    byRepository: toList(countBy(records, (r) => r.repositoryName), 'repository'),
    byExtension: toList(countBy(records, (r) => r.extension || '(sem extensão)'), 'extension'),
    byLocation: {
      name: records.filter((r) => r.matches.some((m) => m.location === 'name')).length,
      content: records.filter((r) => r.matches.some((m) => m.location === 'content')).length,
    },
    bySource: Object.fromEntries(countBy(records, (r) => r.lastUserSource || 'none')),
    byStatus: Object.fromEntries(countBy(records, (r) => r.contentStatus || 'none')),
  };
}

/** Exemplo de ocorrência em uma linha de texto: "... antes [trecho] depois ...". */
export function sampleText(sample) {
  if (!sample) return '';
  const where = sample.where ? ` (${sample.where})` : '';
  return `${sample.before}[${sample.match}]${sample.after}${where}`;
}

export function formatDateTime(iso) {
  if (!iso) return '';
  const d = new Date(iso);
  return Number.isNaN(d.getTime()) ? '' : d.toLocaleString('pt-BR');
}

export function auditText(audit) {
  if (!audit) return '';
  return `${audit.user} – ${audit.action} em ${formatDateTime(audit.time)}`;
}
