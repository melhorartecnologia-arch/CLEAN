// Rótulos, filtros e agregações dos resultados (usados pela API e pelas exportações).
import { foldText } from '../scan/matcher.js';
import { ageBucket, AGE_BUCKETS } from '../retention/policy.js';

export const SOURCE_LABELS = {
  audit: 'Log de auditoria',
  cloud: 'Microsoft 365 (quem alterou por último)',
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

export const LOCATION_LABELS = {
  name: 'Nome',
  content: 'Conteúdo',
  // análises de e-mail
  subject: 'Assunto',
  body: 'Corpo',
  attachmentName: 'Nome do anexo',
  attachment: 'Conteúdo do anexo',
  address: 'Remetente/destinatários',
};

export const MAIL_TYPE_LABELS = { graph: 'Microsoft 365', gmail: 'Google Workspace', imap: 'IMAP' };

export const DELETION_LABELS = {
  deleted: 'Excluído',
  missing: 'Não encontrado (já excluído ou movido)',
  changed: 'Não excluído (alterado depois da análise)',
  failed: 'Falha na exclusão',
};

export const MAIL_DELETION_LABELS = {
  deleted: 'Excluída',
  missing: 'Não encontrada (já excluída ou movida)',
  changed: 'Não excluída (alterada depois da análise)',
  failed: 'Falha na exclusão',
};

/** Como a exclusão foi feita (automática, com quem iniciou a análise, ou manual, com quem excluiu). */
export function deletionHow(d) {
  if (d.mode === 'retention') return `exclusão pela política de retenção${d.by ? ` (${d.by})` : ''}`;
  return d.mode === 'auto' ? `exclusão automática da análise${d.by ? ` iniciada por ${d.by}` : ''}` : `exclusão manual${d.by ? ` por ${d.by}` : ''}`;
}

/** Situação da exclusão de um item, em uma linha (relatórios e exportações). kind: 'files' ou 'mail'. */
export function deletionText(d, kind = 'files') {
  if (!d) return '';
  const o = kind === 'mail' ? 'a' : 'o'; // "arquivo excluído", "mensagem excluída"
  const when = formatDateTime(d.at);
  const how = deletionHow(d);
  if (d.status === 'deleted') {
    const where = d.note ? ` (${d.note})` : d.method === 'trash' ? ` (movid${o} para a lixeira)` : '';
    return `Excluíd${o} em ${when}${where} — ${how}`;
  }
  if (d.status === 'missing') return `Não encontrad${o} em ${when} (já excluíd${o} ou movid${o}) — ${how}`;
  if (d.status === 'changed') return `Não excluíd${o} em ${when}: ${d.error || `alterad${o} depois da análise`} — ${how}`;
  return `Falha ao excluir em ${when}: ${d.error || 'erro desconhecido'} — ${how}`;
}

/**
 * Aplica aos registros a situação de exclusão de cada um (registro deletions.ndjson): vale a última
 * tentativa, mas um item excluído continua excluído (uma tentativa posterior que não o encontre
 * mais, como dois pedidos ao mesmo tempo, não apaga essa informação).
 */
export function applyDeletions(records, deletions) {
  const latest = new Map();
  for (const d of deletions) {
    if (latest.get(d.recordId)?.status === 'deleted') continue;
    latest.set(d.recordId, d);
  }
  for (const r of records) r.deletion = latest.get(r.id) || null;
  return records;
}

/**
 * Filtro "Exclusão": deleted = excluídos; missing = não encontrados na hora da exclusão; failed = não
 * excluídos por falha ou alteração; kept = continuam no lugar (sem exclusão, com falha ou alterados).
 */
function matchesDeletion(record, value) {
  const status = record.deletion?.status;
  if (value === 'deleted') return status === 'deleted';
  if (value === 'missing') return status === 'missing';
  if (value === 'failed') return status === 'failed' || status === 'changed';
  if (value === 'kept') return status !== 'deleted' && status !== 'missing';
  return true;
}

export function deletionTotals(records) {
  const totals = { deleted: 0, missing: 0, changed: 0, failed: 0 };
  for (const r of records) if (r.deletion && totals[r.deletion.status] !== undefined) totals[r.deletion.status]++;
  return totals;
}

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
      [
        record.path,
        record.lastUser,
        record.owner,
        record.metadata?.lastModifiedBy,
        record.metadata?.author,
        record.cloud?.account,
        record.cloud?.accountName,
        record.cloud?.library,
        record.cloud?.lastModifiedBy?.name,
        record.cloud?.lastModifiedBy?.email,
        record.cloud?.createdBy?.name,
        record.cloud?.createdBy?.email,
        ...record.terms,
      ]
        .filter(Boolean)
        .join(' | '),
    );
  }
  return record._search;
}

/** Retenção: data do critério (os mais antigos primeiro). */
const byRetention = (a, b) => String(a.retention?.date || '').localeCompare(String(b.retention?.date || ''));

const SORTERS = {
  oldest: byRetention,
  path: (a, b) => a.path.localeCompare(b.path, 'pt-BR'),
  name: (a, b) => a.name.localeCompare(b.name, 'pt-BR'),
  modified: (a, b) => String(a.modified).localeCompare(String(b.modified)),
  occurrences: (a, b) => a.occurrences - b.occurrences,
  terms: (a, b) => a.terms.length - b.terms.length,
  size: (a, b) => a.size - b.size,
  lastUser: (a, b) => String(a.lastUser || '').localeCompare(String(b.lastUser || ''), 'pt-BR'),
};

/** Parâmetros de filtro aceitos pela API (os demais são ignorados). */
export const FILTER_KEYS = ['q', 'term', 'user', 'repository', 'location', 'extension', 'status', 'deletion', 'age', 'sort', 'dir', 'page'];

/** Faixa de idade (retenção) do item. */
const inAge = (r, key) => Boolean(r.retention) && ageBucket(r.retention.ageDays).key === key;

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
    if (filters.deletion && !matchesDeletion(r, filters.deletion)) return false;
    if (filters.age && !inAge(r, filters.age)) return false;
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

// ---------------------------------------------------------------------------------------------
// Análises de e-mail

function mailHaystack(record) {
  if (!record._search) {
    record._search = foldText(
      [record.subject, record.from, ...(record.to || []), ...(record.cc || []), record.mailbox, record.folder, ...(record.attachments || []).map((a) => a.name), ...record.terms]
        .filter(Boolean)
        .join(' | '),
    );
  }
  return record._search;
}

const byDate = (a, b) => String(a.date || '').localeCompare(String(b.date || ''));

const MAIL_SORTERS = {
  oldest: byRetention,
  date: byDate,
  mailbox: (a, b) => a.mailbox.localeCompare(b.mailbox, 'pt-BR') || byDate(a, b),
  sender: (a, b) => String(a.from || '').localeCompare(String(b.from || ''), 'pt-BR') || byDate(a, b),
  subject: (a, b) => String(a.subject || '').localeCompare(String(b.subject || ''), 'pt-BR'),
  occurrences: (a, b) => a.occurrences - b.occurrences,
  terms: (a, b) => a.terms.length - b.terms.length,
  size: (a, b) => a.size - b.size,
};

export const MAIL_FILTER_KEYS = ['q', 'term', 'mailbox', 'sender', 'location', 'source', 'deletion', 'age', 'sort', 'dir', 'page'];

/** Anexos "de verdade" (sem as imagens embutidas no corpo, como logotipos de assinatura). */
export function realAttachments(record) {
  return (record.attachments || []).filter((a) => !a.inline);
}

/**
 * Filtra e ordena as mensagens.
 * filters: { q, term, mailbox, sender, location, source, sort, dir } — padrão: mais recentes primeiro.
 */
export function filterMailRecords(records, filters = {}) {
  const q = filters.q ? foldText(filters.q) : '';
  let out = records.filter((r) => {
    if (filters.term && !r.terms.includes(filters.term)) return false;
    if (filters.mailbox && r.mailbox !== filters.mailbox) return false;
    if (filters.sender && (r.fromAddress || '') !== filters.sender) return false;
    if (filters.location && !r.matches.some((m) => m.location === filters.location)) return false;
    if (filters.source && r.sourceId !== filters.source) return false;
    if (filters.deletion && !matchesDeletion(r, filters.deletion)) return false;
    if (filters.age && !inAge(r, filters.age)) return false;
    if (q && !mailHaystack(r).includes(q)) return false;
    return true;
  });
  const known = Object.hasOwn(MAIL_SORTERS, filters.sort || '');
  out = out.slice().sort(known ? MAIL_SORTERS[filters.sort] : byDate);
  if (known ? filters.dir === 'desc' : filters.dir !== 'asc') out.reverse();
  return out;
}

/** Agregações do relatório de e-mail. */
export function summarizeMail(records) {
  const terms = new Map();
  for (const r of records) {
    for (const m of r.matches) {
      let t = terms.get(m.termId);
      if (!t) {
        t = { termId: m.termId, term: m.term, list: m.list, kind: m.kind, messages: new Set(), occurrences: 0, inSubject: 0, inBody: 0, inAttachments: 0 };
        terms.set(m.termId, t);
      }
      t.messages.add(r.id);
      t.occurrences += m.count;
      if (m.location === 'subject') t.inSubject++;
      else if (m.location === 'body') t.inBody++;
      else if (m.location === 'attachment' || m.location === 'attachmentName') t.inAttachments++;
    }
  }
  const byTerm = [...terms.values()]
    .map(({ messages, ...t }) => ({ ...t, messages: messages.size }))
    .sort((a, b) => b.messages - a.messages || b.occurrences - a.occurrences);

  const group = (keyFn, make) => {
    const map = new Map();
    for (const r of records) {
      const k = keyFn(r);
      let g = map.get(k);
      if (!g) {
        g = { ...make(r), messages: 0, occurrences: 0 };
        map.set(k, g);
      }
      g.messages++;
      g.occurrences += r.occurrences;
    }
    return [...map.values()].sort((a, b) => b.messages - a.messages || b.occurrences - a.occurrences);
  };
  const locations = Object.fromEntries(['subject', 'body', 'attachmentName', 'attachment', 'address'].map((k) => [k, records.filter((r) => r.matches.some((m) => m.location === k)).length]));
  return {
    messages: records.length,
    occurrences: records.reduce((sum, r) => sum + r.occurrences, 0),
    withAttachments: records.filter((r) => realAttachments(r).length > 0).length,
    byTerm,
    byMailbox: group(
      (r) => r.mailbox,
      (r) => ({ mailbox: r.mailbox, name: r.mailboxName || '' }),
    ),
    bySender: group(
      (r) => r.fromAddress || '',
      (r) => ({ sender: r.fromAddress || '', label: r.from || '(sem remetente)' }),
    ),
    bySource: group(
      (r) => r.sourceId,
      (r) => ({ sourceId: r.sourceId, source: r.sourceName }),
    ),
    byLocation: locations,
    byStatus: Object.fromEntries(countBy(records, (r) => r.contentStatus || 'none')),
  };
}

// ---------------------------------------------------------------------------------------------
// Retenção

/** Itens expirados por faixa de idade (quantidade e tamanho), para o gráfico do relatório. */
export function summarizeRetention(records) {
  const buckets = new Map(AGE_BUCKETS.map((b) => [b.key, { key: b.key, label: b.label, count: 0, bytes: 0 }]));
  let bytes = 0;
  for (const r of records) {
    if (!r.retention) continue;
    const b = buckets.get(ageBucket(r.retention.ageDays).key);
    b.count++;
    b.bytes += Number(r.size) || 0;
    bytes += Number(r.size) || 0;
  }
  return { byAge: [...buckets.values()].filter((b) => b.count > 0), bytes };
}
