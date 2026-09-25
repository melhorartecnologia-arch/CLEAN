// Exportação dos relatórios: Excel (.xlsx), CSV (padrão Excel pt-BR), HTML para impressão e JSON.
// Tudo é gravado em fluxo contínuo na resposta, para suportar análises com muitos resultados.
import { once } from 'node:events';
import { writeXlsx } from './xlsx.js';
import {
  summarize,
  sampleText,
  formatDateTime,
  auditText,
  SOURCE_LABELS,
  STATUS_LABELS,
  LOCATION_LABELS,
  SCAN_STATUS_LABELS,
  DELETION_LABELS,
  deletionText,
} from './model.js';

export const toDate = (iso) => (iso ? new Date(iso) : null);
export const kb = (bytes) => Math.round((bytes / 1024) * 10) / 10;

function sortRecords(records) {
  return records.slice().sort((a, b) => a.repositoryName.localeCompare(b.repositoryName, 'pt-BR') || a.path.localeCompare(b.path, 'pt-BR'));
}

function locations(record) {
  return [...new Set(record.matches.map((m) => LOCATION_LABELS[m.location]))].join(' e ');
}

function folderOf(record) {
  const rel = record.relativePath || '';
  const idx = Math.max(rel.lastIndexOf('\\'), rel.lastIndexOf('/'));
  return idx === -1 ? '' : rel.slice(0, idx);
}

const FILE_COLUMNS = [
  ['Repositório', 18],
  ['Pasta', 30],
  ['Arquivo', 32],
  ['Extensão', 9],
  ['Tamanho (KB)', 12],
  ['Criado em', 17],
  ['Modificado em', 17],
  ['Último usuário', 24],
  ['Fonte do último usuário', 22],
  ['Proprietário (NTFS)', 24],
  ['Salvo por último por (metadados)', 24],
  ['Autor (metadados)', 22],
  ['Último acesso (auditoria)', 24],
  ['Ação (auditoria)', 16],
  ['Data (auditoria)', 17],
  ['Última alteração (auditoria)', 24],
  ['Termos encontrados', 40],
  ['Ocorrências', 11],
  ['Encontrado em', 14],
  ['Situação do conteúdo', 20],
  ['Exclusão', 34],
  ['Caminho completo', 60],
];

function fileRow(r) {
  return [
    r.repositoryName,
    folderOf(r),
    r.name,
    r.extension,
    kb(r.size),
    toDate(r.created),
    toDate(r.modified),
    r.lastUser || '',
    SOURCE_LABELS[r.lastUserSource] || '',
    r.owner || '',
    r.metadata?.lastModifiedBy || '',
    r.metadata?.author || '',
    r.audit?.user || '',
    r.audit?.action || '',
    toDate(r.audit?.time),
    r.audit?.lastWrite ? `${r.audit.lastWrite.user} (${formatDateTime(r.audit.lastWrite.time)})` : '',
    r.terms.join('; '),
    r.occurrences,
    locations(r),
    STATUS_LABELS[r.contentStatus] || r.contentStatus || '',
    deletionText(r.deletion),
    r.path,
  ];
}

/**
 * Aba "Exclusões": todas as exclusões (automáticas e manuais) da análise, na ordem em que
 * aconteceram. label(record) descreve o item (caminho do arquivo ou caixa e assunto).
 */
export function deletionsSheet(deletions, records, columns, label) {
  const byId = new Map(records.map((r) => [r.id, r]));
  return {
    name: 'Exclusões',
    cols: [17, ...columns.map(([, w]) => w), 22, 26, 22, 50],
    header: ['Quando', ...columns.map(([h]) => h), 'Resultado', 'Como', 'Por', 'Detalhe'],
    rows: (function* () {
      for (const d of deletions) {
        const record = byId.get(d.recordId);
        const how = `${d.mode === 'auto' ? 'Automática (na análise)' : 'Manual (relatório)'}${d.method === 'trash' ? ' – para a lixeira' : d.method === 'permanent' ? ' – definitiva' : ''}`;
        yield [toDate(d.at), ...label(record), DELETION_LABELS[d.status] || d.status, how, d.by || '', d.error || ''];
      }
    })(),
  };
}

const MATCH_COLUMNS = [
  ['Repositório', 18],
  ['Arquivo', 30],
  ['Termo', 22],
  ['Lista', 18],
  ['Encontrado em', 13],
  ['Ocorrências', 11],
  ['Valores encontrados', 36],
  ['Exemplo', 70],
  ['Outros exemplos', 70],
  ['Último usuário', 24],
  ['Modificado em', 17],
  ['Caminho completo', 60],
];

function matchRows(r) {
  return r.matches.map((m) => [
    r.repositoryName,
    r.name,
    m.term,
    m.list,
    LOCATION_LABELS[m.location],
    m.count,
    m.values.join('; '),
    sampleText(m.samples[0]),
    m.samples.slice(1).map(sampleText).join('\n'),
    r.lastUser || '',
    toDate(r.modified),
    r.path,
  ]);
}

function scanInfoRows(scan) {
  const s = scan.stats || {};
  const opts = scan.options || {};
  const verifications = [opts.checkName ? `nome (${opts.nameTarget === 'path' ? 'caminho completo' : 'arquivo'})` : null, opts.checkContent ? 'conteúdo' : null]
    .filter(Boolean)
    .join(' e ');
  return [
    ['Análise', scan.name],
    ['Situação', SCAN_STATUS_LABELS[scan.status] || scan.status],
    ['Início', toDate(scan.startedAt)],
    ['Fim', toDate(scan.finishedAt)],
    ['Repositórios', (scan.summary?.repositories || []).map((r) => `${r.name} (${r.path})`).join('; ')],
    ['Listas de referência', (scan.summary?.lists || []).map((l) => `${l.name} (${l.termCount} termos)`).join('; ')],
    ['Verificações', verifications],
    ['Modificados a partir de', opts.modifiedAfter ? toDate(opts.modifiedAfter) : 'todos'],
    ['Arquivos verificados', s.filesSeen ?? 0],
    ['Arquivos com ocorrências', s.filesMatched ?? 0],
    ['Total de ocorrências', s.occurrences ?? 0],
    ['Conteúdos analisados', s.contentAnalyzed ?? 0],
    ['Protegidos por senha', s.contentEncrypted ?? 0],
    ['Grandes demais (só nome)', s.contentSkippedSize ?? 0],
    ['Erros de acesso/leitura', s.errors ?? 0],
    ['Ação', opts.deleteMatches ? 'Analisar e excluir automaticamente' : 'Somente analisar'],
    ...(opts.deleteMatches ? [['Excluídos na análise', s.deleted ?? 0], ['Falhas na exclusão', s.deleteErrors ?? 0]] : []),
  ];
}

/** Grava textos em `out` respeitando o controle de fluxo (espera o "drain" quando necessário). */
export async function writeAll(out, chunks) {
  let buffer = '';
  for (const chunk of chunks) {
    buffer += chunk;
    if (buffer.length >= 64 * 1024) {
      if (!out.write(buffer)) await once(out, 'drain');
      buffer = '';
    }
  }
  if (buffer) out.write(buffer);
}

export async function exportXlsx(scan, records, errors, out, { deletions = [], records: all = records } = {}) {
  const sorted = sortRecords(records);
  const summary = summarize(sorted);

  const resumo = [[{ v: `Relatório CLEAN – ${scan.name}`, s: 'title' }], []];
  for (const [label, value] of scanInfoRows(scan)) resumo.push([{ v: label, s: 'bold' }, value]);
  resumo.push([], [{ v: 'Termos encontrados', s: 'header' }, { v: 'Lista', s: 'header' }, { v: 'Arquivos', s: 'header' }, { v: 'Ocorrências', s: 'header' }]);
  for (const t of summary.byTerm) resumo.push([t.term, t.list, t.files, t.occurrences]);
  resumo.push([], [{ v: 'Último usuário', s: 'header' }, { v: 'Fonte', s: 'header' }, { v: 'Arquivos', s: 'header' }, { v: 'Ocorrências', s: 'header' }]);
  for (const u of summary.byUser) {
    const sources = Object.entries(u.sources).map(([k, n]) => `${SOURCE_LABELS[k]}: ${n}`).join('; ');
    resumo.push([u.user, sources, u.files, u.occurrences]);
  }

  const sheets = [
    { name: 'Resumo', cols: [30, 50, 12, 12], rows: resumo },
    {
      name: 'Arquivos',
      cols: FILE_COLUMNS.map(([, w]) => w),
      header: FILE_COLUMNS.map(([h]) => h),
      rows: (function* () {
        for (const r of sorted) yield fileRow(r);
      })(),
    },
    {
      name: 'Ocorrências',
      cols: MATCH_COLUMNS.map(([, w]) => w),
      header: MATCH_COLUMNS.map(([h]) => h),
      rows: (function* () {
        for (const r of sorted) yield* matchRows(r);
      })(),
    },
  ];
  if (deletions.length) sheets.push(deletionsSheet(deletions, all, [['Arquivo', 70]], (r) => [r?.path || '']));
  if (errors.length) {
    sheets.push({
      name: 'Erros',
      cols: [70, 50, 17],
      header: ['Caminho', 'Erro', 'Quando'],
      rows: (function* () {
        for (const e of errors) yield [e.path, e.message, toDate(e.time)];
      })(),
    });
  }
  await writeXlsx(sheets, out, { title: `Relatório CLEAN – ${scan.name}` });
}

// -- CSV -----------------------------------------------------------------------------------------

/** Data e hora no formato do Excel em português, sem vírgula (reconhecida como data ao abrir). */
function csvDate(date) {
  if (Number.isNaN(date.getTime())) return '';
  const p = (n) => String(n).padStart(2, '0');
  return `${p(date.getDate())}/${p(date.getMonth() + 1)}/${date.getFullYear()} ${p(date.getHours())}:${p(date.getMinutes())}:${p(date.getSeconds())}`;
}

export function csvCell(value) {
  if (value === null || value === undefined) return '';
  if (value instanceof Date) return csvDate(value);
  if (typeof value === 'number') return String(value).replace('.', ',');
  let text = String(value);
  // Evita que o Excel interprete o conteúdo como fórmula (injeção de CSV).
  if (/^[=+\-@\t\r]/.test(text)) text = `'${text}`;
  if (/[";\n\r]/.test(text)) text = `"${text.replace(/"/g, '""')}"`;
  return text;
}

/**
 * CSV com uma linha por arquivo, termo e local (nome ou conteúdo), separador ";" e BOM UTF-8,
 * como o Excel em português espera.
 */
export async function exportCsv(records, out) {
  const header = [...MATCH_COLUMNS.map(([h]) => h), 'Fonte do último usuário', 'Proprietário (NTFS)', 'Salvo por último por (metadados)', 'Último acesso (auditoria)', 'Exclusão'];
  await writeAll(
    out,
    (function* () {
      yield `\uFEFF${header.map(csvCell).join(';')}\r\n`;
      for (const r of sortRecords(records)) {
        const extra = [SOURCE_LABELS[r.lastUserSource] || '', r.owner || '', r.metadata?.lastModifiedBy || '', auditText(r.audit), deletionText(r.deletion)];
        for (const row of matchRows(r)) yield `${[...row, ...extra].map(csvCell).join(';')}\r\n`;
      }
    })(),
  );
}

/** JSON com os dados da análise e todos os resultados. */
export async function exportJson(scan, records, out) {
  await writeAll(
    out,
    (function* () {
      yield `{"scan":${JSON.stringify(scan)},"results":[`;
      let first = true;
      for (const r of records) {
        const { _search, ...rest } = r;
        yield `${first ? '' : ','}\n${JSON.stringify(rest)}`;
        first = false;
      }
      yield '\n]}\n';
    })(),
  );
}

// -- HTML ----------------------------------------------------------------------------------------

export function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

export const REPORT_CSS = `
:root{--ink:#1d2433;--muted:#5b6475;--line:#d9dee7;--accent:#1f4e79;--mark:#fff1a8;--bg:#fff}
*{box-sizing:border-box}body{font:14px/1.45 "Segoe UI",system-ui,sans-serif;color:var(--ink);background:var(--bg);margin:24px}
h1{font-size:22px;margin:0 0 4px}h2{font-size:16px;margin:28px 0 8px;color:var(--accent)}
.muted{color:var(--muted)}table{border-collapse:collapse;width:100%;margin-top:8px}
th,td{border-bottom:1px solid var(--line);padding:6px 8px;text-align:left;vertical-align:top}
th{background:#f3f5f9;font-weight:600;font-size:12px;text-transform:uppercase;letter-spacing:.02em}
td.num,th.num{text-align:right;font-variant-numeric:tabular-nums}
.info td:first-child{width:220px;color:var(--muted)}.path{font-family:Consolas,monospace;font-size:12px;word-break:break-all}
.term{display:inline-block;background:#e8eef7;border-radius:4px;padding:1px 6px;margin:1px 2px;font-size:12px}
.sample{font-size:12px;color:var(--muted);margin:2px 0}.sample mark{background:var(--mark);color:var(--ink)}
@media print{body{margin:0}h2{break-after:avoid}tr{break-inside:avoid}}`;

export async function exportHtml(scan, records, out) {
  const sorted = sortRecords(records);
  const summary = summarize(sorted);
  const info = scanInfoRows(scan)
    .map(([label, value]) => `<tr><td>${escapeHtml(label)}</td><td>${escapeHtml(value instanceof Date ? value.toLocaleString('pt-BR') : value)}</td></tr>`)
    .join('');
  const terms = summary.byTerm
    .map((t) => `<tr><td>${escapeHtml(t.term)}</td><td>${escapeHtml(t.list)}</td><td class="num">${t.files}</td><td class="num">${t.occurrences}</td></tr>`)
    .join('');
  const users = summary.byUser
    .map((u) => `<tr><td>${escapeHtml(u.user)}</td><td class="num">${u.files}</td><td class="num">${u.occurrences}</td></tr>`)
    .join('');
  const sample = (s) =>
    s ? `<div class="sample">${s.where ? `<b>${escapeHtml(s.where)}:</b> ` : ''}${escapeHtml(s.before)}<mark>${escapeHtml(s.match)}</mark>${escapeHtml(s.after)}</div>` : '';
  const row = (r) => {
    const found = r.matches
      .map((m) => `<div><span class="term">${escapeHtml(m.term)}</span> ${escapeHtml(LOCATION_LABELS[m.location])} · ${m.count}×${m.samples.map(sample).join('')}</div>`)
      .join('');
    const user = r.lastUser ? `${escapeHtml(r.lastUser)}<div class="muted">${escapeHtml(SOURCE_LABELS[r.lastUserSource] || '')}</div>` : '<span class="muted">não identificado</span>';
    const deleted = r.deletion ? `<div class="muted">${escapeHtml(deletionText(r.deletion))}</div>` : '';
    return `<tr><td><b>${escapeHtml(r.name)}</b><div class="path">${escapeHtml(r.path)}</div>${deleted}</td><td>${user}</td><td>${escapeHtml(formatDateTime(r.modified))}</td><td>${found}</td></tr>`;
  };
  await writeAll(
    out,
    (function* () {
      yield `<!doctype html><html lang="pt-BR"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>${escapeHtml(`Relatório CLEAN – ${scan.name}`)}</title><style>${REPORT_CSS}</style></head><body>
<h1>Relatório CLEAN</h1><div class="muted">${escapeHtml(scan.name)} · gerado em ${escapeHtml(new Date().toLocaleString('pt-BR'))}</div>
<h2>Resumo</h2><table class="info">${info}</table>
<h2>Termos encontrados</h2><table><thead><tr><th>Termo</th><th>Lista</th><th class="num">Arquivos</th><th class="num">Ocorrências</th></tr></thead><tbody>${terms}</tbody></table>
<h2>Últimos usuários</h2><table><thead><tr><th>Usuário</th><th class="num">Arquivos</th><th class="num">Ocorrências</th></tr></thead><tbody>${users}</tbody></table>
<h2>Arquivos com ocorrências (${sorted.length})</h2><table><thead><tr><th>Arquivo</th><th>Último usuário</th><th>Modificado em</th><th>Informação encontrada</th></tr></thead><tbody>`;
      for (const r of sorted) yield row(r);
      yield '</tbody></table>\n</body></html>';
    })(),
  );
}
