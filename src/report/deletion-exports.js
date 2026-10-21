// Exportação do registro geral de exclusões: Excel (.xlsx), CSV (padrão Excel pt-BR) e HTML para
// impressão. O resumo vem da passada com os filtros (DeletionLog.query); as linhas — uma por tentativa
// de exclusão, das mais recentes para as mais antigas — são lidas do arquivo aos poucos, então
// registros grandes não esgotam a memória.
import { writeXlsx } from './xlsx.js';
import { writeAll, csvCell, escapeHtml, REPORT_CSS, toDate, kb } from './exports.js';
import { KIND_LABELS, PURPOSE_LABELS, ORIGIN_LABELS, NO_POLICY, localDay } from './deletion-log.js';
import { DELETION_LABELS } from './model.js';

/** Linhas da tabela do HTML (o Excel e o CSV trazem todas). */
export const HTML_LIMIT = 10000;

const mb = (bytes) => Math.round(((Number(bytes) || 0) / 1048576) * 100) / 100;
const DAY_RE = /^(\d{4})-(\d{2})-(\d{2})$/;

/** "AAAA-MM-DD" (dia no fuso do servidor) ou ISO -> Date. */
function dateOf(value) {
  if (!value) return null;
  const m = DAY_RE.exec(value);
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : toDate(value);
}

/** "AAAA-MM-DD" -> "dd/mm/aaaa". */
const dayText = (day) => (DAY_RE.test(day || '') ? day.split('-').reverse().join('/') : day || '');

const COLUMNS = [
  ['Quando', 17],
  ['Situação', 24],
  ['Tipo', 16],
  ['Repositório/Conexão', 24],
  ['Política/Agendamento', 28],
  ['Finalidade', 18],
  ['Origem', 20],
  ['Item', 60],
  ['Pasta relativa', 30],
  ['Tamanho (KB)', 12],
  ['Data do critério', 17],
  ['Data de corte', 17],
  ['Como', 22],
  ['Por', 34],
  ['Detalhe', 44],
  ['Análise', 32],
  ['ID da análise', 38],
];

function eventRow(n) {
  return [
    toDate(n.at),
    n.statusLabel,
    n.kindLabel,
    n.repositoryName,
    n.scheduleName || 'Sem política',
    n.purposeLabel,
    n.originLabel,
    n.item,
    n.folder,
    n.size === null ? '' : kb(n.size),
    dateOf(n.criterionDate),
    dateOf(n.cutoff || n.since),
    n.methodLabel,
    n.by,
    n.detail,
    n.scanName,
    n.scanId,
  ];
}

/** Os filtros aplicados, por extenso: [[rótulo, valor]]. options: as opções do resumo (nomes). */
export function filterRows(filters, options = {}) {
  const rows = [];
  if (filters.from || filters.to) {
    rows.push(['Período', filters.from && filters.to ? `de ${dayText(filters.from)} a ${dayText(filters.to)}` : filters.from ? `a partir de ${dayText(filters.from)}` : `até ${dayText(filters.to)}`]);
  }
  if (filters.status) rows.push(['Situação', filters.status.split(',').map((s) => DELETION_LABELS[s] || s).join('; ')]);
  if (filters.mode) rows.push(['Origem', ORIGIN_LABELS[filters.mode] || filters.mode]);
  if (filters.purpose) rows.push(['Finalidade', PURPOSE_LABELS[filters.purpose] || filters.purpose]);
  if (filters.kind) rows.push(['Tipo de local', KIND_LABELS[filters.kind] || filters.kind]);
  if (filters.scheduleId) {
    const name = filters.scheduleId === 'none' ? NO_POLICY : (options.schedules || []).find((s) => s.id === filters.scheduleId)?.name || filters.scheduleId;
    rows.push(['Política/agendamento', name]);
  }
  if (filters.repository) rows.push(['Repositório/conexão', (options.repositories || []).find((r) => r.id === filters.repository)?.name || filters.repository]);
  if (filters.scanId) rows.push(['Análise (ID)', filters.scanId]);
  if (filters.q) rows.push(['Busca', filters.q]);
  return rows.length ? rows : [['Filtros', 'nenhum (todas as exclusões registradas)']];
}

/** Linhas do resumo: [rótulo, valor]. */
function totalsRows(summary) {
  const t = summary.totals;
  return [
    ['Primeira exclusão do recorte', toDate(summary.range.first)],
    ['Última exclusão do recorte', toDate(summary.range.last)],
    ['Itens excluídos', t.uniqueDeleted],
    ['Espaço liberado (MB)', mb(t.bytesDeleted)],
    ...(t.bytesUnknown ? [['Excluídos sem o tamanho registrado', t.bytesUnknown]] : []),
    ['Tentativas de exclusão', t.attempts],
    ['Não encontrados (já excluídos ou movidos)', t.missing],
    ['Não excluídos (alterados depois da análise)', t.changed],
    ['Falhas na exclusão', t.failed],
    ['Itens que continuam no lugar (falha ou alteração na última tentativa)', t.uniquePending],
  ];
}

const pending = (g) => g.failed + g.changed;

/** Tabelas do resumo: [título, cabeçalho, linhas, primeira coluna numérica]. */
function summaryTables(summary) {
  return [
    [
      'Por política ou agendamento',
      ['Política/agendamento', 'Finalidade', 'Excluídos', 'Espaço liberado (MB)', 'Não encontrados', 'Falhas/alterados', 'Tentativas'],
      summary.byPolicy.map((g) => [g.name, g.key === null ? '—' : g.purposeLabel, g.uniqueDeleted, mb(g.bytesDeleted), g.missing, pending(g), g.attempts]),
      2,
    ],
    [
      'Por repositório ou conexão',
      ['Repositório/conexão', 'Tipo', 'Excluídos', 'Espaço liberado (MB)', 'Não encontrados', 'Falhas/alterados', 'Tentativas'],
      summary.byRepository.map((g) => [g.name, g.kindLabel, g.uniqueDeleted, mb(g.bytesDeleted), g.missing, pending(g), g.attempts]),
      2,
    ],
    ['Por situação', ['Situação', 'Tentativas'], summary.byStatus.map((g) => [g.label, g.count]), 1],
    [
      'Por dia',
      ['Dia', 'Excluídos', 'Espaço liberado (MB)', 'Não encontrados', 'Falhas/alterados', 'Tentativas'],
      summary.byDay.map((g) => [dayText(g.day), g.uniqueDeleted, mb(g.bytesDeleted), g.missing, pending(g), g.attempts]),
      1,
    ],
  ];
}

/** Nome do arquivo baixado, com o período: registro-de-exclusoes-2026-10-01-a-2026-10-21.xlsx. */
export function exportFileName(filters, summary, ext) {
  const day = (iso) => (iso && Number.isFinite(Date.parse(iso)) ? localDay(Date.parse(iso)) : '');
  const today = localDay(Date.now());
  const start = filters.from || day(summary.range.first);
  const end = filters.to || (filters.from ? today : day(summary.range.last)) || today;
  const period = !start || start === end ? end : `${start}-a-${end}`;
  return `registro-de-exclusoes-${period}.${ext}`;
}

const TITLE = 'Registro de exclusões – CLEAN';

/**
 * Excel: aba "Resumo" (filtros, período, totais e tabelas) e aba "Exclusões" (uma linha por
 * tentativa). data: { filters, summary, options, events: iterável assíncrono dos eventos }.
 */
export async function exportDeletionsXlsx({ filters, summary, options, events }, out) {
  const resumo = [[{ v: TITLE, s: 'title' }], [], [{ v: 'Gerado em', s: 'bold' }, new Date()]];
  for (const [label, value] of filterRows(filters, options)) resumo.push([{ v: label, s: 'bold' }, value]);
  for (const [label, value] of totalsRows(summary)) resumo.push([{ v: label, s: 'bold' }, value]);
  for (const [title, header, rows] of summaryTables(summary)) {
    resumo.push([], [{ v: title, s: 'bold' }], header.map((v) => ({ v, s: 'header' })));
    resumo.push(...rows);
  }
  const sheets = [
    { name: 'Resumo', cols: [44, 28, 12, 18, 14, 14, 12], rows: resumo },
    {
      name: 'Exclusões',
      cols: COLUMNS.map(([, w]) => w),
      header: COLUMNS.map(([h]) => h),
      rows: (async function* () {
        for await (const n of events) yield eventRow(n);
      })(),
    },
  ];
  await writeXlsx(sheets, out, { title: TITLE });
}

/** CSV com uma linha por tentativa de exclusão (separador ";" e BOM UTF-8, como o Excel em português espera). */
export async function exportDeletionsCsv({ events }, out) {
  await writeAll(
    out,
    (async function* () {
      yield `﻿${COLUMNS.map(([h]) => csvCell(h)).join(';')}\r\n`;
      for await (const n of events) yield `${eventRow(n).map(csvCell).join(';')}\r\n`;
    })(),
  );
}

const STATUS_CLASS = { deleted: 'ok', missing: 'muted', changed: 'warn', failed: 'fail' };
const HTML_CSS = `${REPORT_CSS}
.ok{color:#0b6b0b;font-weight:600}.warn{color:#7a5200;font-weight:600}.fail{color:#9e2323;font-weight:600}.note{background:#fdf1d6;border-radius:6px;padding:8px 10px}`;

/** HTML para impressão: o resumo e a tabela (até HTML_LIMIT linhas, as mais recentes). */
export async function exportDeletionsHtml({ filters, summary, options, events, total }, out) {
  const cell = (value) => escapeHtml(value instanceof Date ? (Number.isNaN(value.getTime()) ? '' : value.toLocaleString('pt-BR')) : value);
  const info = [['Gerado em', new Date()], ...filterRows(filters, options), ...totalsRows(summary)]
    .map(([label, value]) => `<tr><td>${escapeHtml(label)}</td><td>${cell(typeof value === 'number' ? value.toLocaleString('pt-BR') : value)}</td></tr>`)
    .join('');
  const tables = summaryTables(summary)
    .map(
      ([title, header, rows, numFrom]) =>
        `<h2>${escapeHtml(title)}</h2>${
          rows.length
            ? `<table><thead><tr>${header.map((h, i) => `<th${i >= numFrom ? ' class="num"' : ''}>${escapeHtml(h)}</th>`).join('')}</tr></thead><tbody>${rows
                .map((r) => `<tr>${r.map((v) => (typeof v === 'number' ? `<td class="num">${escapeHtml(v.toLocaleString('pt-BR'))}</td>` : `<td>${cell(v)}</td>`)).join('')}</tr>`)
                .join('')}</tbody></table>`
            : '<p class="muted">Nenhuma exclusão.</p>'
        }`,
    )
    .join('\n');
  const shown = Math.min(total, HTML_LIMIT);
  const note =
    total > HTML_LIMIT
      ? `<p class="note">Mostrando as ${HTML_LIMIT.toLocaleString('pt-BR')} exclusões mais recentes de ${total.toLocaleString('pt-BR')}: exporte em Excel para todas.</p>`
      : '';
  const row = (n) => {
    const where = n.folder ? `<div class="muted">${escapeHtml(n.folder)}</div>` : '';
    const size = n.size === null ? '' : `<div class="muted">${escapeHtml(kb(n.size).toLocaleString('pt-BR'))} KB</div>`;
    const policy = n.scheduleName ? `${escapeHtml(n.scheduleName)}<div class="muted">${escapeHtml(n.purposeLabel)}</div>` : `<span class="muted">Sem política</span>${n.purpose ? `<div class="muted">${escapeHtml(n.purposeLabel)}</div>` : ''}`;
    return `<tr><td>${cell(toDate(n.at))}</td><td class="${STATUS_CLASS[n.status] || ''}">${escapeHtml(n.statusLabel)}</td><td>${escapeHtml(n.kindLabel)}</td><td><div class="path">${escapeHtml(n.item)}</div>${where}${size}</td><td>${escapeHtml(n.repositoryName)}</td><td>${policy}</td><td>${escapeHtml(n.originLabel)}<div class="muted">${escapeHtml(n.methodLabel)}</div></td><td>${escapeHtml(n.by)}</td><td>${escapeHtml(n.detail)}</td></tr>`;
  };
  await writeAll(
    out,
    (async function* () {
      yield `<!doctype html><html lang="pt-BR"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>${escapeHtml(TITLE)}</title><style>${HTML_CSS}</style></head><body>
<h1>Registro de exclusões</h1><div class="muted">CLEAN · todas as exclusões registradas (automáticas, manuais e em lote) · gerado em ${escapeHtml(new Date().toLocaleString('pt-BR'))}</div>
<h2>Resumo</h2><table class="info">${info}</table>
${tables}
<h2>Exclusões (${escapeHtml(total.toLocaleString('pt-BR'))})</h2>${note}<table><thead><tr><th>Quando</th><th>Situação</th><th>Tipo</th><th>Item</th><th>Repositório/conexão</th><th>Política/agendamento</th><th>Origem</th><th>Por</th><th>Detalhe</th></tr></thead><tbody>`;
      let count = 0;
      for await (const n of events) {
        if (count++ >= shown) break;
        yield row(n);
      }
      yield `</tbody></table>${total === 0 ? '<p class="muted">Nenhuma exclusão registrada com estes filtros.</p>' : ''}\n</body></html>`;
    })(),
  );
}
