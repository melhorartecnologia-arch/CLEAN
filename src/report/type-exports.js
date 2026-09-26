// Exportação dos relatórios das buscas por tipo de arquivo: Excel, CSV e HTML. Cada linha é um
// arquivo encontrado (sem termos), com o tipo, a extensão e como ele foi encontrado (pela extensão
// ou pelo tipo real no conteúdo).
import { writeXlsx } from './xlsx.js';
import { writeAll, csvCell, escapeHtml, REPORT_CSS, toDate, kb, deletionsSheet, deletionInfoRows } from './exports.js';
import { summarizeTypes, deletionText, formatDateTime, SOURCE_LABELS, SCAN_STATUS_LABELS } from './model.js';
import { CATEGORIES, categoryLabel, describeFileTypes } from '../types/catalog.js';

const mb = (bytes) => Math.round(((Number(bytes) || 0) / 1048576) * 100) / 100;

/** Pasta do arquivo (no OneDrive/SharePoint, com a conta ou o site e a biblioteca). */
function folderOf(record) {
  const rel = record.relativePath || '';
  const idx = Math.max(rel.lastIndexOf('\\'), rel.lastIndexOf('/'));
  const dir = idx === -1 ? '' : rel.slice(0, idx);
  const c = record.cloud;
  if (!c) return dir;
  const start = `${c.accountName || c.account} › ${c.library}`;
  return dir ? `${start} › ${dir}` : start;
}

const sortItems = (records) => records.slice().sort((a, b) => a.repositoryName.localeCompare(b.repositoryName, 'pt-BR') || a.path.localeCompare(b.path, 'pt-BR'));
const foundBy = (r) => (r.typeMatch?.by === 'content' ? 'Conteúdo (tipo real)' : 'Extensão');

export function typeInfoRows(scan) {
  const s = scan.stats || {};
  const o = scan.options || {};
  const t = scan.fileTypes || {};
  return [
    ['Análise', scan.name],
    ...(scan.scheduleId ? [['Agendamento', scan.scheduleName || '']] : []),
    ['Situação', SCAN_STATUS_LABELS[scan.status] || scan.status],
    ['Início', toDate(scan.startedAt)],
    ['Fim', toDate(scan.finishedAt)],
    ['Repositórios', (scan.summary?.repositories || []).map((r) => `${r.name} (${r.path})`).join('; ')],
    ['Tipos procurados', describeFileTypes(t)],
    ['Categorias', (t.categories || []).map((c) => CATEGORIES[c]?.label || c).join('; ') || 'nenhuma'],
    ['Outras extensões', (t.extensions || []).join('; ') || 'nenhuma'],
    ['Tamanho mínimo', t.minSizeMB ? `${String(t.minSizeMB).replace('.', ',')} MB` : 'todos os tamanhos'],
    ['Tipo real pelo conteúdo', t.checkContent ? 'conferido (pastas do Windows)' : 'não conferido (só a extensão)'],
    ['Alterados a partir de', o.modifiedAfter ? toDate(o.modifiedAfter) : 'todos'],
    ['Arquivos verificados', s.filesSeen ?? 0],
    ['Arquivos encontrados', s.filesMatched ?? 0],
    ['Tamanho dos encontrados (MB)', mb(s.bytesFound)],
    ['Encontrados pelo conteúdo (outra extensão)', s.typesByContent ?? 0],
    ['Abaixo do tamanho mínimo', s.filesSkippedBySize ?? 0],
    ['Erros de acesso/leitura', s.errors ?? 0],
    ...deletionInfoRows(o, s, {
      blocked: scan.deletionBlocked,
      revoked: scan.deletionRevoked,
      labels: {
        action: `Procurar e excluir automaticamente${t.maxDeletions ? ` (até ${t.maxDeletions} por execução)` : ''}`,
        listOnly: 'Somente procurar (revisão no relatório)',
        extra: [
          ['Não excluídos (limite da execução)', s.deleteSkipped ?? 0],
          ['Em locais protegidos (não excluídos)', s.deleteProtected ?? 0],
        ],
      },
    }),
  ];
}

/** Tabelas do resumo: [título, rótulo da coluna, grupos { label, count, bytes }]. */
function summaryTables(records) {
  const s = summarizeTypes(records);
  const rows = (groups, label) => groups.map((g) => ({ label: label(g), count: g.count, bytes: g.bytes }));
  return [
    ['Tipos encontrados', 'Tipo', rows(s.byType, (g) => g.label || categoryLabel(g.key))],
    ['Extensões', 'Extensão', rows(s.byExtension, (g) => g.key || '(sem extensão)')],
    ['Repositórios', 'Repositório', rows(s.byRepository, (g) => g.name)],
    ['Últimos usuários', 'Usuário', rows(s.byUser, (g) => g.key || '(não identificado)')],
  ];
}

const COLUMNS = [
  ['Repositório', 18],
  ['Pasta', 30],
  ['Arquivo', 32],
  ['Extensão', 9],
  ['Tipo', 22],
  ['Encontrado por', 18],
  ['Formato real', 24],
  ['Tamanho (KB)', 12],
  ['Criado em', 17],
  ['Modificado em', 17],
  ['Último acesso', 17],
  ['Último usuário', 24],
  ['Fonte do último usuário', 22],
  ['Proprietário (NTFS) / dono do OneDrive', 26],
  ['Exclusão', 34],
  ['Caminho completo', 60],
];

function row(r) {
  return [
    r.repositoryName,
    folderOf(r),
    r.name,
    r.extension,
    categoryLabel(r.typeMatch?.category),
    foundBy(r),
    r.typeMatch?.format || '',
    kb(r.size),
    toDate(r.created),
    toDate(r.modified),
    toDate(r.accessed),
    r.lastUser || '',
    SOURCE_LABELS[r.lastUserSource] || '',
    r.owner || '',
    deletionText(r.deletion),
    r.path,
  ];
}

export async function exportTypesXlsx(scan, records, errors, out, { deletions = [], records: all = records } = {}) {
  const sorted = sortItems(records);
  const resumo = [[{ v: `Relatório CLEAN – ${scan.name}`, s: 'title' }], []];
  for (const [label, value] of typeInfoRows(scan)) resumo.push([{ v: label, s: 'bold' }, value]);
  for (const [, column, groups] of summaryTables(sorted)) {
    resumo.push([], [column, 'Arquivos', 'Tamanho (MB)'].map((v) => ({ v, s: 'header' })));
    for (const g of groups) resumo.push([g.label, g.count, mb(g.bytes)]);
  }
  const sheets = [
    { name: 'Resumo', cols: [36, 50, 14], rows: resumo },
    {
      name: 'Arquivos encontrados',
      cols: COLUMNS.map(([, w]) => w),
      header: COLUMNS.map(([h]) => h),
      rows: (function* () {
        for (const r of sorted) yield row(r);
      })(),
    },
  ];
  if (deletions.length) sheets.push(deletionsSheet(deletions, all, [['Arquivo', 70]], (r, d) => [r?.path || d.item || '']));
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

/** CSV com uma linha por arquivo encontrado (separador ";" e BOM UTF-8, como o Excel em português espera). */
export async function exportTypesCsv(records, out) {
  await writeAll(
    out,
    (function* () {
      yield `﻿${COLUMNS.map(([h]) => csvCell(h)).join(';')}\r\n`;
      for (const r of sortItems(records)) yield `${row(r).map(csvCell).join(';')}\r\n`;
    })(),
  );
}

export async function exportTypesHtml(scan, records, out) {
  const sorted = sortItems(records);
  const cell = (value) => escapeHtml(value instanceof Date ? value.toLocaleString('pt-BR') : value);
  const info = typeInfoRows(scan)
    .map(([label, value]) => `<tr><td>${escapeHtml(label)}</td><td>${cell(value)}</td></tr>`)
    .join('');
  const tables = summaryTables(sorted)
    .map(
      ([title, column, groups]) =>
        `<h2>${escapeHtml(title)}</h2><table><thead><tr><th>${escapeHtml(column)}</th><th class="num">Arquivos</th><th class="num">Tamanho (MB)</th></tr></thead><tbody>${groups
          .map((g) => `<tr><td>${escapeHtml(g.label)}</td><td class="num">${g.count}</td><td class="num">${mb(g.bytes).toLocaleString('pt-BR')}</td></tr>`)
          .join('')}</tbody></table>`,
    )
    .join('\n');
  const line = (r) => {
    const deleted = r.deletion ? `<div class="muted">${escapeHtml(deletionText(r.deletion))}</div>` : '';
    const how = r.typeMatch?.by === 'content' ? `<div class="muted">tipo real: ${escapeHtml(r.typeMatch.format || '')}</div>` : '';
    return `<tr><td><b>${escapeHtml(r.name)}</b><div class="path">${escapeHtml(r.path)}</div>${deleted}</td><td>${escapeHtml(categoryLabel(r.typeMatch?.category))}${how}</td><td>${r.lastUser ? escapeHtml(r.lastUser) : '<span class="muted">não identificado</span>'}</td><td>${escapeHtml(formatDateTime(r.modified))}</td><td class="num">${kb(r.size || 0).toLocaleString('pt-BR')}</td></tr>`;
  };
  await writeAll(
    out,
    (function* () {
      yield `<!doctype html><html lang="pt-BR"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>${escapeHtml(`Relatório CLEAN – ${scan.name}`)}</title><style>${REPORT_CSS}</style></head><body>
<h1>Relatório CLEAN – busca por tipo</h1><div class="muted">${escapeHtml(scan.name)} · gerado em ${escapeHtml(new Date().toLocaleString('pt-BR'))}</div>
<h2>Resumo</h2><table class="info">${info}</table>
${tables}
<h2>Arquivos encontrados (${sorted.length})</h2><table><thead><tr><th>Arquivo</th><th>Tipo</th><th>Último usuário</th><th>Modificado em</th><th class="num">Tamanho (KB)</th></tr></thead><tbody>`;
      for (const r of sorted) yield line(r);
      yield '</tbody></table>\n</body></html>';
    })(),
  );
}

/** Exportações de uma busca por tipo (mesmas assinaturas das análises comuns). */
export const TYPE_EXPORTS = { xlsx: exportTypesXlsx, csv: exportTypesCsv, html: exportTypesHtml };
