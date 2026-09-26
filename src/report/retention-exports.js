// Exportação dos relatórios das políticas de retenção (arquivos e e-mail): Excel, CSV e HTML.
// Os itens expirados não têm termos: cada linha é um arquivo ou uma mensagem, com a data do
// critério da política e a idade (a análise comum tem uma linha por termo encontrado).
import { writeXlsx } from './xlsx.js';
import { writeAll, csvCell, escapeHtml, REPORT_CSS, toDate, kb, deletionsSheet, deletionInfoRows } from './exports.js';
import { sourceText } from './mail-exports.js';
import { summarizeRetention, deletionText, formatDateTime, SOURCE_LABELS, SCAN_STATUS_LABELS, MAIL_DELETION_LABELS } from './model.js';
import { FILE_CRITERIA, MAIL_CRITERIA, amountText, describeRetention, ageBucket } from '../retention/policy.js';

const mb = (bytes) => Math.round(((Number(bytes) || 0) / 1048576) * 100) / 100;
const isMail = (scan) => scan.kind === 'mail';

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

const ageOf = (r) => (r.retention ? r.retention.ageDays : null);
const bucketOf = (r) => (r.retention ? ageBucket(r.retention.ageDays).label : '');

function sortItems(scan, records) {
  if (isMail(scan)) {
    return records.slice().sort((a, b) => a.mailbox.localeCompare(b.mailbox, 'pt-BR') || String(a.date || '').localeCompare(String(b.date || '')));
  }
  return records.slice().sort((a, b) => a.repositoryName.localeCompare(b.repositoryName, 'pt-BR') || a.path.localeCompare(b.path, 'pt-BR'));
}

/** Regra da política gravada na análise (com a data de corte do momento em que ela foi criada). */
function policyRows(scan) {
  const r = scan.retention;
  const mail = isMail(scan);
  const criterion = (mail ? MAIL_CRITERIA : FILE_CRITERIA)[r.criterion];
  return [
    ['Regra', describeRetention(r, mail ? 'mail' : 'files')],
    ['Critério de data', criterion?.label || r.criterion],
    ['Idade máxima', amountText(r)],
    ['Data de corte (expiram os anteriores)', toDate(r.cutoff)],
    mail
      ? ['Lixeira / Lixo eletrônico', `${r.includeTrash ? 'inclui' : 'ignora'} a lixeira; ${r.includeJunk ? 'inclui' : 'ignora'} o lixo eletrônico`]
      : ['Nomes de arquivo', r.patterns?.length ? r.patterns.join('; ') : 'todos'],
    ['Limite de exclusões por execução', r.maxDeletions ? r.maxDeletions : 'sem limite'],
  ];
}

const CLOUD_TYPES = new Set(['onedrive', 'sharepoint']);

/**
 * Forma da exclusão por extenso. Nas pastas do Windows ela é sempre definitiva: "para a lixeira"
 * vale só para o e-mail, o OneDrive e o SharePoint.
 */
export function retentionModeText(scan) {
  if (scan.retention?.deleteMode !== 'trash') return 'definitivamente';
  if (isMail(scan)) return 'para a lixeira';
  const repos = scan.summary?.repositories || [];
  const cloud = repos.some((r) => CLOUD_TYPES.has(r.type));
  const local = repos.some((r) => !CLOUD_TYPES.has(r.type));
  if (cloud && local) return 'para a lixeira no OneDrive e no SharePoint; definitivamente nas pastas do Windows';
  return cloud ? 'para a lixeira' : 'definitivamente, nas pastas do Windows';
}

export function retentionInfoRows(scan) {
  const s = scan.stats || {};
  const o = scan.options || {};
  const mail = isMail(scan);
  const counts = mail
    ? [
        ['Caixas analisadas', s.mailboxesDone ?? 0],
        ['Caixas ignoradas (sem e-mail)', s.mailboxesSkipped ?? 0],
        ['Mensagens expiradas', s.messagesMatched ?? 0],
        ['Tamanho das expiradas (MB)', mb(s.bytesExpired)],
        ['Sem data de recebimento válida (mantidas)', s.retentionUnknown ?? 0],
        ['Erros', s.errors ?? 0],
      ]
    : [
        ['Arquivos verificados', s.filesSeen ?? 0],
        ['Arquivos expirados', s.filesMatched ?? 0],
        ['Tamanho dos expirados (MB)', mb(s.bytesExpired)],
        ['Sem a data do critério (mantidos)', s.retentionUnknown ?? 0],
        ['Erros de acesso/leitura', s.errors ?? 0],
      ];
  const retention = {
    mode: retentionModeText(scan),
    extra: mail ? [['Já estavam na lixeira (não movidas de novo)', s.alreadyInTrash ?? 0]] : [['Em locais protegidos (não excluídos)', s.deleteProtected ?? 0]],
  };
  const stops = { blocked: scan.deletionBlocked, revoked: scan.deletionRevoked };
  const deletion = mail
    ? { noun: 'Excluídas', gone: 'Já não existiam', changed: 'Alteradas depois da análise (mantidas)', skipped: 'Não excluídas (limite da execução)', retention, ...stops }
    : { changed: 'Alterados ou não mais expirados (mantidos)', retention, ...stops };
  return [
    ['Análise', scan.name],
    ...(scan.scheduleId ? [['Política de retenção', scan.scheduleName || '']] : []),
    ['Situação', SCAN_STATUS_LABELS[scan.status] || scan.status],
    ['Início', toDate(scan.startedAt)],
    ['Fim', toDate(scan.finishedAt)],
    mail
      ? ['Conexões de e-mail', (scan.summary?.sources || []).map(sourceText).join('; ')]
      : ['Repositórios', (scan.summary?.repositories || []).map((r) => `${r.name} (${r.path})`).join('; ')],
    ...policyRows(scan),
    ...counts,
    ...deletionInfoRows(o, s, deletion),
  ];
}

/** Tabelas do resumo: [título, rótulo da coluna, grupos { label, count, bytes }]. */
function summaryTables(scan, records) {
  const s = summarizeRetention(records, isMail(scan) ? 'mail' : 'files');
  const rows = (groups, label) => groups.map((g) => ({ label: label(g), count: g.count, bytes: g.bytes }));
  if (isMail(scan)) {
    return [
      ['Idade das mensagens expiradas', 'Faixa de idade', rows(s.byAge, (g) => g.label)],
      ['Caixas', 'Caixa', rows(s.byMailbox, (g) => (g.name ? `${g.name} <${g.key}>` : g.key))],
      ['Pastas', 'Pasta', rows(s.byFolder, (g) => g.key || '(sem pasta)')],
    ];
  }
  return [
    ['Idade dos arquivos expirados', 'Faixa de idade', rows(s.byAge, (g) => g.label)],
    ['Repositórios', 'Repositório', rows(s.byRepository, (g) => g.name)],
    ['Extensões', 'Extensão', rows(s.byExtension, (g) => g.key || '(sem extensão)')],
    ['Últimos usuários', 'Usuário', rows(s.byUser, (g) => g.key || '(não identificado)')],
  ];
}

const FILE_COLUMNS = [
  ['Repositório', 18],
  ['Pasta', 30],
  ['Arquivo', 32],
  ['Extensão', 9],
  ['Tamanho (KB)', 12],
  ['Data considerada', 17],
  ['Idade (dias)', 11],
  ['Faixa de idade', 14],
  ['Criado em', 17],
  ['Modificado em', 17],
  ['Último acesso', 17],
  ['Último usuário', 24],
  ['Fonte do último usuário', 22],
  ['Proprietário (NTFS) / dono do OneDrive', 26],
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
    toDate(r.retention?.date),
    ageOf(r),
    bucketOf(r),
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

const MESSAGE_COLUMNS = [
  ['Conexão', 16],
  ['Caixa', 28],
  ['Pasta', 22],
  ['Recebida em', 17],
  ['Idade (dias)', 11],
  ['Faixa de idade', 14],
  ['Remetente', 32],
  ['Assunto', 44],
  ['Tamanho (KB)', 12],
  ['Exclusão', 34],
  ['Message-ID', 40],
  ['Link (Outlook na Web)', 30],
];

function messageRow(r) {
  return [
    r.sourceName,
    r.mailbox,
    r.folder,
    toDate(r.retention?.date || r.date),
    ageOf(r),
    bucketOf(r),
    r.from,
    r.subject,
    kb(r.size || 0),
    deletionText(r.deletion, 'mail'),
    r.internetMessageId || '',
    r.webLink || '',
  ];
}

const layout = (scan) =>
  isMail(scan) ? { sheet: 'Mensagens expiradas', columns: MESSAGE_COLUMNS, row: messageRow } : { sheet: 'Arquivos expirados', columns: FILE_COLUMNS, row: fileRow };

export async function exportRetentionXlsx(scan, records, errors, out, { deletions = [], records: all = records } = {}) {
  const sorted = sortItems(scan, records);
  const { sheet, columns, row } = layout(scan);
  const mail = isMail(scan);
  const resumo = [[{ v: `Relatório CLEAN – ${scan.name}`, s: 'title' }], []];
  for (const [label, value] of retentionInfoRows(scan)) resumo.push([{ v: label, s: 'bold' }, value]);
  for (const [, column, groups] of summaryTables(scan, sorted)) {
    resumo.push([], [column, mail ? 'Mensagens' : 'Arquivos', 'Tamanho (MB)'].map((v) => ({ v, s: 'header' })));
    for (const g of groups) resumo.push([g.label, g.count, mb(g.bytes)]);
  }
  const sheets = [
    { name: 'Resumo', cols: [36, 50, 14], rows: resumo },
    {
      name: sheet,
      cols: columns.map(([, w]) => w),
      header: columns.map(([h]) => h),
      rows: (function* () {
        for (const r of sorted) yield row(r);
      })(),
    },
  ];
  if (deletions.length) {
    sheets.push(
      mail
        ? deletionsSheet(
            deletions,
            all,
            [['Caixa', 28], ['Pasta', 20], ['Assunto', 40]],
            (r, d) => (r ? [r.mailbox || '', r.folder || '', r.subject || ''] : [d.item || '', '', '']),
            MAIL_DELETION_LABELS,
          )
        : deletionsSheet(deletions, all, [['Arquivo', 70]], (r, d) => [r?.path || d.item || '']),
    );
  }
  if (errors.length) {
    sheets.push({
      name: 'Erros',
      cols: [70, 50, 17],
      header: [mail ? 'Local' : 'Caminho', 'Erro', 'Quando'],
      rows: (function* () {
        for (const e of errors) yield [e.path, e.message, toDate(e.time)];
      })(),
    });
  }
  await writeXlsx(sheets, out, { title: `Relatório CLEAN – ${scan.name}` });
}

/** CSV com uma linha por item expirado (separador ";" e BOM UTF-8, como o Excel em português espera). */
export async function exportRetentionCsv(records, out, scan) {
  const { columns, row } = layout(scan);
  await writeAll(
    out,
    (function* () {
      yield `﻿${columns.map(([h]) => csvCell(h)).join(';')}\r\n`;
      for (const r of sortItems(scan, records)) yield `${row(r).map(csvCell).join(';')}\r\n`;
    })(),
  );
}

export async function exportRetentionHtml(scan, records, out) {
  const sorted = sortItems(scan, records);
  const mail = isMail(scan);
  const cell = (value) => escapeHtml(value instanceof Date ? value.toLocaleString('pt-BR') : value);
  const info = retentionInfoRows(scan)
    .map(([label, value]) => `<tr><td>${escapeHtml(label)}</td><td>${cell(value)}</td></tr>`)
    .join('');
  const noun = mail ? 'Mensagens' : 'Arquivos';
  const tables = summaryTables(scan, sorted)
    .map(
      ([title, column, groups]) =>
        `<h2>${escapeHtml(title)}</h2><table><thead><tr><th>${escapeHtml(column)}</th><th class="num">${noun}</th><th class="num">Tamanho (MB)</th></tr></thead><tbody>${groups
          .map((g) => `<tr><td>${escapeHtml(g.label)}</td><td class="num">${g.count}</td><td class="num">${mb(g.bytes).toLocaleString('pt-BR')}</td></tr>`)
          .join('')}</tbody></table>`,
    )
    .join('\n');
  const deleted = (r) => (r.deletion ? `<div class="muted">${escapeHtml(deletionText(r.deletion, mail ? 'mail' : 'files'))}</div>` : '');
  const size = (r) => kb(r.size || 0).toLocaleString('pt-BR');
  const row = mail
    ? (r) =>
        `<tr><td><b>${escapeHtml(r.subject || '(sem assunto)')}</b><div class="muted">${escapeHtml(r.mailbox)} › ${escapeHtml(r.folder)}</div>${deleted(r)}</td><td>${escapeHtml(r.from)}</td><td>${escapeHtml(formatDateTime(r.retention?.date || r.date))}</td><td class="num">${ageOf(r) ?? ''}</td><td class="num">${size(r)}</td></tr>`
    : (r) =>
        `<tr><td><b>${escapeHtml(r.name)}</b><div class="path">${escapeHtml(r.path)}</div>${deleted(r)}</td><td>${r.lastUser ? escapeHtml(r.lastUser) : '<span class="muted">não identificado</span>'}</td><td>${escapeHtml(formatDateTime(r.retention?.date))}</td><td class="num">${ageOf(r) ?? ''}</td><td class="num">${size(r)}</td></tr>`;
  const head = mail
    ? '<th>Mensagem</th><th>Remetente</th><th>Recebida em</th><th class="num">Idade (dias)</th><th class="num">Tamanho (KB)</th>'
    : '<th>Arquivo</th><th>Último usuário</th><th>Data considerada</th><th class="num">Idade (dias)</th><th class="num">Tamanho (KB)</th>';
  await writeAll(
    out,
    (function* () {
      yield `<!doctype html><html lang="pt-BR"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>${escapeHtml(`Relatório CLEAN – ${scan.name}`)}</title><style>${REPORT_CSS}</style></head><body>
<h1>Relatório CLEAN – retenção</h1><div class="muted">${escapeHtml(scan.name)} · gerado em ${escapeHtml(new Date().toLocaleString('pt-BR'))}</div>
<h2>Resumo</h2><table class="info">${info}</table>
${tables}
<h2>${mail ? 'Mensagens expiradas' : 'Arquivos expirados'} (${sorted.length})</h2><table><thead><tr>${head}</tr></thead><tbody>`;
      for (const r of sorted) yield row(r);
      yield '</tbody></table>\n</body></html>';
    })(),
  );
}

/** Exportações de uma análise de retenção (mesmas assinaturas das análises comuns). */
export const RETENTION_EXPORTS = { xlsx: exportRetentionXlsx, csv: exportRetentionCsv, html: exportRetentionHtml };
