// Exportação dos relatórios de análises do Microsoft Teams: Excel, CSV (padrão Excel pt-BR) e HTML.
// (O JSON é o comum das exportações.)
import { writeXlsx } from './xlsx.js';
import { writeAll, csvCell, escapeHtml, REPORT_CSS, toDate, kb, deletionsSheet, deletionInfoRows } from './exports.js';
import { summarizeTeams, sampleText, formatDateTime, deletionText, STATUS_LABELS, LOCATION_LABELS, SCAN_STATUS_LABELS, MAIL_DELETION_LABELS } from './model.js';

const SCOPE = { channel: 'Canal', chat: 'Chat' };

function sortMessages(records) {
  return records.slice().sort((a, b) => a.folder.localeCompare(b.folder, 'pt-BR') || String(b.date || '').localeCompare(String(a.date || '')));
}

const locations = (r) => [...new Set(r.matches.map((m) => LOCATION_LABELS[m.location] || m.location))].join(', ');
const attachmentNames = (r) => (r.attachments || []).map((a) => a.name).join('; ');

const MESSAGE_COLUMNS = [
  ['Conexão', 16],
  ['Âmbito', 10],
  ['Equipe', 26],
  ['Canal / Chat', 30],
  ['Autor', 28],
  ['Data', 17],
  ['Assunto', 32],
  ['Anexos', 32],
  ['Tamanho (KB)', 12],
  ['Termos encontrados', 32],
  ['Ocorrências', 11],
  ['Encontrado em', 22],
  ['Situação', 16],
  ['Exclusão', 34],
  ['Link', 30],
];

function messageRow(r) {
  return [
    r.sourceName,
    SCOPE[r.scopeKind] || r.scopeKind,
    r.team || '',
    r.scopeKind === 'channel' ? r.channel : r.folder.replace(/^Chat:\s*/, ''),
    r.from,
    toDate(r.date),
    r.subject,
    attachmentNames(r),
    kb(r.size || 0),
    r.terms.join('; '),
    r.occurrences,
    locations(r),
    STATUS_LABELS[r.contentStatus] || r.contentStatus || '',
    deletionText(r.deletion, 'mail'),
    r.webUrl || '',
  ];
}

const MATCH_COLUMNS = [
  ['Âmbito', 10],
  ['Equipe', 24],
  ['Canal / Chat', 28],
  ['Autor', 26],
  ['Data', 17],
  ['Assunto', 30],
  ['Termo', 22],
  ['Lista', 18],
  ['Encontrado em', 18],
  ['Ocorrências', 11],
  ['Valores encontrados', 32],
  ['Exemplo', 70],
  ['Outros exemplos', 70],
];

function matchRows(r) {
  const channelOrChat = r.scopeKind === 'channel' ? r.channel : r.folder.replace(/^Chat:\s*/, '');
  return r.matches.map((m) => [
    SCOPE[r.scopeKind] || r.scopeKind,
    r.team || '',
    channelOrChat,
    r.from,
    toDate(r.date),
    r.subject,
    m.term,
    m.list,
    LOCATION_LABELS[m.location] || m.location,
    m.count,
    m.values.join('; '),
    sampleText(m.samples[0]),
    m.samples.slice(1).map(sampleText).join('\n'),
  ]);
}

function scanInfoRows(scan) {
  const s = scan.stats || {};
  const o = scan.options || {};
  const t = scan.teams || {};
  const checks = [o.checkSubject && 'assunto', o.checkBody && 'corpo', o.checkAttachmentNames && 'nomes dos anexos', o.checkAttachments && 'conteúdo dos anexos'].filter(Boolean);
  const scope = t.scanChannels && t.scanChats ? 'canais e chats' : t.scanChannels ? 'canais' : 'chats';
  return [
    ['Análise', scan.name],
    ...(scan.scheduleId ? [['Agendamento', scan.scheduleName || '']] : []),
    ['Situação', SCAN_STATUS_LABELS[scan.status] || scan.status],
    ['Início', toDate(scan.startedAt)],
    ['Fim', toDate(scan.finishedAt)],
    ['Conexões Microsoft 365', (scan.summary?.sources || []).map((x) => x.name).join('; ')],
    ['Listas de referência', (scan.summary?.lists || []).map((l) => `${l.name} (${l.termCount} termos)`).join('; ')],
    ['Âmbito', `${scope}${t.scope === 'list' ? ' (lista selecionada)' : ' (todo o locatário)'}`],
    ['Verificações', checks.join(', ')],
    ['A partir de', o.receivedAfter ? toDate(o.receivedAfter) : 'todas as mensagens'],
    ['Conversas analisadas', s.conversationsDone ?? 0],
    ['Mensagens verificadas', s.messagesSeen ?? 0],
    ['Mensagens com ocorrências', s.messagesMatched ?? 0],
    ['Total de ocorrências', s.occurrences ?? 0],
    ['Anexos lidos', s.attachmentsAnalyzed ?? 0],
    ['Erros', s.errors ?? 0],
    ...deletionInfoRows(o, s, { noun: 'Excluídas (softDelete)', gone: 'Já não existiam', changed: 'Mantidas', blocked: scan.deletionBlocked, revoked: scan.deletionRevoked }),
    ...(o.deleteMatches && s.deleteSkipped ? [['Chats não excluídos (sem suporte do Graph)', s.deleteSkipped]] : []),
  ];
}

export async function exportTeamsXlsx(scan, records, errors, out, { deletions = [], records: all = records } = {}) {
  const sorted = sortMessages(records);
  const summary = summarizeTeams(sorted);
  const header = (...labels) => labels.map((v) => ({ v, s: 'header' }));
  const resumo = [[{ v: `Relatório CLEAN – Teams – ${scan.name}`, s: 'title' }], []];
  for (const [label, value] of scanInfoRows(scan)) resumo.push([{ v: label, s: 'bold' }, value]);
  resumo.push([], header('Termos encontrados', 'Lista', 'Mensagens', 'Ocorrências'));
  for (const t of summary.byTerm) resumo.push([t.term, t.list, t.messages, t.occurrences]);
  resumo.push([], header('Equipe', 'Mensagens', 'Ocorrências'));
  for (const g of summary.byTeam) resumo.push([g.team, g.messages, g.occurrences]);
  resumo.push([], header('Autor', 'Mensagens', 'Ocorrências'));
  for (const g of summary.bySender) resumo.push([g.label, g.messages, g.occurrences]);

  const sheets = [
    { name: 'Resumo', cols: [34, 40, 12, 12], rows: resumo },
    {
      name: 'Mensagens',
      cols: MESSAGE_COLUMNS.map(([, w]) => w),
      header: MESSAGE_COLUMNS.map(([h]) => h),
      rows: (function* () {
        for (const r of sorted) yield messageRow(r);
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
  if (deletions.length) {
    sheets.push(
      deletionsSheet(
        deletions,
        all,
        [['Âmbito', 10], ['Equipe/Canal ou Chat', 40], ['Assunto', 32]],
        (r) => (r ? [SCOPE[r.scopeKind] || '', r.folder || '', r.subject || ''] : ['', '', '']),
        MAIL_DELETION_LABELS,
      ),
    );
  }
  if (errors.length) {
    sheets.push({
      name: 'Erros',
      cols: [60, 60, 17],
      header: ['Local', 'Erro', 'Quando'],
      rows: (function* () {
        for (const e of errors) yield [e.path, e.message, toDate(e.time)];
      })(),
    });
  }
  await writeXlsx(sheets, out, { title: `Relatório CLEAN – Teams – ${scan.name}` });
}

export async function exportTeamsCsv(records, out) {
  const header = [...MATCH_COLUMNS.map(([h]) => h), 'Anexos', 'Conexão', 'Exclusão'];
  await writeAll(
    out,
    (function* () {
      yield `﻿${header.map(csvCell).join(';')}\r\n`;
      for (const r of sortMessages(records)) {
        const extra = [attachmentNames(r), r.sourceName, deletionText(r.deletion, 'mail')];
        for (const row of matchRows(r)) yield `${[...row, ...extra].map(csvCell).join(';')}\r\n`;
      }
    })(),
  );
}

export async function exportTeamsHtml(scan, records, out) {
  const sorted = sortMessages(records);
  const summary = summarizeTeams(sorted);
  const info = scanInfoRows(scan)
    .map(([label, value]) => `<tr><td>${escapeHtml(label)}</td><td>${escapeHtml(value instanceof Date ? value.toLocaleString('pt-BR') : value)}</td></tr>`)
    .join('');
  const terms = summary.byTerm
    .map((t) => `<tr><td>${escapeHtml(t.term)}</td><td>${escapeHtml(t.list)}</td><td class="num">${t.messages}</td><td class="num">${t.occurrences}</td></tr>`)
    .join('');
  const sample = (s) =>
    s ? `<div class="sample">${s.where ? `<b>${escapeHtml(s.where)}:</b> ` : ''}${escapeHtml(s.before)}<mark>${escapeHtml(s.match)}</mark>${escapeHtml(s.after)}</div>` : '';
  const row = (r) => {
    const found = r.matches
      .map((m) => `<div><span class="term">${escapeHtml(m.term)}</span> ${escapeHtml(LOCATION_LABELS[m.location] || m.location)} · ${m.count}×${m.samples.map(sample).join('')}</div>`)
      .join('');
    const files = attachmentNames(r);
    const deleted = r.deletion ? `<div class="muted">${escapeHtml(deletionText(r.deletion, 'mail'))}</div>` : '';
    return `<tr><td><b>${escapeHtml(r.subject || '(sem assunto)')}</b><div class="muted">${escapeHtml(r.folder)}</div>${files ? `<div class="path">Anexos: ${escapeHtml(files)}</div>` : ''}${deleted}</td><td>${escapeHtml(r.from)}</td><td>${escapeHtml(formatDateTime(r.date))}</td><td>${found}</td></tr>`;
  };
  await writeAll(
    out,
    (function* () {
      yield `<!doctype html><html lang="pt-BR"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>${escapeHtml(`Relatório CLEAN – Teams – ${scan.name}`)}</title><style>${REPORT_CSS}</style></head><body>
<h1>Relatório CLEAN – Microsoft Teams</h1><div class="muted">${escapeHtml(scan.name)} · gerado em ${escapeHtml(new Date().toLocaleString('pt-BR'))}</div>
<h2>Resumo</h2><table class="info">${info}</table>
<h2>Termos encontrados</h2><table><thead><tr><th>Termo</th><th>Lista</th><th class="num">Mensagens</th><th class="num">Ocorrências</th></tr></thead><tbody>${terms}</tbody></table>
<h2>Mensagens com ocorrências (${sorted.length})</h2><table><thead><tr><th>Mensagem</th><th>Autor</th><th>Data</th><th>Informação encontrada</th></tr></thead><tbody>`;
      for (const r of sorted) yield row(r);
      yield '</tbody></table>\n</body></html>';
    })(),
  );
}
