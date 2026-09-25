// Exportação dos relatórios de análises de e-mail: Excel, CSV (padrão Excel pt-BR) e HTML.
// (O JSON é o mesmo das análises de arquivos.)
import { writeXlsx } from './xlsx.js';
import { writeAll, csvCell, escapeHtml, REPORT_CSS, toDate, kb, deletionsSheet, deletionInfoRows } from './exports.js';
import {
  summarizeMail,
  sampleText,
  formatDateTime,
  realAttachments,
  deletionText,
  STATUS_LABELS,
  LOCATION_LABELS,
  SCAN_STATUS_LABELS,
  MAIL_TYPE_LABELS,
  MAIL_DELETION_LABELS,
} from './model.js';

function sortMessages(records) {
  return records.slice().sort((a, b) => a.mailbox.localeCompare(b.mailbox, 'pt-BR') || String(b.date || '').localeCompare(String(a.date || '')));
}

function locations(record) {
  return [...new Set(record.matches.map((m) => LOCATION_LABELS[m.location]))].join(', ');
}

function attachmentNames(record) {
  return realAttachments(record)
    .map((a) => a.name)
    .join('; ');
}

const MESSAGE_COLUMNS = [
  ['Conexão', 16],
  ['Caixa', 28],
  ['Pasta', 22],
  ['Data', 17],
  ['Remetente', 32],
  ['Destinatários', 40],
  ['Cc', 30],
  ['Assunto', 44],
  ['Anexos', 36],
  ['Tamanho (KB)', 12],
  ['Termos encontrados', 36],
  ['Ocorrências', 11],
  ['Encontrado em', 24],
  ['Situação do conteúdo', 20],
  ['Exclusão', 34],
  ['Message-ID', 40],
  ['Link (Outlook na Web)', 30],
];

function messageRow(r) {
  return [
    r.sourceName,
    r.mailbox,
    r.folder,
    toDate(r.date),
    r.from,
    (r.to || []).join('; '),
    (r.cc || []).join('; '),
    r.subject,
    attachmentNames(r),
    kb(r.size || 0),
    r.terms.join('; '),
    r.occurrences,
    locations(r),
    STATUS_LABELS[r.contentStatus] || r.contentStatus || '',
    deletionText(r.deletion, 'mail'),
    r.internetMessageId || '',
    r.webLink || '',
  ];
}

const MATCH_COLUMNS = [
  ['Caixa', 28],
  ['Pasta', 20],
  ['Data', 17],
  ['Remetente', 30],
  ['Assunto', 36],
  ['Termo', 22],
  ['Lista', 18],
  ['Encontrado em', 18],
  ['Ocorrências', 11],
  ['Valores encontrados', 32],
  ['Exemplo', 70],
  ['Outros exemplos', 70],
];

function matchRows(r) {
  return r.matches.map((m) => [
    r.mailbox,
    r.folder,
    toDate(r.date),
    r.from,
    r.subject,
    m.term,
    m.list,
    LOCATION_LABELS[m.location],
    m.count,
    m.values.join('; '),
    sampleText(m.samples[0]),
    m.samples.slice(1).map(sampleText).join('\n'),
  ]);
}

function sourceText(s) {
  const scope = s.scope === 'all' ? 'todas as caixas' : `${s.mailboxCount ?? 0} caixa(s)`;
  return `${s.name} (${MAIL_TYPE_LABELS[s.type] || s.type}, ${scope})`;
}

function scanInfoRows(scan) {
  const s = scan.stats || {};
  const o = scan.options || {};
  const checks = [
    o.checkSubject && 'assunto',
    o.checkBody && 'corpo',
    o.checkAttachmentNames && 'nomes dos anexos',
    o.checkAttachments && 'conteúdo dos anexos',
    o.checkAddresses && 'remetente e destinatários',
  ].filter(Boolean);
  return [
    ['Análise', scan.name],
    ['Situação', SCAN_STATUS_LABELS[scan.status] || scan.status],
    ['Início', toDate(scan.startedAt)],
    ['Fim', toDate(scan.finishedAt)],
    ['Conexões de e-mail', (scan.summary?.sources || []).map(sourceText).join('; ')],
    ['Listas de referência', (scan.summary?.lists || []).map((l) => `${l.name} (${l.termCount} termos)`).join('; ')],
    ['Verificações', checks.join(', ')],
    ['Recebidas a partir de', o.receivedAfter ? toDate(o.receivedAfter) : 'todas'],
    ['Lixeira / Lixo eletrônico', `${o.includeTrash ? 'inclui' : 'ignora'} a lixeira; ${o.includeJunk ? 'inclui' : 'ignora'} o lixo eletrônico`],
    ['Caixas analisadas', s.mailboxesDone ?? 0],
    ['Caixas ignoradas (sem e-mail)', s.mailboxesSkipped ?? 0],
    ['Mensagens verificadas', s.messagesSeen ?? 0],
    ['Mensagens com ocorrências', s.messagesMatched ?? 0],
    ['Total de ocorrências', s.occurrences ?? 0],
    ['Anexos lidos', s.attachmentsAnalyzed ?? 0],
    ['Anexos protegidos por senha', s.attachmentsEncrypted ?? 0],
    ['Mensagens criptografadas', s.messagesEncrypted ?? 0],
    ['Erros', s.errors ?? 0],
    ...deletionInfoRows(o, s, { noun: 'Excluídas', gone: 'Já não existiam', changed: 'Alteradas depois da análise (mantidas)' }),
  ];
}

export async function exportMailXlsx(scan, records, errors, out, { deletions = [], records: all = records } = {}) {
  const sorted = sortMessages(records);
  const summary = summarizeMail(sorted);
  const header = (...labels) => labels.map((v) => ({ v, s: 'header' }));
  const resumo = [[{ v: `Relatório CLEAN – ${scan.name}`, s: 'title' }], []];
  for (const [label, value] of scanInfoRows(scan)) resumo.push([{ v: label, s: 'bold' }, value]);
  resumo.push([], header('Termos encontrados', 'Lista', 'Mensagens', 'Ocorrências'));
  for (const t of summary.byTerm) resumo.push([t.term, t.list, t.messages, t.occurrences]);
  resumo.push([], header('Caixa', 'Nome', 'Mensagens', 'Ocorrências'));
  for (const m of summary.byMailbox) resumo.push([m.mailbox, m.name, m.messages, m.occurrences]);
  resumo.push([], header('Remetente', '', 'Mensagens', 'Ocorrências'));
  for (const s of summary.bySender) resumo.push([s.label, '', s.messages, s.occurrences]);

  const sheets = [
    { name: 'Resumo', cols: [34, 50, 12, 12], rows: resumo },
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
        [['Caixa', 28], ['Pasta', 20], ['Assunto', 40]],
        (r, d) => (r ? [r.mailbox || '', r.folder || '', r.subject || ''] : [d.item || '', '', '']),
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
  await writeXlsx(sheets, out, { title: `Relatório CLEAN – ${scan.name}` });
}

/** CSV com uma linha por mensagem, termo e local (assunto, corpo, anexo...). */
export async function exportMailCsv(records, out) {
  const header = [...MATCH_COLUMNS.map(([h]) => h), 'Destinatários', 'Anexos', 'Conexão', 'Exclusão'];
  await writeAll(
    out,
    (function* () {
      yield `\uFEFF${header.map(csvCell).join(';')}\r\n`;
      for (const r of sortMessages(records)) {
        const extra = [(r.to || []).join('; '), attachmentNames(r), r.sourceName, deletionText(r.deletion, 'mail')];
        for (const row of matchRows(r)) yield `${[...row, ...extra].map(csvCell).join(';')}\r\n`;
      }
    })(),
  );
}

export async function exportMailHtml(scan, records, out) {
  const sorted = sortMessages(records);
  const summary = summarizeMail(sorted);
  const info = scanInfoRows(scan)
    .map(([label, value]) => `<tr><td>${escapeHtml(label)}</td><td>${escapeHtml(value instanceof Date ? value.toLocaleString('pt-BR') : value)}</td></tr>`)
    .join('');
  const terms = summary.byTerm
    .map((t) => `<tr><td>${escapeHtml(t.term)}</td><td>${escapeHtml(t.list)}</td><td class="num">${t.messages}</td><td class="num">${t.occurrences}</td></tr>`)
    .join('');
  const boxes = summary.byMailbox
    .map((m) => `<tr><td>${escapeHtml(m.mailbox)}</td><td class="num">${m.messages}</td><td class="num">${m.occurrences}</td></tr>`)
    .join('');
  const sample = (s) =>
    s ? `<div class="sample">${s.where ? `<b>${escapeHtml(s.where)}:</b> ` : ''}${escapeHtml(s.before)}<mark>${escapeHtml(s.match)}</mark>${escapeHtml(s.after)}</div>` : '';
  const row = (r) => {
    const found = r.matches
      .map((m) => `<div><span class="term">${escapeHtml(m.term)}</span> ${escapeHtml(LOCATION_LABELS[m.location])} · ${m.count}×${m.samples.map(sample).join('')}</div>`)
      .join('');
    const files = attachmentNames(r);
    const deleted = r.deletion ? `<div class="muted">${escapeHtml(deletionText(r.deletion, 'mail'))}</div>` : '';
    return `<tr><td><b>${escapeHtml(r.subject || '(sem assunto)')}</b><div class="muted">${escapeHtml(r.mailbox)} › ${escapeHtml(r.folder)}</div>${files ? `<div class="path">Anexos: ${escapeHtml(files)}</div>` : ''}${deleted}</td><td>${escapeHtml(r.from)}<div class="muted">para ${escapeHtml((r.to || []).join(', '))}</div></td><td>${escapeHtml(formatDateTime(r.date))}</td><td>${found}</td></tr>`;
  };
  await writeAll(
    out,
    (function* () {
      yield `<!doctype html><html lang="pt-BR"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>${escapeHtml(`Relatório CLEAN – ${scan.name}`)}</title><style>${REPORT_CSS}</style></head><body>
<h1>Relatório CLEAN – e-mail</h1><div class="muted">${escapeHtml(scan.name)} · gerado em ${escapeHtml(new Date().toLocaleString('pt-BR'))}</div>
<h2>Resumo</h2><table class="info">${info}</table>
<h2>Termos encontrados</h2><table><thead><tr><th>Termo</th><th>Lista</th><th class="num">Mensagens</th><th class="num">Ocorrências</th></tr></thead><tbody>${terms}</tbody></table>
<h2>Caixas</h2><table><thead><tr><th>Caixa</th><th class="num">Mensagens</th><th class="num">Ocorrências</th></tr></thead><tbody>${boxes}</tbody></table>
<h2>Mensagens com ocorrências (${sorted.length})</h2><table><thead><tr><th>Mensagem</th><th>Remetente</th><th>Data</th><th>Informação encontrada</th></tr></thead><tbody>`;
      for (const r of sorted) yield row(r);
      yield '</tbody></table>\n</body></html>';
    })(),
  );
}
