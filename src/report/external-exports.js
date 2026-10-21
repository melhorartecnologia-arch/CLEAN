// Exportações da varredura de destinatários externos ao domínio (somente leitura): Excel, CSV
// (padrão Excel pt-BR) e HTML. O JSON é o comum das exportações. Uma linha por mensagem com
// destinatário externo, mais as apurações por remetente (quem envia para fora), por destinatário e
// por domínio.
import { writeXlsx } from './xlsx.js';
import { writeAll, csvCell, escapeHtml, REPORT_CSS, toDate, kb } from './exports.js';
import { summarizeExternal, formatDateTime, SCAN_STATUS_LABELS, MAIL_TYPE_LABELS } from './model.js';

const FIELD_TEXT = { to: 'Para', cc: 'Cc', bcc: 'Cco' };
const fieldsText = (fields) => (fields || []).map((f) => FIELD_TEXT[f] || f).join(', ');
const dateOnly = (iso) => (iso ? new Date(iso).toLocaleDateString('pt-BR') : '');

/** Conexão de e-mail no resumo: "Nome (tipo, N caixa(s))". */
function sourceText(s) {
  const scope = s.scope === 'all' ? 'todas as caixas' : `${s.mailboxCount ?? 0} caixa(s)`;
  return `${s.name} (${MAIL_TYPE_LABELS[s.type] || s.type}, ${scope})`;
}

function periodText(o) {
  const from = o.receivedAfter ? new Date(o.receivedAfter).toLocaleDateString('pt-BR') : null;
  const to = o.receivedBefore ? new Date(o.receivedBefore).toLocaleDateString('pt-BR') : null;
  if (from && to) return `de ${from} a ${to}`;
  if (from) return `a partir de ${from}`;
  if (to) return `até ${to}`;
  return 'todo o período';
}

function infoRows(scan, summary) {
  const s = scan.stats || {};
  const o = scan.options || {};
  const own = scan.listing?.ownDomains || [];
  return [
    ['Varredura', scan.name],
    ...(scan.scheduleId ? [['Agendamento', scan.scheduleName || '']] : []),
    ['Situação', SCAN_STATUS_LABELS[scan.status] || scan.status],
    ['Início', toDate(scan.startedAt)],
    ['Fim', toDate(scan.finishedAt)],
    ['Conexões de e-mail', (scan.summary?.sources || []).map(sourceText).join('; ')],
    ['Domínios internos', own.length ? own.join('; ') : 'o domínio de cada caixa analisada'],
    ['Pastas', scan.listing?.folderScope === 'sent' ? 'apenas Enviados' : 'todas as pastas'],
    ['Período', periodText(o)],
    ['Caixas varridas', s.mailboxesDone ?? 0],
    ['Mensagens verificadas', s.messagesSeen ?? 0],
    ['Mensagens com destinatário externo', summary.messages],
    ['Remetentes internos que enviaram para fora', summary.sendersTotal ?? 0],
    ['Destinatários externos (distintos)', summary.recipientsTotal],
    ['Domínios externos (distintos)', summary.domainsTotal],
    ['Mensagens de remetente externo', summary.externalSenders],
    ['Erros', s.errors ?? 0],
  ];
}

const SENDER_COLUMNS = [
  ['Remetente (quem enviou para fora)', 34],
  ['Nome', 26],
  ['Mensagens enviadas para fora', 16],
  ['Destinatários externos', 15],
  ['Domínios externos', 14],
  ['Principais domínios', 40],
  ['Primeira', 17],
  ['Última', 17],
];

const senderRow = (s) => [s.address, s.name, s.messages, s.recipients, s.domains, (s.topDomains || []).join(', '), toDate(s.first), toDate(s.last)];

// ---------------------------------------------------------------------------------------------

const RECIPIENT_COLUMNS = [
  ['Destinatário externo', 34],
  ['Nome', 26],
  ['Domínio', 26],
  ['Mensagens', 11],
  ['Caixas internas', 13],
  ['Campos', 16],
  ['Primeira', 17],
  ['Última', 17],
];

const recipientRow = (r) => [r.address, r.name, r.domain, r.messages, r.mailboxes, fieldsText(r.fields), toDate(r.first), toDate(r.last)];

const DOMAIN_COLUMNS = [
  ['Domínio externo', 30],
  ['Destinatários', 13],
  ['Mensagens', 11],
  ['Primeira', 17],
  ['Última', 17],
];

const domainRow = (d) => [d.domain, d.recipients, d.messages, toDate(d.first), toDate(d.last)];

const MESSAGE_COLUMNS = [
  ['Conexão', 16],
  ['Caixa', 28],
  ['Pasta', 22],
  ['Recebida', 17],
  ['Enviada', 17],
  ['Remetente', 32],
  ['Remetente externo', 11],
  ['Assunto', 42],
  ['Destinatários externos', 50],
  ['Domínios externos', 28],
  ['Tem anexos', 11],
  ['Tamanho (KB)', 12],
  ['Message-ID', 40],
];

const externalText = (r) => (r.external || []).map((e) => (e.name ? `${e.name} <${e.address}>` : e.address)).join('; ');
const externalDomains = (r) => [...new Set((r.external || []).map((e) => e.domain).filter(Boolean))].join('; ');

function messageRow(r) {
  return [
    r.sourceName,
    r.mailbox,
    r.folder,
    toDate(r.date),
    toDate(r.sent),
    r.from,
    r.fromExternal ? 'Sim' : 'Não',
    r.subject,
    externalText(r),
    externalDomains(r),
    r.hasAttachments === true ? 'Sim' : r.hasAttachments === false ? 'Não' : '',
    kb(r.size || 0),
    r.internetMessageId || '',
  ];
}

function sortMessages(records) {
  return records.slice().sort((a, b) => String(b.date || '').localeCompare(String(a.date || '')));
}

export async function exportExternalXlsx(scan, records, errors, out) {
  const sorted = sortMessages(records);
  const summary = summarizeExternal(sorted);
  const header = (...labels) => labels.map((v) => ({ v, s: 'header' }));
  const resumo = [[{ v: `Destinatários externos – ${scan.name}`, s: 'title' }], []];
  for (const [label, value] of infoRows(scan, summary)) resumo.push([{ v: label, s: 'bold' }, value]);
  resumo.push([], header('Domínio externo', 'Destinatários', 'Mensagens'));
  for (const d of summary.byDomain.slice(0, 100)) resumo.push([d.domain, d.recipients, d.messages]);
  const sheets = [
    { name: 'Resumo', cols: [38, 40, 14], rows: resumo },
    {
      name: 'Remetentes',
      cols: SENDER_COLUMNS.map(([, w]) => w),
      header: SENDER_COLUMNS.map(([h]) => h),
      rows: (function* () {
        for (const s of summary.bySender) yield senderRow(s);
      })(),
    },
    {
      name: 'Destinatários',
      cols: RECIPIENT_COLUMNS.map(([, w]) => w),
      header: RECIPIENT_COLUMNS.map(([h]) => h),
      rows: (function* () {
        for (const r of summary.byRecipient) yield recipientRow(r);
      })(),
    },
    {
      name: 'Domínios',
      cols: DOMAIN_COLUMNS.map(([, w]) => w),
      header: DOMAIN_COLUMNS.map(([h]) => h),
      rows: (function* () {
        for (const d of summary.byDomain) yield domainRow(d);
      })(),
    },
    {
      name: 'Mensagens',
      cols: MESSAGE_COLUMNS.map(([, w]) => w),
      header: MESSAGE_COLUMNS.map(([h]) => h),
      rows: (function* () {
        for (const r of sorted) yield messageRow(r);
      })(),
    },
  ];
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
  await writeXlsx(sheets, out, { title: `Destinatários externos – ${scan.name}` });
}

export async function exportExternalCsv(records, out) {
  // O CSV traz a apuração por destinatário externo (a visão principal da varredura).
  const summary = summarizeExternal(records);
  const header = RECIPIENT_COLUMNS.map(([h]) => h);
  await writeAll(
    out,
    (function* () {
      yield `﻿${header.map(csvCell).join(';')}\r\n`;
      for (const r of summary.byRecipient) {
        const row = recipientRow(r).map((v) => (v instanceof Date ? v.toLocaleString('pt-BR') : v));
        yield `${row.map(csvCell).join(';')}\r\n`;
      }
    })(),
  );
}

export async function exportExternalHtml(scan, records, out) {
  const sorted = sortMessages(records);
  const summary = summarizeExternal(sorted);
  const info = infoRows(scan, summary)
    .map(([label, value]) => `<tr><td>${escapeHtml(label)}</td><td>${escapeHtml(value instanceof Date ? value.toLocaleString('pt-BR') : value)}</td></tr>`)
    .join('');
  const senders = summary.bySender
    .slice(0, 1000)
    .map(
      (s) =>
        `<tr><td><b>${escapeHtml(s.address)}</b>${s.name ? `<div class="muted">${escapeHtml(s.name)}</div>` : ''}${
          (s.topDomains || []).length ? `<div class="muted">→ ${escapeHtml(s.topDomains.join(', '))}${s.domains > s.topDomains.length ? '…' : ''}</div>` : ''
        }</td><td class="num">${s.messages}</td><td class="num">${s.recipients}</td><td class="num">${s.domains}</td></tr>`,
    )
    .join('');
  const domains = summary.byDomain
    .slice(0, 100)
    .map((d) => `<tr><td>${escapeHtml(d.domain)}</td><td class="num">${d.recipients}</td><td class="num">${d.messages}</td></tr>`)
    .join('');
  const recipients = summary.byRecipient
    .slice(0, 1000)
    .map(
      (r) =>
        `<tr><td><b>${escapeHtml(r.address)}</b>${r.name ? `<div class="muted">${escapeHtml(r.name)}</div>` : ''}</td><td>${escapeHtml(r.domain)}</td><td class="num">${
          r.messages
        }</td><td>${escapeHtml(fieldsText(r.fields))}</td><td>${escapeHtml(dateOnly(r.last))}</td></tr>`,
    )
    .join('');
  const row = (r) =>
    `<tr><td><b>${escapeHtml(r.subject || '(sem assunto)')}</b><div class="muted">${escapeHtml(r.mailbox)} › ${escapeHtml(r.folder)}</div></td><td>${escapeHtml(
      r.from,
    )}${r.fromExternal ? '<div class="muted">remetente externo</div>' : ''}</td><td>${escapeHtml(externalText(r))}</td><td>${escapeHtml(formatDateTime(r.date))}</td></tr>`;
  await writeAll(
    out,
    (function* () {
      yield `<!doctype html><html lang="pt-BR"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>${escapeHtml(
        `Destinatários externos – ${scan.name}`,
      )}</title><style>${REPORT_CSS}</style></head><body>
<h1>Destinatários externos ao domínio</h1><div class="muted">${escapeHtml(scan.name)} · gerado em ${escapeHtml(new Date().toLocaleString('pt-BR'))}</div>
<h2>Resumo</h2><table class="info">${info}</table>
<h2>Remetentes que enviam para fora (${summary.sendersTotal})</h2><table><thead><tr><th>Remetente</th><th class="num">Mensagens</th><th class="num">Destinatários</th><th class="num">Domínios</th></tr></thead><tbody>${senders}</tbody></table>
<h2>Domínios externos (${summary.domainsTotal})</h2><table><thead><tr><th>Domínio</th><th class="num">Destinatários</th><th class="num">Mensagens</th></tr></thead><tbody>${domains}</tbody></table>
<h2>Destinatários externos (${summary.recipientsTotal})</h2><table><thead><tr><th>Destinatário</th><th>Domínio</th><th class="num">Mensagens</th><th>Campos</th><th>Última</th></tr></thead><tbody>${recipients}</tbody></table>
<h2>Mensagens (${sorted.length})</h2><table><thead><tr><th>Mensagem</th><th>Remetente</th><th>Destinatários externos</th><th>Recebida</th></tr></thead><tbody>`;
      for (const r of sorted) yield row(r);
      yield '</tbody></table>\n</body></html>';
    })(),
  );
}
