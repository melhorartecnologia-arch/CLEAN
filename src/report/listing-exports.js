// Exportações das listagens de e-mail (somente leitura): catálogo de contas do domínio e listagem de
// mensagens por caixa. Excel, CSV (padrão Excel pt-BR) e HTML. (O JSON é o comum das exportações.)
import { writeXlsx } from './xlsx.js';
import { writeAll, csvCell, escapeHtml, REPORT_CSS, toDate, kb } from './exports.js';
import { summarizeAccounts, summarizeMessages, formatDateTime, SCAN_STATUS_LABELS, MAIL_TYPE_LABELS } from './model.js';

const yesNo = (v) => (v === true ? 'Sim' : v === false ? 'Não' : '');
const stateText = (r) => (r.enabled === true ? 'Ativa' : r.enabled === false ? 'Inativa/bloqueada' : '');
const dateOnly = (iso) => (iso ? new Date(iso).toLocaleDateString('pt-BR') : '');

/** Conexão de e-mail no resumo: "Nome (tipo, N caixa(s))". */
function sourceText(s) {
  const scope = s.scope === 'all' ? 'todas as caixas' : `${s.mailboxCount ?? 0} caixa(s)`;
  return `${s.name} (${MAIL_TYPE_LABELS[s.type] || s.type}, ${scope})`;
}

function baseInfoRows(scan) {
  return [
    ['Listagem', scan.name],
    ...(scan.scheduleId ? [['Agendamento', scan.scheduleName || '']] : []),
    ['Situação', SCAN_STATUS_LABELS[scan.status] || scan.status],
    ['Início', toDate(scan.startedAt)],
    ['Fim', toDate(scan.finishedAt)],
    ['Conexões de e-mail', (scan.summary?.sources || []).map(sourceText).join('; ')],
  ];
}

// ---------------------------------------------------------------------------------------------
// Catálogo de contas do domínio

const ACCOUNT_COLUMNS = [
  ['Endereço principal', 32],
  ['Nome', 28],
  ['Login (UPN)', 32],
  ['Apelidos (aliases)', 40],
  ['Situação', 16],
  ['Tipo', 18],
  ['Licenciada', 11],
  ['Criada em', 14],
  ['Último acesso', 14],
  ['Departamento', 22],
  ['Cargo', 22],
  ['Local', 18],
  ['Telefone', 18],
  ['Unidade organizacional', 24],
  ['Conexão', 18],
  ['Observação', 40],
];

function accountRow(r) {
  return [
    r.address,
    r.name,
    r.login,
    (r.aliases || []).join('; '),
    stateText(r),
    r.type,
    yesNo(r.licensed),
    toDate(r.created),
    toDate(r.lastActivity),
    r.department,
    r.title,
    r.location,
    r.phone,
    r.orgUnit,
    r.sourceName,
    r.note || '',
  ];
}

function accountsInfoRows(scan, summary) {
  const s = scan.stats || {};
  return [
    ...baseInfoRows(scan),
    ['Contas listadas', summary.accounts],
    ['Contas com apelidos', summary.withAliases],
    ['Total de apelidos', summary.aliasesTotal],
    ['Contas licenciadas', summary.licensed],
    ['Contas inativas/bloqueadas', summary.disabled],
    ['Erros', s.errors ?? 0],
  ];
}

export async function exportAccountsXlsx(scan, records, errors, out) {
  const sorted = records.slice().sort((a, b) => String(a.address || '').localeCompare(String(b.address || ''), 'pt-BR'));
  const summary = summarizeAccounts(sorted);
  const header = (...labels) => labels.map((v) => ({ v, s: 'header' }));
  const resumo = [[{ v: `Contas do domínio – ${scan.name}`, s: 'title' }], []];
  for (const [label, value] of accountsInfoRows(scan, summary)) resumo.push([{ v: label, s: 'bold' }, value]);
  resumo.push([], header('Situação', 'Contas'));
  for (const g of summary.byState) resumo.push([g.label, g.accounts]);
  resumo.push([], header('Tipo', 'Contas'));
  for (const g of summary.byType) resumo.push([g.label, g.accounts]);
  const sheets = [
    { name: 'Resumo', cols: [30, 50], rows: resumo },
    {
      name: 'Contas',
      cols: ACCOUNT_COLUMNS.map(([, w]) => w),
      header: ACCOUNT_COLUMNS.map(([h]) => h),
      rows: (function* () {
        for (const r of sorted) yield accountRow(r);
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
  await writeXlsx(sheets, out, { title: `Contas do domínio – ${scan.name}` });
}

export async function exportAccountsCsv(records, out) {
  const header = ACCOUNT_COLUMNS.map(([h]) => h);
  const sorted = records.slice().sort((a, b) => String(a.address || '').localeCompare(String(b.address || ''), 'pt-BR'));
  await writeAll(
    out,
    (function* () {
      yield `﻿${header.map(csvCell).join(';')}\r\n`;
      for (const r of sorted) {
        const row = accountRow(r).map((v) => (v instanceof Date ? v.toLocaleString('pt-BR') : v));
        yield `${row.map(csvCell).join(';')}\r\n`;
      }
    })(),
  );
}

export async function exportAccountsHtml(scan, records, out) {
  const sorted = records.slice().sort((a, b) => String(a.address || '').localeCompare(String(b.address || ''), 'pt-BR'));
  const summary = summarizeAccounts(sorted);
  const info = accountsInfoRows(scan, summary)
    .map(([label, value]) => `<tr><td>${escapeHtml(label)}</td><td>${escapeHtml(value instanceof Date ? value.toLocaleString('pt-BR') : value)}</td></tr>`)
    .join('');
  const row = (r) =>
    `<tr><td><b>${escapeHtml(r.address)}</b>${r.name ? `<div class="muted">${escapeHtml(r.name)}</div>` : ''}${
      (r.aliases || []).length ? `<div class="path">Apelidos: ${escapeHtml(r.aliases.join(', '))}</div>` : ''
    }</td><td>${escapeHtml(stateText(r))}${r.licensed === true ? '<div class="muted">licenciada</div>' : ''}</td><td>${escapeHtml(r.type)}</td><td>${escapeHtml(
      [r.department, r.title].filter(Boolean).join(' · '),
    )}</td><td>${escapeHtml(dateOnly(r.created))}</td><td>${escapeHtml(r.note || '')}</td></tr>`;
  await writeAll(
    out,
    (function* () {
      yield `<!doctype html><html lang="pt-BR"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>${escapeHtml(
        `Contas do domínio – ${scan.name}`,
      )}</title><style>${REPORT_CSS}</style></head><body>
<h1>Contas do domínio</h1><div class="muted">${escapeHtml(scan.name)} · gerado em ${escapeHtml(new Date().toLocaleString('pt-BR'))}</div>
<h2>Resumo</h2><table class="info">${info}</table>
<h2>Contas (${sorted.length})</h2><table><thead><tr><th>Conta</th><th>Situação</th><th>Tipo</th><th>Departamento / cargo</th><th>Criada em</th><th>Observação</th></tr></thead><tbody>`;
      for (const r of sorted) yield row(r);
      yield '</tbody></table>\n</body></html>';
    })(),
  );
}

// ---------------------------------------------------------------------------------------------
// Listagem de mensagens por caixa

const MESSAGE_COLUMNS = [
  ['Conexão', 16],
  ['Caixa', 28],
  ['Pasta', 22],
  ['Recebida', 17],
  ['Enviada', 17],
  ['Remetente', 32],
  ['Destinatários', 40],
  ['Cc', 30],
  ['Assunto', 44],
  ['Tem anexos', 11],
  ['Tamanho (KB)', 12],
  ['Message-ID', 40],
  ['Link (Outlook na Web)', 30],
];

function messageRow(r) {
  return [
    r.sourceName,
    r.mailbox,
    r.folder,
    toDate(r.date),
    toDate(r.sent),
    r.from,
    (r.to || []).join('; '),
    (r.cc || []).join('; '),
    r.subject,
    r.hasAttachments === true ? 'Sim' : r.hasAttachments === false ? 'Não' : '',
    kb(r.size || 0),
    r.internetMessageId || '',
    r.webLink || '',
  ];
}

function sortMessages(records) {
  return records.slice().sort((a, b) => a.mailbox.localeCompare(b.mailbox, 'pt-BR') || String(b.date || '').localeCompare(String(a.date || '')));
}

function messagesInfoRows(scan, summary) {
  const s = scan.stats || {};
  const o = scan.options || {};
  return [
    ...baseInfoRows(scan),
    ['Recebidas a partir de', o.receivedAfter ? toDate(o.receivedAfter) : 'todas'],
    ['Lixeira / Lixo eletrônico', `${o.includeTrash ? 'inclui' : 'ignora'} a lixeira; ${o.includeJunk ? 'inclui' : 'ignora'} o lixo eletrônico`],
    ['Caixas listadas', s.mailboxesDone ?? 0],
    ['Mensagens listadas', summary.messages],
    ['Com anexos', summary.withAttachments],
    ['Espaço total', `${kb(summary.bytes)} KB`],
    ['E-mail mais antigo', summary.oldest ? toDate(summary.oldest) : ''],
    ['E-mail mais recente', summary.newest ? toDate(summary.newest) : ''],
    ['Erros', s.errors ?? 0],
  ];
}

export async function exportMessagesXlsx(scan, records, errors, out) {
  const sorted = sortMessages(records);
  const summary = summarizeMessages(sorted);
  const header = (...labels) => labels.map((v) => ({ v, s: 'header' }));
  const resumo = [[{ v: `Listagem de mensagens – ${scan.name}`, s: 'title' }], []];
  for (const [label, value] of messagesInfoRows(scan, summary)) resumo.push([{ v: label, s: 'bold' }, value]);
  resumo.push([], header('Caixa', 'Nome', 'Mensagens', 'Espaço (KB)'));
  for (const m of summary.byMailbox) resumo.push([m.mailbox, m.name, m.messages, kb(m.bytes)]);
  resumo.push([], header('Pasta', 'Mensagens'));
  for (const f of summary.byFolder) resumo.push([f.folder, f.messages]);
  const sheets = [
    { name: 'Resumo', cols: [30, 40, 12, 12], rows: resumo },
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
  await writeXlsx(sheets, out, { title: `Listagem de mensagens – ${scan.name}` });
}

export async function exportMessagesCsv(records, out) {
  const header = MESSAGE_COLUMNS.map(([h]) => h);
  await writeAll(
    out,
    (function* () {
      yield `﻿${header.map(csvCell).join(';')}\r\n`;
      for (const r of sortMessages(records)) {
        const row = messageRow(r).map((v) => (v instanceof Date ? v.toLocaleString('pt-BR') : v));
        yield `${row.map(csvCell).join(';')}\r\n`;
      }
    })(),
  );
}

export async function exportMessagesHtml(scan, records, out) {
  const sorted = sortMessages(records);
  const summary = summarizeMessages(sorted);
  const info = messagesInfoRows(scan, summary)
    .map(([label, value]) => `<tr><td>${escapeHtml(label)}</td><td>${escapeHtml(value instanceof Date ? value.toLocaleString('pt-BR') : value)}</td></tr>`)
    .join('');
  const boxes = summary.byMailbox
    .map((m) => `<tr><td>${escapeHtml(m.mailbox)}</td><td>${escapeHtml(m.name)}</td><td class="num">${m.messages}</td></tr>`)
    .join('');
  const row = (r) =>
    `<tr><td><b>${escapeHtml(r.subject || '(sem assunto)')}</b><div class="muted">${escapeHtml(r.mailbox)} › ${escapeHtml(r.folder)}${
      r.hasAttachments ? ' · com anexos' : ''
    }</div></td><td>${escapeHtml(r.from)}${(r.to || []).length ? `<div class="muted">para ${escapeHtml(r.to.join(', '))}</div>` : ''}</td><td>${escapeHtml(
      formatDateTime(r.date),
    )}</td></tr>`;
  await writeAll(
    out,
    (function* () {
      yield `<!doctype html><html lang="pt-BR"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>${escapeHtml(
        `Listagem de mensagens – ${scan.name}`,
      )}</title><style>${REPORT_CSS}</style></head><body>
<h1>Listagem de mensagens</h1><div class="muted">${escapeHtml(scan.name)} · gerado em ${escapeHtml(new Date().toLocaleString('pt-BR'))}</div>
<h2>Resumo</h2><table class="info">${info}</table>
<h2>Caixas</h2><table><thead><tr><th>Caixa</th><th>Nome</th><th class="num">Mensagens</th></tr></thead><tbody>${boxes}</tbody></table>
<h2>Mensagens (${sorted.length})</h2><table><thead><tr><th>Mensagem</th><th>Remetente</th><th>Recebida</th></tr></thead><tbody>`;
      for (const r of sorted) yield row(r);
      yield '</tbody></table>\n</body></html>';
    })(),
  );
}
