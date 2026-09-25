// Relatório de uma análise (de arquivos ou de e-mail): progresso, indicadores, gráficos, filtros,
// resultados e exportações. O que muda entre os dois tipos fica nos perfis FILES e MAIL.
import { get, post } from '../api.js';
import {
  html,
  render as paint,
  icon,
  toast,
  confirmDialog,
  fmtNum,
  fmtCompact,
  fmtDateTime,
  fmtDuration,
  fmtBytes,
  statusBadge,
  plural,
  bindTooltips,
  copyText,
  debounce,
  redraw,
} from '../ui.js';
import { replaceQuery, setActiveNav } from '../nav.js';

const CONTENT_STATUS = {
  ok: 'Analisado',
  partial: 'Analisado parcialmente',
  encrypted: 'Protegido por senha',
  unsupported: 'Formato sem texto',
  'skipped-size': 'Muito grande (só o nome)',
  empty: 'Vazio',
  error: 'Erro de leitura',
  'not-requested': 'Conteúdo não verificado',
};
const TOP = 10;
const PAGE_SIZE = 50;

const isActive = (scan) => scan.status === 'running' || scan.status === 'queued';
const sampleHtml = (s) => html`<div class="sample">${s.where ? html`<span class="where">${s.where}</span>` : ''}${s.before}<mark>${s.match}</mark>${s.after}</div>`;

/** Lista de ocorrências de um resultado (termo, lista, local, contagem, valores e exemplos). */
function matchesHtml(record, phrase) {
  return record.matches.map(
    (mt) => html`<div class="match">
      <div class="match-head">
        <span class="chip"><b>${mt.term}</b></span>
        <span class="muted small">lista ${mt.list} · ${phrase[mt.location] || mt.location} · ${plural(mt.count, 'ocorrência', 'ocorrências')}${mt.truncated ? '+' : ''}</span>
      </div>
      ${mt.values?.length && mt.kind === 'regex' ? html`<div class="small"><span class="muted">Valores:</span> ${mt.values.join(' · ')}</div>` : ''}
      ${mt.samples.map(sampleHtml)}
    </div>`,
  );
}

const option = (value, label, current) => html`<option value="${value}" ${current === value ? 'selected' : ''}>${label}</option>`;

// ---------- Exclusão dos itens encontrados ----------

/** Textos da exclusão no gênero do item ("arquivo excluído", "mensagem excluída"). */
const DELETION_WORDS = {
  arquivo: {
    o: 'o',
    notAllowed: 'Exclusão não permitida neste repositório (ative "Permitir exclusão" em Repositórios).',
    removed: 'O repositório deste arquivo foi removido do cadastro: a exclusão pelo relatório não está disponível.',
  },
  mensagem: {
    o: 'a',
    notAllowed: 'Exclusão não permitida nesta conexão (ative "Permitir exclusão" em Caixas de e-mail).',
    removed: 'A conexão desta mensagem foi removida do cadastro: a exclusão pelo relatório não está disponível.',
  },
};

function deletionLabel(status, o) {
  const labels = {
    deleted: `Excluíd${o}`,
    missing: `Não encontrad${o} (já excluíd${o} ou movid${o})`,
    changed: `Não excluíd${o}: alterad${o} depois da análise`,
    failed: 'Falha na exclusão',
  };
  return labels[status] || status;
}

const deletionFilter = (filters, o) => html`<label class="field"><span>Exclusão</span>
    <select name="deletion">
      <option value="">${o === 'a' ? 'Todas' : 'Todos'}</option>
      ${option('kept', `Não excluíd${o}s`, filters.deletion)}
      ${option('deleted', `Excluíd${o}s`, filters.deletion)}
      ${option('missing', `Não encontrad${o}s ao excluir`, filters.deletion)}
      ${option('failed', `Com falha ou alterad${o}s`, filters.deletion)}
    </select>
  </label>`;

const isGone = (d) => d?.status === 'deleted' || d?.status === 'missing';

function deletionChip(r, o) {
  const d = r.deletion;
  if (!d) return '';
  if (d.status === 'failed') return html` <span class="chip danger">falha ao excluir</span>`;
  if (d.status === 'changed') return html` <span class="chip">mantid${o}: alterad${o}</span>`;
  return html` <span class="chip deleted">${d.status === 'missing' ? `não encontrad${o}` : d.method === 'trash' ? 'na lixeira' : `excluíd${o}`}</span>`;
}

/** Situação da exclusão e botão "Excluir" no detalhe de um item. */
function deletionBlock(r, noun, { active, deleting }) {
  const w = DELETION_WORDS[noun];
  const d = r.deletion;
  const how = d ? (d.mode === 'auto' ? `exclusão automática da análise${d.by ? ` iniciada por ${d.by}` : ''}` : `exclusão manual${d.by ? ` por ${d.by}` : ''}`) : '';
  const where = d?.status === 'deleted' ? (d.note ? ` (${d.note})` : d.method === 'trash' ? ` (movid${w.o} para a lixeira)` : '') : '';
  const status = d
    ? html`<p class="small ${d.status === 'failed' ? 'danger-text' : ''}"><b>${deletionLabel(d.status, w.o)}</b> em ${fmtDateTime(d.at)}${where} — ${how}${d.status === 'failed' && d.error ? html`<br />${d.error}` : ''}</p>`
    : html`<p class="muted small">Não excluíd${w.o}.</p>`;
  let action = '';
  if (deleting.has(r.id) || r.deleting) {
    action = html`<button type="button" class="btn small danger" disabled>${icon('trash')} Excluindo…</button>`;
  } else if (r.canDelete) {
    action = html`<button type="button" class="btn small danger" data-action="delete-item" data-rid="${r.id}">${icon('trash')} ${d && !isGone(d) ? 'Tentar excluir de novo' : `Excluir ${noun}`}</button>`;
  } else if (!isGone(d)) {
    action = html`<p class="muted small">${active ? 'A exclusão manual fica disponível ao fim da análise.' : r.deleteBlocked === 'removed' ? w.removed : w.notAllowed}</p>`;
  }
  return html`<h4 class="spaced">Exclusão</h4>${status}${action}`;
}

// ---------------------------------------------------------------------------------------------
// Análises de arquivos

const SOURCE = { audit: 'Log de auditoria', metadata: 'Metadados do documento', owner: 'Proprietário do arquivo (NTFS)' };
const SOURCE_SHORT = { audit: 'auditoria', metadata: 'metadados', owner: 'proprietário' };
const FILE_LOCATION = { name: 'nome', content: 'conteúdo' };

function folderOf(record) {
  const rel = record.relativePath || '';
  const idx = Math.max(rel.lastIndexOf('\\'), rel.lastIndexOf('/'));
  return idx === -1 ? record.repositoryName : `${record.repositoryName} › ${rel.slice(0, idx)}`;
}

const FILES = {
  base: '#/analises',
  nav: 'analises',
  filterKeys: ['q', 'term', 'user', 'location', 'extension', 'deletion', 'sort', 'page'],
  criteria: ['q', 'term', 'user', 'location', 'extension', 'deletion'],
  descSorts: new Set(['modified', 'occurrences', 'terms', 'size']),
  defaultSort: 'path',
  noun: ['arquivo', 'arquivos'],
  o: 'o', // gênero: "arquivos excluídos"
  resultsTitle: 'Arquivos com ocorrências',
  views: { terms: 'chart', users: 'chart' }, // forma inicial de exibição de cada gráfico

  subtitle: (scan) => {
    const s = scan.summary || {};
    return `${(s.repositories || []).map((r) => r.name).join(', ')} · ${(s.lists || []).map((l) => `${l.name} (${fmtNum(l.termCount)})`).join(', ')}`;
  },

  filterFields: (filters) => html`<label class="field grow"><span>Buscar</span><input type="search" name="q" value="${filters.q}" placeholder="Caminho, usuário ou termo" /></label>
    <label class="field"><span>Termo</span><select name="term"><option value="">Todos</option></select></label>
    <label class="field"><span>Último usuário</span><select name="user"><option value="">Todos</option></select></label>
    <label class="field"><span>Encontrado em</span>
      <select name="location">
        <option value="">Nome ou conteúdo</option>
        ${option('name', 'Nome', filters.location)}
        ${option('content', 'Conteúdo', filters.location)}
      </select>
    </label>
    <label class="field"><span>Extensão</span><select name="extension"><option value="">Todas</option></select></label>
    ${deletionFilter(filters, 'o')}
    <label class="field"><span>Ordenar por</span>
      <select name="sort">
        ${[
          ['path', 'Caminho'],
          ['occurrences', 'Mais ocorrências'],
          ['terms', 'Mais termos'],
          ['modified', 'Modificados recentemente'],
          ['lastUser', 'Último usuário'],
          ['size', 'Maiores arquivos'],
        ].map(([value, label]) => option(value, label, filters.sort))}
      </select>
    </label>`,

  fillOptions: (form, options, filters, fill) => {
    const o = options || { terms: [], users: [], extensions: [] };
    fill(form.elements.term, o.terms, filters.term);
    fill(form.elements.user, o.users, filters.user);
    fill(form.elements.extension, o.extensions, filters.extension);
  },

  progress: (st) => html`<span><b>${fmtNum(st.filesSeen)}</b> arquivos verificados</span>
    <span><b>${fmtNum(st.directories)}</b> pastas</span>
    <span><b>${fmtNum(st.filesMatched)}</b> com ocorrências</span>
    <span><b>${fmtBytes(st.bytesAnalyzed)}</b> de conteúdo lido</span>
    <span><b>${fmtNum(st.errors)}</b> erros</span>
    <span>repositório <b>${Math.min((st.repositoriesDone || 0) + 1, st.repositoriesTotal || 1)}</b> de <b>${st.repositoriesTotal || 1}</b></span>`,

  tiles: (st) => {
    const pct = st.filesSeen ? Math.round((st.filesMatched / st.filesSeen) * 1000) / 10 : 0;
    const notRead = (st.contentEncrypted || 0) + (st.contentSkippedSize || 0) + (st.contentErrors || 0);
    return html`<div class="tile"><div class="label">Arquivos verificados</div><div class="value">${fmtCompact(st.filesSeen)}</div><div class="detail">em ${plural(st.directories || 0, 'pasta', 'pastas')}${st.filesSkippedByDate ? ` · ${fmtNum(st.filesSkippedByDate)} fora do período` : ''}</div></div>
      <div class="tile"><div class="label">Arquivos com ocorrências</div><div class="value">${fmtCompact(st.filesMatched)}</div><div class="detail">${pct.toLocaleString('pt-BR')}% dos verificados</div></div>
      <div class="tile"><div class="label">Ocorrências</div><div class="value">${fmtCompact(st.occurrences)}</div><div class="detail">somando nome e conteúdo</div></div>
      <div class="tile"><div class="label">Conteúdos lidos</div><div class="value">${fmtCompact(st.contentAnalyzed)}</div><div class="detail">${notRead ? `${fmtNum(st.contentEncrypted)} com senha · ${fmtNum(st.contentSkippedSize)} grandes · ${fmtNum(st.contentErrors)} com erro` : fmtBytes(st.bytesAnalyzed)}</div></div>
      <div class="tile"><div class="label">Erros de acesso ou leitura</div><div class="value">${fmtCompact(st.errors)}</div><div class="detail">${st.errors ? 'veja a aba Erros' : 'nenhum'}</div></div>`;
  },

  charts: (summary, barChart) => {
    const terms = summary.byTerm.map((t) => ({
      label: t.term,
      value: t.files,
      filterValue: t.term,
      tipValue: `${plural(t.files, 'arquivo', 'arquivos')} · ${plural(t.occurrences, 'ocorrência', 'ocorrências')}`,
      tipLabel: `${t.term} — lista ${t.list}${t.inName ? ` · ${fmtNum(t.inName)} no nome` : ''}`,
      raw: t,
    }));
    const users = summary.byUser.map((u) => ({
      label: u.user,
      value: u.files,
      filterValue: u.identified ? u.user : '',
      tipValue: `${plural(u.files, 'arquivo', 'arquivos')} · ${plural(u.occurrences, 'ocorrência', 'ocorrências')}`,
      tipLabel: `${u.user} — ${Object.entries(u.sources).map(([k, n]) => `${SOURCE_SHORT[k]}: ${n}`).join(', ') || 'sem fonte'}`,
      raw: u,
    }));
    return html`${barChart({
      key: 'terms',
      title: 'Termos encontrados',
      subtitle: 'Arquivos em que cada termo aparece. Clique para filtrar.',
      rows: terms,
      filterKey: 'term',
      emptyText: 'Nenhum termo encontrado.',
      tableHead: html`<tr><th>Termo</th><th>Lista</th><th class="num">Arquivos</th><th class="num">Ocorrências</th><th class="num">No nome</th></tr>`,
      tableRow: (r) => html`<tr><td>${r.raw.term}</td><td>${r.raw.list}</td><td class="num">${fmtNum(r.raw.files)}</td><td class="num">${fmtNum(r.raw.occurrences)}</td><td class="num">${fmtNum(r.raw.inName)}</td></tr>`,
    })}
    ${barChart({
      key: 'users',
      title: 'Últimos usuários',
      subtitle: 'Quem interagiu por último com os arquivos encontrados. Clique para filtrar.',
      rows: users,
      filterKey: 'user',
      emptyText: 'Nenhum arquivo encontrado.',
      tableHead: html`<tr><th>Usuário</th><th>Fontes</th><th class="num">Arquivos</th><th class="num">Ocorrências</th></tr>`,
      tableRow: (r) => html`<tr><td>${r.raw.user}</td><td class="small">${Object.entries(r.raw.sources).map(([k, n]) => `${SOURCE_SHORT[k]}: ${n}`).join(', ')}</td><td class="num">${fmtNum(r.raw.files)}</td><td class="num">${fmtNum(r.raw.occurrences)}</td></tr>`,
    })}`;
  },

  tableHead: html`<tr><th><span class="sr-only">Detalhes</span></th><th>Arquivo</th><th>Último usuário</th><th>Modificado em</th><th>Informação encontrada</th></tr>`,

  row: (r) => html`<td><div class="name">${r.name}</div>${deletionChip(r, 'o')}<div class="path">${folderOf(r)}</div></td>
    <td>${r.lastUser ? html`${r.lastUser}<div><span class="chip source">${SOURCE_SHORT[r.lastUserSource]}</span></div>` : html`<span class="muted">não identificado</span>`}</td>
    <td class="nowrap">${fmtDateTime(r.modified)}</td>
    <td><div class="chips">${r.matches.map((m) => html`<span class="chip"><b>${m.term}</b> ${fmtNum(m.count)}× · ${FILE_LOCATION[m.location]}</span>`)}</div></td>`,

  rowLabel: (r) => r.name,

  detail: (r, ctx) => {
    const m = r.metadata || {};
    const a = r.audit;
    return html`<div class="detail-grid">
      <div>
        <h4>Arquivo</h4>
        <dl class="kv">
          <dt>Caminho</dt>
          <dd><span class="mono">${r.path}</span> <button type="button" class="btn small" data-action="copy" data-copy="${r.path}" data-copied="Caminho copiado.">${icon('copy')} Copiar</button></dd>
          <dt>Tamanho</dt><dd>${fmtBytes(r.size)}</dd>
          <dt>Criado em</dt><dd>${fmtDateTime(r.created)}</dd>
          <dt>Modificado em</dt><dd>${fmtDateTime(r.modified)}</dd>
          <dt>Tipo</dt><dd>${(r.contentType || r.extension || '—').toString().toUpperCase()} · ${CONTENT_STATUS[r.contentStatus] || r.contentStatus || '—'}</dd>
          ${r.contentNote ? html`<dt>Observação</dt><dd>${r.contentNote}</dd>` : ''}
          ${m.title ? html`<dt>Título</dt><dd>${m.title}</dd>` : ''}
        </dl>
        ${deletionBlock(r, 'arquivo', ctx)}
      </div>
      <div>
        <h4>Quem interagiu com o arquivo</h4>
        <dl class="kv">
          <dt>Último usuário</dt><dd><b>${r.lastUser || 'não identificado'}</b>${r.lastUserSource ? html`<br /><span class="muted small">fonte: ${SOURCE[r.lastUserSource]}</span>` : ''}</dd>
          ${a ? html`<dt>Último acesso (auditoria)</dt><dd>${a.user} · ${a.action} · ${fmtDateTime(a.time)}</dd>` : ''}
          ${a?.lastWrite ? html`<dt>Última alteração (auditoria)</dt><dd>${a.lastWrite.user} · ${a.lastWrite.action} · ${fmtDateTime(a.lastWrite.time)}</dd>` : ''}
          ${m.lastModifiedBy ? html`<dt>Salvo por último por</dt><dd>${m.lastModifiedBy}${m.modified ? html` <span class="muted small">em ${fmtDateTime(m.modified)}</span>` : ''}</dd>` : ''}
          ${m.author ? html`<dt>Autor</dt><dd>${m.author}${m.created ? html` <span class="muted small">em ${fmtDateTime(m.created)}</span>` : ''}</dd>` : ''}
          <dt>Proprietário (NTFS)</dt><dd>${r.owner || html`<span class="muted">${r.ownerError ? `não obtido: ${r.ownerError}` : 'não verificado'}</span>`}</dd>
        </dl>
      </div>
      <div>
        <h4>Informação encontrada</h4>
        ${matchesHtml(r, { name: 'no nome', content: 'no conteúdo' })}
      </div>
    </div>`;
  },

  empty: { filtered: 'Nenhum arquivo corresponde aos filtros.', running: 'Nenhuma ocorrência encontrada até agora.', none: 'Nenhum termo da lista foi encontrado nos arquivos analisados.' },
  errors: {
    column: 'Caminho',
    help: 'Pastas ou arquivos que a conta do CLEAN não conseguiu abrir (permissão, arquivo em uso, caminho longo...). Eles não foram analisados.',
    empty: 'Nenhum erro de acesso ou leitura.',
  },
};

// ---------------------------------------------------------------------------------------------
// Análises de e-mail

const TYPE_LABELS = { graph: 'Microsoft 365', gmail: 'Google Workspace', imap: 'IMAP' };
const MAIL_LOCATION = { subject: 'assunto', body: 'corpo', attachmentName: 'nome do anexo', attachment: 'anexo', address: 'remetente/destinatários' };
const MAIL_PHRASE = { subject: 'no assunto', body: 'no corpo', attachmentName: 'no nome do anexo', attachment: 'no conteúdo do anexo', address: 'no remetente ou destinatários' };
const LOCATION_TITLES = { subject: 'Assunto', body: 'Corpo', attachmentName: 'Nome do anexo', attachment: 'Conteúdo do anexo', address: 'Remetente/destinatários' };

const MAIL = {
  base: '#/email/analises',
  nav: 'email-analises',
  filterKeys: ['q', 'term', 'mailbox', 'sender', 'location', 'deletion', 'sort', 'page'],
  criteria: ['q', 'term', 'mailbox', 'sender', 'location', 'deletion'],
  descSorts: new Set(['date', 'occurrences', 'terms', 'size']),
  defaultSort: 'date',
  noun: ['mensagem', 'mensagens'],
  o: 'a', // gênero: "mensagens excluídas"
  resultsTitle: 'Mensagens com ocorrências',
  views: { terms: 'chart', mailboxes: 'chart', senders: 'chart', locations: 'chart' },

  subtitle: (scan) => {
    const s = scan.summary || {};
    const sources = (s.sources || []).map((x) => `${x.name} (${TYPE_LABELS[x.type] || x.type}${x.scope === 'all' ? ', todas as caixas' : ''})`);
    return `${sources.join(', ')} · ${(s.lists || []).map((l) => `${l.name} (${fmtNum(l.termCount)})`).join(', ')}`;
  },

  filterFields: (filters) => html`<label class="field grow"><span>Buscar</span><input type="search" name="q" value="${filters.q}" placeholder="Assunto, pessoa, anexo ou termo" /></label>
    <label class="field"><span>Termo</span><select name="term"><option value="">Todos</option></select></label>
    <label class="field"><span>Caixa</span><select name="mailbox"><option value="">Todas</option></select></label>
    <label class="field"><span>Remetente</span><select name="sender"><option value="">Todos</option></select></label>
    <label class="field"><span>Encontrado em</span>
      <select name="location">
        <option value="">Qualquer parte</option>
        ${Object.entries(LOCATION_TITLES).map(([value, label]) => option(value, label, filters.location))}
      </select>
    </label>
    ${deletionFilter(filters, 'a')}
    <label class="field"><span>Ordenar por</span>
      <select name="sort">
        ${[
          ['date', 'Mais recentes'],
          ['occurrences', 'Mais ocorrências'],
          ['terms', 'Mais termos'],
          ['mailbox', 'Caixa'],
          ['sender', 'Remetente'],
          ['subject', 'Assunto'],
          ['size', 'Maiores mensagens'],
        ].map(([value, label]) => option(value, label, filters.sort))}
      </select>
    </label>`,

  fillOptions: (form, options, filters, fill) => {
    const o = options || { terms: [], mailboxes: [], senders: [] };
    fill(form.elements.term, o.terms, filters.term);
    fill(form.elements.mailbox, o.mailboxes, filters.mailbox);
    const senders = new Map((o.senders || []).map((s) => [s.value, s.label]));
    fill(form.elements.sender, [...senders.keys()], filters.sender, (v) => senders.get(v) || v);
  },

  progress: (st) => html`<span><b>${fmtNum(st.messagesSeen)}</b> mensagens verificadas</span>
    <span><b>${fmtNum(st.messagesMatched)}</b> com ocorrências</span>
    <span><b>${fmtNum(st.attachmentsAnalyzed)}</b> anexos lidos</span>
    <span><b>${fmtBytes(st.bytesDownloaded)}</b> baixados</span>
    <span><b>${fmtNum(st.errors)}</b> erros</span>
    <span>${st.mailboxesTotal ? html`caixa <b>${Math.min((st.mailboxesDone || 0) + 1, st.mailboxesTotal)}</b> de <b>${fmtNum(st.mailboxesTotal)}</b>` : 'listando as caixas…'}</span>`,

  tiles: (st) => {
    const pct = st.messagesSeen ? Math.round((st.messagesMatched / st.messagesSeen) * 1000) / 10 : 0;
    const notRead = (st.attachmentsEncrypted || 0) + (st.attachmentsSkippedSize || 0) + (st.attachmentsErrors || 0);
    return html`<div class="tile"><div class="label">Mensagens verificadas</div><div class="value">${fmtCompact(st.messagesSeen)}</div><div class="detail">em ${plural(Math.max(0, (st.mailboxesDone || 0) - (st.mailboxesSkipped || 0)), 'caixa', 'caixas')}${st.mailboxesSkipped ? ` · ${fmtNum(st.mailboxesSkipped)} sem e-mail` : ''}</div></div>
      <div class="tile"><div class="label">Mensagens com ocorrências</div><div class="value">${fmtCompact(st.messagesMatched)}</div><div class="detail">${pct.toLocaleString('pt-BR')}% das verificadas</div></div>
      <div class="tile"><div class="label">Ocorrências</div><div class="value">${fmtCompact(st.occurrences)}</div><div class="detail">no assunto, no corpo e nos anexos</div></div>
      <div class="tile"><div class="label">Anexos lidos</div><div class="value">${fmtCompact(st.attachmentsAnalyzed)}</div><div class="detail">${notRead ? `${fmtNum(st.attachmentsEncrypted)} com senha · ${fmtNum(st.attachmentsSkippedSize)} grandes · ${fmtNum(st.attachmentsErrors)} com erro` : `${fmtBytes(st.bytesDownloaded)} baixados`}</div></div>
      <div class="tile"><div class="label">Erros</div><div class="value">${fmtCompact(st.errors)}</div><div class="detail">${st.errors ? 'veja a aba Erros' : st.messagesEncrypted ? `${fmtNum(st.messagesEncrypted)} mensagens criptografadas` : 'nenhum'}</div></div>`;
  },

  charts: (summary, barChart) => {
    const count = (n) => plural(n, 'mensagem', 'mensagens');
    const terms = summary.byTerm.map((t) => ({
      label: t.term,
      value: t.messages,
      filterValue: t.term,
      tipValue: `${count(t.messages)} · ${plural(t.occurrences, 'ocorrência', 'ocorrências')}`,
      tipLabel: `${t.term} — lista ${t.list}`,
      raw: t,
    }));
    const boxes = summary.byMailbox.map((m) => ({
      label: m.mailbox,
      value: m.messages,
      filterValue: m.mailbox,
      tipValue: `${count(m.messages)} · ${plural(m.occurrences, 'ocorrência', 'ocorrências')}`,
      tipLabel: m.name ? `${m.name} <${m.mailbox}>` : m.mailbox,
      raw: m,
    }));
    const senders = summary.bySender.map((s) => ({
      label: s.label,
      value: s.messages,
      filterValue: s.sender,
      tipValue: `${count(s.messages)} · ${plural(s.occurrences, 'ocorrência', 'ocorrências')}`,
      tipLabel: s.label,
      raw: s,
    }));
    const locations = Object.entries(summary.byLocation)
      .filter(([, n]) => n > 0)
      .sort((a, b) => b[1] - a[1])
      .map(([key, n]) => ({ label: LOCATION_TITLES[key], value: n, filterValue: key, tipValue: count(n), tipLabel: `Termos encontrados ${MAIL_PHRASE[key]}`, raw: { key, n } }));
    return html`${barChart({
      key: 'terms',
      title: 'Termos encontrados',
      subtitle: 'Mensagens em que cada termo aparece. Clique para filtrar.',
      rows: terms,
      filterKey: 'term',
      emptyText: 'Nenhum termo encontrado.',
      tableHead: html`<tr><th>Termo</th><th>Lista</th><th class="num">Mensagens</th><th class="num">Ocorrências</th><th class="num">No assunto</th><th class="num">No corpo</th><th class="num">Em anexos</th></tr>`,
      tableRow: (r) =>
        html`<tr><td>${r.raw.term}</td><td>${r.raw.list}</td><td class="num">${fmtNum(r.raw.messages)}</td><td class="num">${fmtNum(r.raw.occurrences)}</td><td class="num">${fmtNum(r.raw.inSubject)}</td><td class="num">${fmtNum(r.raw.inBody)}</td><td class="num">${fmtNum(r.raw.inAttachments)}</td></tr>`,
    })}
    ${barChart({
      key: 'mailboxes',
      title: 'Caixas',
      subtitle: 'Caixas que guardam mensagens com os termos. Clique para filtrar.',
      rows: boxes,
      filterKey: 'mailbox',
      emptyText: 'Nenhuma mensagem encontrada.',
      tableHead: html`<tr><th>Caixa</th><th>Nome</th><th class="num">Mensagens</th><th class="num">Ocorrências</th></tr>`,
      tableRow: (r) => html`<tr><td>${r.raw.mailbox}</td><td>${r.raw.name}</td><td class="num">${fmtNum(r.raw.messages)}</td><td class="num">${fmtNum(r.raw.occurrences)}</td></tr>`,
    })}
    ${barChart({
      key: 'senders',
      title: 'Remetentes',
      subtitle: 'Quem enviou as mensagens encontradas. Clique para filtrar.',
      rows: senders,
      filterKey: 'sender',
      emptyText: 'Nenhuma mensagem encontrada.',
      tableHead: html`<tr><th>Remetente</th><th class="num">Mensagens</th><th class="num">Ocorrências</th></tr>`,
      tableRow: (r) => html`<tr><td>${r.raw.label}</td><td class="num">${fmtNum(r.raw.messages)}</td><td class="num">${fmtNum(r.raw.occurrences)}</td></tr>`,
    })}
    ${barChart({
      key: 'locations',
      title: 'Onde foi encontrado',
      subtitle: 'Mensagens por parte em que os termos aparecem. Clique para filtrar.',
      rows: locations,
      filterKey: 'location',
      emptyText: 'Nenhuma mensagem encontrada.',
      tableHead: html`<tr><th>Parte da mensagem</th><th class="num">Mensagens</th></tr>`,
      tableRow: (r) => html`<tr><td>${r.label}</td><td class="num">${fmtNum(r.value)}</td></tr>`,
    })}`;
  },

  tableHead: html`<tr><th><span class="sr-only">Detalhes</span></th><th>Mensagem</th><th>Remetente</th><th>Data</th><th>Informação encontrada</th></tr>`,

  row: (r) => {
    const files = (r.attachments || []).filter((a) => !a.inline);
    return html`<td><div class="name">${r.subject || '(sem assunto)'}</div>${deletionChip(r, 'a')}<div class="path">${r.mailbox} › ${r.folder}${files.length ? html` · ${plural(files.length, 'anexo', 'anexos')}` : ''}</div></td>
      <td>${r.from || html`<span class="muted">sem remetente</span>`}${r.to?.length ? html`<div class="muted small">para ${r.to[0]}${r.to.length > 1 ? ` e mais ${r.to.length - 1}` : ''}</div>` : ''}</td>
      <td class="nowrap">${fmtDateTime(r.date)}</td>
      <td><div class="chips">${r.matches.map((m) => html`<span class="chip"><b>${m.term}</b> ${fmtNum(m.count)}× · ${MAIL_LOCATION[m.location] || m.location}</span>`)}</div></td>`;
  },

  rowLabel: (r) => r.subject || 'mensagem sem assunto',

  detail: (r, ctx) => {
    const safeLink = /^https:\/\//i.test(r.webLink || '') ? r.webLink : null;
    const messageId = String(r.internetMessageId || '').replace(/^<|>$/g, '');
    const people = (list) => (list?.length ? list.join('; ') : '—');
    return html`<div class="detail-grid">
      <div>
        <h4>Mensagem</h4>
        <dl class="kv">
          <dt>Assunto</dt><dd><b>${r.subject || '(sem assunto)'}</b></dd>
          <dt>Caixa</dt><dd>${r.mailboxName ? `${r.mailboxName} <${r.mailbox}>` : r.mailbox}</dd>
          <dt>Pasta</dt><dd>${r.folder}</dd>
          <dt>Recebida em</dt><dd>${fmtDateTime(r.date)}</dd>
          ${r.sent && r.sent !== r.date ? html`<dt>Enviada em</dt><dd>${fmtDateTime(r.sent)}</dd>` : ''}
          <dt>Tamanho</dt><dd>${fmtBytes(r.size)}</dd>
          <dt>Conteúdo</dt><dd>${CONTENT_STATUS[r.contentStatus] || r.contentStatus}${r.contentNote ? html`<br /><span class="muted small">${r.contentNote}</span>` : ''}</dd>
          ${messageId
            ? html`<dt>Message-ID</dt><dd><span class="mono small">${messageId}</span> <button type="button" class="btn small" data-action="copy" data-copy="${messageId}" data-copied="Message-ID copiado.">${icon('copy')} Copiar</button></dd>`
            : ''}
          <dt>Conexão</dt><dd>${r.sourceName} (${TYPE_LABELS[r.sourceType] || r.sourceType})</dd>
          ${safeLink ? html`<dt>Abrir</dt><dd><a href="${safeLink}" target="_blank" rel="noopener noreferrer">Abrir no Outlook na Web</a> <span class="muted small">(exige acesso à caixa)</span></dd>` : ''}
        </dl>
        ${deletionBlock(r, 'mensagem', ctx)}
      </div>
      <div>
        <h4>Remetente e destinatários</h4>
        <dl class="kv">
          <dt>De</dt><dd><b>${r.from || '—'}</b></dd>
          <dt>Para</dt><dd>${people(r.to)}</dd>
          ${r.cc?.length ? html`<dt>Cc</dt><dd>${people(r.cc)}</dd>` : ''}
        </dl>
        <h4 class="spaced">Anexos</h4>
        ${r.attachments?.length
          ? html`<ul class="attachment-list">
              ${r.attachments.map(
                (a) => html`<li>${icon('file')} ${a.name} <span class="muted small">· ${fmtBytes(a.size)}${a.inline ? ' · imagem no corpo' : ''}${a.status ? ` · ${CONTENT_STATUS[a.status] || a.status}` : ''}</span>${a.note ? html`<div class="muted small">${a.note}</div>` : ''}</li>`,
              )}
            </ul>`
          : html`<p class="muted small">Sem anexos.</p>`}
      </div>
      <div>
        <h4>Informação encontrada</h4>
        ${matchesHtml(r, MAIL_PHRASE)}
      </div>
    </div>`;
  },

  empty: { filtered: 'Nenhuma mensagem corresponde aos filtros.', running: 'Nenhuma ocorrência encontrada até agora.', none: 'Nenhum termo da lista foi encontrado nas mensagens analisadas.' },
  errors: {
    column: 'Local',
    help: 'Caixas, pastas ou mensagens que não puderam ser lidas (credenciais, permissões, limites do provedor...). Elas não foram analisadas.',
    empty: 'Nenhum erro de acesso ou leitura.',
  },
};

// ---------------------------------------------------------------------------------------------

export async function render(root, { params, query, isCurrent = () => true }) {
  const id = params[0];
  let scan = await get(`/api/scans/${id}`);
  if (!isCurrent()) return null; // o usuário já foi para outra tela
  const P = scan.kind === 'mail' ? MAIL : FILES;
  // Endereço e menu de acordo com o tipo da análise (ex.: link antigo para uma análise de e-mail).
  if (!location.hash.startsWith(`${P.base}/`)) history.replaceState(null, '', `${P.base}/${id}${query.toString() ? `?${query}` : ''}`);
  setActiveNav(P.nav);

  const filters = Object.fromEntries(P.filterKeys.map((k) => [k, query.get(k) || '']));
  let tab = ['arquivos', 'erros', 'registro'].includes(query.get('aba')) ? query.get('aba') : 'arquivos';
  let results = null;
  let summary = null;
  let errors = null;
  const expanded = new Set();
  const deletingNow = new Set(); // itens com exclusão manual em andamento nesta tela
  const chartView = { ...P.views };
  let stopped = false;
  let timer = null;
  let lastResults = 0;
  let loading = null;
  const stale = { results: false, errors: false };
  let lastAnnounced = '';

  const queryString = (values, extra = {}) => {
    const out = new URLSearchParams();
    for (const [key, value] of Object.entries({ ...values, ...extra })) if (value !== '' && value !== null && value !== undefined) out.set(key, value);
    if (values.sort) out.set('dir', P.descSorts.has(values.sort) ? 'desc' : 'asc');
    return out.toString();
  };

  // ---------- Estrutura fixa (os blocos abaixo são redesenhados separadamente) ----------
  paint(
    root,
    html`<div data-head></div>
      <div data-alerts></div>
      <div data-progress></div>
      <section class="tiles" data-tiles aria-label="Números da análise"></section>
      <div class="tabs" role="tablist">
        <button type="button" role="tab" data-tab="arquivos">${P.resultsTitle}</button>
        <button type="button" role="tab" data-tab="erros">Erros <span data-error-count></span></button>
        <button type="button" role="tab" data-tab="registro">Registro</button>
      </div>
      <div data-panel="arquivos">
        <form class="filters" data-filters role="search">
          ${P.filterFields(filters)}
          <button type="button" class="btn" data-action="clear-filters">Limpar filtros</button>
        </form>
        <div class="grid-2" data-charts></div>
        <section class="card" data-results></section>
      </div>
      <p class="sr-only" aria-live="polite" data-live></p>
      <div data-panel="erros" hidden><section class="card" data-errors></section></div>
      <div data-panel="registro" hidden><section class="card" data-log></section></div>`,
  );

  const $ = (sel) => root.querySelector(sel);
  const filtersForm = $('[data-filters]');

  // ---------- Cabeçalho, avisos, progresso e indicadores ----------

  const drawHead = () => {
    const started = scan.startedAt ? fmtDateTime(scan.startedAt) : '—';
    const end = scan.finishedAt ? new Date(scan.finishedAt) : new Date();
    const duration = scan.startedAt ? fmtDuration(end - new Date(scan.startedAt)) : '—';
    const qs = queryString({ ...filters, page: '' });
    const filtered = P.criteria.some((k) => filters[k]);
    const exportsLabel = filtered ? 'Exportar (com os filtros atuais):' : 'Exportar:';
    paint(
      $('[data-head]'),
      html`<div class="page-head">
        <div>
          <div class="inline"><h1>${scan.name}</h1>${statusBadge(scan.status)}${scan.options?.deleteMatches ? html`<span class="badge deleting">com exclusão automática</span>` : ''}</div>
          <div class="sub">Início ${started} · duração ${duration} · ${P.subtitle(scan)}</div>
        </div>
        <div class="actions">
          ${isActive(scan) ? html`<button type="button" class="btn danger" data-action="cancel">${icon('stop')} Cancelar análise</button>` : ''}
          <span class="muted small nowrap">${exportsLabel}</span>
          <a class="btn small" href="/api/scans/${id}/export.xlsx?${qs}" download>${icon('download')} Excel</a>
          <a class="btn small" href="/api/scans/${id}/export.csv?${qs}" download>CSV</a>
          <a class="btn small" href="/api/scans/${id}/export.html?${qs}" download>HTML</a>
          <a class="btn small" href="/api/scans/${id}/export.json?${qs}" download>JSON</a>
        </div>
      </div>`,
    );
  };

  const drawAlerts = () => {
    let content = '';
    if (scan.status === 'failed') {
      content = html`<div class="alert error">${icon('alert')}<div><b>A análise falhou.</b> ${scan.error || ''} Os resultados encontrados até a falha estão abaixo.</div></div>`;
    } else if (scan.status === 'interrupted') {
      content = html`<div class="alert">${icon('alert')}<div><b>A análise foi interrompida</b> (o servidor do CLEAN foi encerrado durante a execução). Os resultados parciais estão abaixo; inicie uma nova análise para completar.</div></div>`;
    } else if (scan.status === 'cancelled') {
      content = html`<div class="alert">${icon('info')}<div><b>Análise cancelada.</b> Os resultados encontrados até o cancelamento estão abaixo.</div></div>`;
    }
    const warnings = (scan.log || []).filter((l) => l.level === 'warn');
    if (warnings.length) {
      content = html`${content}<div class="alert">${icon('alert')}<div><b>${plural(warnings.length, 'aviso', 'avisos')} durante a análise.</b> ${warnings.at(-1).message}
          ${warnings.length > 1 ? html`<a href="#" data-action="show-log">Ver todos</a>` : ''}</div></div>`;
    }
    paint($('[data-alerts]'), content);
  };

  const drawProgress = () => {
    if (!isActive(scan)) {
      paint($('[data-progress]'), '');
      return;
    }
    paint(
      $('[data-progress]'),
      html`<section class="card progress-card">
        <div class="card-head">
          <h2>${scan.status === 'queued' ? 'Aguardando na fila…' : 'Análise em andamento'}</h2>
          <span class="muted small">Os resultados aparecem abaixo conforme são encontrados.</span>
        </div>
        <div class="progress-line" role="progressbar" aria-label="Análise em andamento"></div>
        <div class="progress-stats">${P.progress(scan.stats || {})}${scan.options?.deleteMatches ? html`<span><b>${fmtNum(scan.stats?.deleted)}</b> excluíd${P.o}s${scan.stats?.deleteErrors ? ` · ${fmtNum(scan.stats.deleteErrors)} com falha` : ''}</span>` : ''}</div>
        ${scan.current?.path ? html`<div class="current">${scan.current.path}</div>` : ''}
      </section>`,
    );
  };

  const drawTiles = () => {
    const st = scan.stats || {};
    // Excluídos: pelo registro de exclusões (automáticas e manuais) quando o resumo já chegou.
    const t = summary?.deletions || { deleted: st.deleted || 0, missing: st.deleteMissing || 0, changed: st.deleteChanged || 0, failed: st.deleteErrors || 0 };
    const o = P.o;
    const detail =
      [
        t.missing ? `${fmtNum(t.missing)} já não existia${t.missing > 1 ? 'm' : ''}` : '',
        t.changed ? `${fmtNum(t.changed)} mantid${o}${t.changed > 1 ? 's' : ''} (alterad${o}${t.changed > 1 ? 's' : ''} depois da análise)` : '',
        t.failed ? `${fmtNum(t.failed)} com falha` : '',
      ]
        .filter(Boolean)
        .join(' · ') || (scan.options?.deleteMatches ? 'exclusão automática ligada' : 'pelo relatório');
    const tile =
      scan.options?.deleteMatches || t.deleted || t.missing || t.changed || t.failed
        ? html`<div class="tile"><div class="label">Excluíd${o}s</div><div class="value">${fmtCompact(t.deleted)}</div><div class="detail">${detail}</div></div>`
        : '';
    paint($('[data-tiles]'), html`${P.tiles(st)}${tile}`);
    $('[data-error-count]').textContent = st.errors ? `(${fmtNum(st.errors)})` : '';
  };

  // ---------- Filtros ----------

  const fillSelect = (select, values, current, labelFn = (v) => v) => {
    const keep = select.querySelector('option[value=""]');
    const options = values.map((v) => {
      const el = document.createElement('option');
      el.value = v;
      el.textContent = labelFn(v);
      return el;
    });
    if (current && !values.includes(current)) {
      const el = document.createElement('option');
      el.value = current;
      el.textContent = labelFn(current);
      options.push(el);
    }
    select.replaceChildren(keep, ...options);
    select.value = current || '';
  };

  const drawFilterOptions = () => P.fillOptions(filtersForm, summary?.options, filters, fillSelect);

  // ---------- Gráficos (barras horizontais de uma série: cor única, valor na ponta) ----------

  const barChart = ({ key, title, subtitle, rows, filterKey, emptyText, tableHead, tableRow }) => {
    const view = chartView[key];
    const top = rows.slice(0, TOP);
    const max = rows[0]?.value || 0;
    const body =
      rows.length === 0
        ? html`<div class="empty">${emptyText}</div>`
        : view === 'table'
          ? html`<div class="table-wrap"><table class="data"><thead>${tableHead}</thead><tbody>${rows.map(tableRow)}</tbody></table></div>`
          : html`<div class="bars" role="list">
                ${top.map((r) => {
                  const selected = filterKey && filters[filterKey] && filters[filterKey] === r.filterValue;
                  const width = max ? Math.max(1, Math.round((r.value / max) * 88)) : 0;
                  return html`<button type="button" class="bar-row" role="listitem" ${r.filterValue ? html`data-filter-key="${filterKey}" data-filter-value="${r.filterValue}"` : html`aria-disabled="true"`}
                      aria-pressed="${selected ? 'true' : 'false'}" aria-label="${r.label}: ${r.tipValue}"
                      data-tip-value="${r.tipValue}" data-tip-label="${r.tipLabel}">
                    <span class="bar-label">${r.label}</span>
                    <span class="bar-track"><span class="bar" data-w="${width}"></span><span class="bar-value">${fmtNum(r.value)}</span></span>
                  </button>`;
                })}
              </div>
              ${rows.length > TOP ? html`<p class="bars-more">+ ${fmtNum(rows.length - TOP)} não exibidos — veja a tabela.</p>` : ''}`;
    return html`<figure class="card chart-card" data-chart="${key}">
      <figcaption>
        <div><h2>${title}</h2><p>${subtitle}</p></div>
        <div class="segmented" role="group" aria-label="Forma de exibição">
          <button type="button" data-chart-view="chart" aria-pressed="${view === 'chart' ? 'true' : 'false'}">Gráfico</button>
          <button type="button" data-chart-view="table" aria-pressed="${view === 'table' ? 'true' : 'false'}">Tabela</button>
        </div>
      </figcaption>
      ${body}
    </figure>`;
  };

  const drawCharts = () => {
    if (!summary) return;
    paint($('[data-charts]'), P.charts(summary, barChart));
  };

  // ---------- Tabela de resultados ----------

  const drawResults = () => {
    const box = $('[data-results]');
    if (!results) {
      paint(box, html`<p class="loading">Carregando resultados…</p>`);
      return;
    }
    const [one, many] = P.noun;
    const filtered = results.total !== results.totalAll;
    const heading = html`<div class="card-head">
      <h2 tabindex="-1">${plural(results.total, one, many)}${filtered ? html` <span class="muted">de ${fmtNum(results.totalAll)}</span>` : ''}</h2>
      ${isActive(scan) ? html`<span class="muted small">atualizando enquanto a análise roda…</span>` : ''}
    </div>`;
    if (results.total === 0) {
      paint(box, html`${heading}<div class="empty">${filtered ? P.empty.filtered : isActive(scan) ? P.empty.running : P.empty.none}</div>`);
      return;
    }
    paint(
      box,
      html`${heading}
        <div class="table-wrap">
          <table class="data results">
            <thead>${P.tableHead}</thead>
            <tbody>
              ${results.items.map((r) => {
                const open = expanded.has(r.id);
                const expandedText = open ? 'true' : 'false';
                return html`<tr data-id="${r.id}" aria-expanded="${expandedText}" class="${isGone(r.deletion) ? 'is-deleted' : ''}">
                    <td><button type="button" class="icon-btn" data-action="toggle" aria-label="${open ? 'Ocultar' : 'Mostrar'} detalhes de ${P.rowLabel(r)}" aria-expanded="${expandedText}"><span class="row-toggle">${icon('chevron')}</span></button></td>
                    ${P.row(r)}
                  </tr>
                  ${open ? html`<tr class="detail"><td colspan="5">${P.detail(r, { active: isActive(scan), deleting: deletingNow })}</td></tr>` : ''}`;
              })}
            </tbody>
          </table>
        </div>
        ${results.pages > 1
          ? html`<div class="pager">
              <span class="muted small">Página ${results.page} de ${results.pages}</span>
              <div class="inline">
                <button type="button" class="btn small" data-action="page" data-page="${results.page - 1}" ${results.page <= 1 ? 'disabled' : ''}>Anterior</button>
                <button type="button" class="btn small" data-action="page" data-page="${results.page + 1}" ${results.page >= results.pages ? 'disabled' : ''}>Próxima</button>
              </div>
            </div>`
          : ''}`,
    );
  };

  const drawErrors = () => {
    const box = $('[data-errors]');
    if (!errors) {
      paint(box, html`<p class="loading">Carregando…</p>`);
      return;
    }
    paint(
      box,
      errors.total === 0
        ? html`<div class="empty">${P.errors.empty}</div>`
        : html`<div class="card-head"><h2>${plural(errors.total, 'erro', 'erros')}</h2>${errors.total > errors.items.length ? html`<span class="muted small">Mostrando ${fmtNum(errors.items.length)}. Exporte para Excel para ver todos.</span>` : ''}</div>
            <p class="muted small">${P.errors.help}</p>
            <div class="table-wrap">
              <table class="data">
                <thead><tr><th>${P.errors.column}</th><th>Erro</th><th>Quando</th></tr></thead>
                <tbody>${errors.items.map((e) => html`<tr><td class="path">${e.path}</td><td>${e.message}</td><td class="nowrap">${fmtDateTime(e.time)}</td></tr>`)}</tbody>
              </table>
            </div>`,
    );
  };

  const drawLog = () => {
    const log = scan.log || [];
    const level = { info: 'Info', warn: 'Aviso', error: 'Erro' };
    paint(
      $('[data-log]'),
      log.length === 0
        ? html`<div class="empty">Sem registros.</div>`
        : html`<ul class="log">${log.map((l) => html`<li class="${l.level}"><span class="muted">${fmtDateTime(l.time)}</span><span>${level[l.level] || l.level}</span><span>${l.message}</span></li>`)}</ul>`,
    );
  };

  const drawTabs = () => {
    root.querySelectorAll('[data-tab]').forEach((b) => b.setAttribute('aria-selected', String(b.dataset.tab === tab)));
    root.querySelectorAll('[data-panel]').forEach((p) => {
      p.hidden = p.dataset.panel !== tab;
    });
  };

  // ---------- Carregamento ----------

  const syncUrl = () => replaceQuery({ ...filters, aba: tab === 'arquivos' ? '' : tab });

  /** Mensagem curta para leitores de tela (só quando muda). */
  const announce = (message) => {
    if (message === lastAnnounced) return;
    lastAnnounced = message;
    const live = root.querySelector('[data-live]');
    if (live) live.textContent = message;
  };

  const loadResults = async () => {
    lastResults = Date.now();
    const busy = [$('[data-charts]'), $('[data-results]')];
    busy.forEach((el) => el.classList.add('is-loading'));
    const request = Promise.all([
      get(`/api/scans/${id}/results?${queryString(filters, { pageSize: PAGE_SIZE })}`),
      get(`/api/scans/${id}/summary?${queryString({ ...filters, page: '' })}`),
    ]);
    loading = request;
    try {
      const [res, sum] = await request;
      if (stopped || loading !== request) return;
      results = res;
      summary = sum;
      drawTiles();
      if (results.page !== Number(filters.page || 1)) filters.page = results.page > 1 ? String(results.page) : '';
      stale.results = false;
      drawFilterOptions();
      redraw(root, () => {
        drawCharts();
        drawResults();
      });
      const [one, many] = P.noun;
      announce(`${plural(results.total, `${one} encontrad${one === 'mensagem' ? 'a' : 'o'}`, `${many} encontrad${one === 'mensagem' ? 'as' : 'os'}`)}.`);
    } catch (err) {
      if (!stopped) toast(err.message, 'error');
    } finally {
      if (loading === request) busy.forEach((el) => el.classList.remove('is-loading'));
    }
  };

  const loadErrors = async () => {
    try {
      const latest = await get(`/api/scans/${id}/errors?limit=500`);
      if (stopped) return;
      errors = latest;
      stale.errors = false;
      drawErrors();
    } catch (err) {
      if (!stopped) toast(err.message, 'error');
    }
  };

  const drawScan = () => {
    drawHead();
    drawAlerts();
    drawProgress();
    drawTiles();
    if (tab === 'registro') drawLog();
  };

  const poll = async () => {
    if (stopped) return;
    try {
      const wasActive = isActive(scan);
      const latest = await get(`/api/scans/${id}`);
      if (stopped) return;
      scan = latest;
      drawScan();
      const finished = wasActive && !isActive(scan);
      if (finished) {
        // As abas não visíveis serão recarregadas ao serem abertas.
        stale.results = true;
        stale.errors = true;
        announce(`Análise ${scan.status === 'completed' ? 'concluída' : 'encerrada'}.`);
        toast(scan.status === 'completed' ? 'Análise concluída.' : 'A análise foi encerrada.', scan.status === 'completed' ? 'success' : 'info');
      } else if (isActive(scan)) {
        stale.results = true;
        stale.errors = true;
      }
      if (tab === 'arquivos' && (finished || Date.now() - lastResults > 4000)) await loadResults();
      if (tab === 'erros' && (finished || scan.stats?.errors !== errors?.total)) await loadErrors();
    } catch {
      // tenta de novo no próximo ciclo
    }
    if (!stopped && isActive(scan)) timer = setTimeout(poll, 2000);
  };

  const applyFilters = (changes) => {
    if (stopped) return;
    Object.assign(filters, changes);
    if (!('page' in changes)) filters.page = '';
    syncUrl();
    drawHead();
    loadResults();
  };

  // ---------- Eventos ----------

  const onClick = async (event) => {
    const tabButton = event.target.closest('[data-tab]');
    if (tabButton) {
      tab = tabButton.dataset.tab;
      drawTabs();
      syncUrl();
      if (tab === 'erros' && (!errors || stale.errors)) loadErrors();
      if (tab === 'registro') drawLog();
      if (tab === 'arquivos' && (!results || stale.results)) loadResults();
      return;
    }
    const viewButton = event.target.closest('[data-chart-view]');
    if (viewButton) {
      chartView[viewButton.closest('[data-chart]').dataset.chart] = viewButton.dataset.chartView;
      redraw(root, drawCharts);
      return;
    }
    const bar = event.target.closest('.bar-row[data-filter-key]');
    if (bar) {
      const key = bar.dataset.filterKey;
      const value = bar.dataset.filterValue;
      const changes = { [key]: filters[key] === value ? '' : value };
      if (key === 'location' && filtersForm.elements.location) filtersForm.elements.location.value = changes.location;
      applyFilters(changes);
      document.getElementById('tooltip').hidden = true;
      return;
    }
    const el = event.target.closest('[data-action]');
    if (!el) return;
    const action = el.dataset.action;
    if (action === 'toggle') {
      const rid = Number(el.closest('tr').dataset.id);
      if (expanded.has(rid)) expanded.delete(rid);
      else expanded.add(rid);
      redraw(root, drawResults);
    } else if (action === 'page') {
      applyFilters({ page: el.dataset.page });
      $('[data-results]').scrollIntoView({ block: 'start', behavior: 'smooth' });
    } else if (action === 'clear-filters') {
      filtersForm.elements.q.value = '';
      filtersForm.elements.location.value = '';
      filtersForm.elements.deletion.value = '';
      filtersForm.elements.sort.value = P.defaultSort;
      applyFilters({ ...Object.fromEntries(P.criteria.map((k) => [k, ''])), sort: '' });
    } else if (action === 'copy') {
      try {
        await copyText(el.dataset.copy);
        toast(el.dataset.copied || 'Copiado.', 'success');
      } catch {
        toast('Não foi possível copiar. Selecione o texto e copie manualmente.', 'error');
      }
    } else if (action === 'delete-item') {
      const rid = Number(el.dataset.rid);
      const record = results?.items.find((r) => r.id === rid);
      if (!record || deletingNow.has(rid)) return;
      const mail = scan.kind === 'mail';
      const what = mail ? `a mensagem "${record.subject || '(sem assunto)'}" da caixa ${record.mailbox}` : `o arquivo ${record.path}`;
      const how =
        record.deleteMethod === 'trash'
          ? `Ela será movida para ${record.sourceType === 'graph' ? 'a pasta Itens Excluídos' : 'a Lixeira'} da caixa.`
          : record.deleteMethod === 'permanent'
            ? 'A exclusão é definitiva: a mensagem não fica na lixeira do usuário.'
            : 'A exclusão é definitiva: o arquivo não vai para a Lixeira.';
      if (!(await confirmDialog(`Excluir ${what}? ${how}`, { title: mail ? 'Excluir mensagem' : 'Excluir arquivo', confirmLabel: 'Excluir' }))) return;
      if (deletingNow.has(rid) || stopped) return;
      // Enquanto o pedido não termina, o botão fica "Excluindo…" (também se a tabela for redesenhada).
      deletingNow.add(rid);
      redraw(root, drawResults);
      // A forma mostrada na confirmação vai junto: se o cadastro mudou, o servidor recusa (409).
      const send = (force) => post(`/api/scans/${id}/results/${rid}/delete`, { confirm: true, force, method: record.deleteMethod });
      try {
        let res;
        try {
          res = await send(false);
        } catch (err) {
          if (err.code !== 'changed') throw err;
          if (!(await confirmDialog(err.message, { title: 'Arquivo alterado depois da análise', confirmLabel: 'Excluir mesmo assim' }))) return;
          res = await send(true);
        }
        const d = res.deletion;
        const noun = mail ? 'A mensagem' : 'O arquivo';
        if (d.status === 'failed') toast(`Não foi possível excluir: ${d.error}`, 'error');
        else if (d.status === 'missing') toast(`${noun} já não existia (registrad${mail ? 'a' : 'o'} como não encontrad${mail ? 'a' : 'o'}).`, 'success');
        else toast(mail ? 'Mensagem excluída.' : 'Arquivo excluído.', 'success');
      } catch (err) {
        toast(err.status === 409 && err.code === 'method-changed' ? `${err.message} A tela foi atualizada.` : err.message, 'error');
      } finally {
        deletingNow.delete(rid);
        // Recarrega sempre: o item pode ter sido excluído em outra aba ou a permissão pode ter mudado.
        try {
          scan = await get(`/api/scans/${id}`);
          if (!stopped) drawScan();
        } catch {
          // mantém a tela como está
        }
        if (!stopped) {
          await loadResults();
          // O foco volta para a linha do item (ou para o título da lista, se ela saiu do filtro).
          const toggle = root.querySelector(`tr[data-id="${rid}"] [data-action="toggle"]`);
          (toggle || root.querySelector('[data-results] h2'))?.focus();
        }
      }
    } else if (action === 'cancel') {
      if (!(await confirmDialog('Cancelar esta análise? Os resultados encontrados até agora serão mantidos.', { confirmLabel: 'Cancelar análise' }))) return;
      try {
        scan = await post(`/api/scans/${id}/cancel`);
        toast('Cancelamento solicitado.');
        drawScan();
      } catch (err) {
        toast(err.message, 'error');
      }
    } else if (action === 'show-log') {
      event.preventDefault();
      tab = 'registro';
      drawTabs();
      drawLog();
      syncUrl();
    }
  };

  const onFilterChange = (event) => {
    const { name, value } = event.target;
    if (name && name !== 'q' && P.filterKeys.includes(name)) applyFilters({ [name]: value });
  };
  const onSearch = debounce((value) => {
    if (!stopped) applyFilters({ q: value.trim() });
  }, 350);
  const onInput = (event) => {
    if (event.target.name === 'q') onSearch(event.target.value);
  };
  const onSubmit = (event) => event.preventDefault();

  root.addEventListener('click', onClick);
  filtersForm.addEventListener('change', onFilterChange);
  filtersForm.addEventListener('input', onInput);
  filtersForm.addEventListener('submit', onSubmit);
  bindTooltips($('[data-charts]'));

  drawTabs();
  drawScan();
  drawResults();
  if (tab === 'arquivos') await loadResults();
  if (tab === 'erros') await loadErrors();
  if (isActive(scan)) timer = setTimeout(poll, 2000);

  return () => {
    stopped = true;
    onSearch.cancel();
    clearTimeout(timer);
    root.removeEventListener('click', onClick);
  };
}
