// Relatório de uma análise: progresso, indicadores, gráficos, filtros, arquivos encontrados e exportações.
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
} from '../ui.js';
import { replaceQuery } from '../nav.js';

const SOURCE = { audit: 'Log de auditoria', metadata: 'Metadados do documento', owner: 'Proprietário do arquivo (NTFS)' };
const SOURCE_SHORT = { audit: 'auditoria', metadata: 'metadados', owner: 'proprietário' };
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
const LOCATION = { name: 'nome', content: 'conteúdo' };
const FILTER_KEYS = ['q', 'term', 'user', 'location', 'extension', 'sort', 'page'];
const DESC_SORTS = new Set(['modified', 'occurrences', 'terms', 'size']);
const TOP = 10;
const PAGE_SIZE = 50;

const isActive = (scan) => scan.status === 'running' || scan.status === 'queued';

function folderOf(record) {
  const rel = record.relativePath || '';
  const idx = Math.max(rel.lastIndexOf('\\'), rel.lastIndexOf('/'));
  return idx === -1 ? record.repositoryName : `${record.repositoryName} › ${rel.slice(0, idx)}`;
}

function queryString(filters, extra = {}) {
  const params = new URLSearchParams();
  for (const [key, value] of Object.entries({ ...filters, ...extra })) if (value !== '' && value !== null && value !== undefined) params.set(key, value);
  if (filters.sort) params.set('dir', DESC_SORTS.has(filters.sort) ? 'desc' : 'asc');
  return params.toString();
}

export async function render(root, { params, query }) {
  const id = params[0];
  let scan = await get(`/api/scans/${id}`);
  const filters = Object.fromEntries(FILTER_KEYS.map((k) => [k, query.get(k) || '']));
  let tab = ['arquivos', 'erros', 'registro'].includes(query.get('aba')) ? query.get('aba') : 'arquivos';
  let results = null;
  let summary = null;
  let errors = null;
  const expanded = new Set();
  const chartView = { terms: 'chart', users: 'chart' };
  let stopped = false;
  let timer = null;
  let lastResults = 0;
  let loading = null;

  // ---------- Estrutura fixa (os blocos abaixo são redesenhados separadamente) ----------
  paint(
    root,
    html`<div data-head></div>
      <div data-alerts></div>
      <div data-progress></div>
      <section class="tiles" data-tiles aria-label="Números da análise"></section>
      <div class="tabs" role="tablist">
        <button type="button" role="tab" data-tab="arquivos">Arquivos com ocorrências</button>
        <button type="button" role="tab" data-tab="erros">Erros <span data-error-count></span></button>
        <button type="button" role="tab" data-tab="registro">Registro</button>
      </div>
      <div data-panel="arquivos">
        <form class="filters" data-filters role="search">
          <label class="field grow"><span>Buscar</span><input type="search" name="q" value="${filters.q}" placeholder="Caminho, usuário ou termo" /></label>
          <label class="field"><span>Termo</span><select name="term"><option value="">Todos</option></select></label>
          <label class="field"><span>Último usuário</span><select name="user"><option value="">Todos</option></select></label>
          <label class="field"><span>Encontrado em</span>
            <select name="location">
              <option value="">Nome ou conteúdo</option>
              <option value="name" ${filters.location === 'name' ? 'selected' : ''}>Nome</option>
              <option value="content" ${filters.location === 'content' ? 'selected' : ''}>Conteúdo</option>
            </select>
          </label>
          <label class="field"><span>Extensão</span><select name="extension"><option value="">Todas</option></select></label>
          <label class="field"><span>Ordenar por</span>
            <select name="sort">
              ${[
                ['path', 'Caminho'],
                ['occurrences', 'Mais ocorrências'],
                ['terms', 'Mais termos'],
                ['modified', 'Modificados recentemente'],
                ['lastUser', 'Último usuário'],
                ['size', 'Maiores arquivos'],
              ].map(([value, label]) => html`<option value="${value}" ${filters.sort === value ? 'selected' : ''}>${label}</option>`)}
            </select>
          </label>
          <button type="button" class="btn" data-action="clear-filters">Limpar filtros</button>
        </form>
        <div class="grid-2" data-charts></div>
        <section class="card" data-results aria-live="polite"></section>
      </div>
      <div data-panel="erros" hidden><section class="card" data-errors></section></div>
      <div data-panel="registro" hidden><section class="card" data-log></section></div>`,
  );

  const $ = (sel) => root.querySelector(sel);
  const filtersForm = $('[data-filters]');

  // ---------- Cabeçalho, avisos, progresso e indicadores ----------

  const drawHead = () => {
    const s = scan.summary || {};
    const started = scan.startedAt ? fmtDateTime(scan.startedAt) : '—';
    const end = scan.finishedAt ? new Date(scan.finishedAt) : new Date();
    const duration = scan.startedAt ? fmtDuration(end - new Date(scan.startedAt)) : '—';
    const qs = queryString({ ...filters, page: '' });
    const filtered = ['q', 'term', 'user', 'location', 'extension'].some((k) => filters[k]);
    const exportsLabel = filtered ? 'Exportar (com os filtros atuais):' : 'Exportar:';
    paint(
      $('[data-head]'),
      html`<div class="page-head">
        <div>
          <div class="inline"><h1>${scan.name}</h1>${statusBadge(scan.status)}</div>
          <div class="sub">
            Início ${started} · duração ${duration} ·
            ${(s.repositories || []).map((r) => r.name).join(', ')} · ${(s.lists || []).map((l) => `${l.name} (${fmtNum(l.termCount)})`).join(', ')}
          </div>
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
    const st = scan.stats || {};
    paint(
      $('[data-progress]'),
      html`<section class="card progress-card" aria-live="polite">
        <div class="card-head">
          <h2>${scan.status === 'queued' ? 'Aguardando na fila…' : 'Análise em andamento'}</h2>
          <span class="muted small">Os resultados aparecem abaixo conforme são encontrados.</span>
        </div>
        <div class="progress-line" role="progressbar" aria-label="Análise em andamento"></div>
        <div class="progress-stats">
          <span><b>${fmtNum(st.filesSeen)}</b> arquivos verificados</span>
          <span><b>${fmtNum(st.directories)}</b> pastas</span>
          <span><b>${fmtNum(st.filesMatched)}</b> com ocorrências</span>
          <span><b>${fmtBytes(st.bytesAnalyzed)}</b> de conteúdo lido</span>
          <span><b>${fmtNum(st.errors)}</b> erros</span>
          <span>repositório <b>${Math.min((st.repositoriesDone || 0) + 1, st.repositoriesTotal || 1)}</b> de <b>${st.repositoriesTotal || 1}</b></span>
        </div>
        ${scan.current?.path ? html`<div class="current">${scan.current.path}</div>` : ''}
      </section>`,
    );
  };

  const drawTiles = () => {
    const st = scan.stats || {};
    const pct = st.filesSeen ? Math.round((st.filesMatched / st.filesSeen) * 1000) / 10 : 0;
    const notRead = (st.contentEncrypted || 0) + (st.contentSkippedSize || 0) + (st.contentErrors || 0);
    paint(
      $('[data-tiles]'),
      html`<div class="tile"><div class="label">Arquivos verificados</div><div class="value">${fmtCompact(st.filesSeen)}</div><div class="detail">em ${plural(st.directories || 0, 'pasta', 'pastas')}${st.filesSkippedByDate ? ` · ${fmtNum(st.filesSkippedByDate)} fora do período` : ''}</div></div>
        <div class="tile"><div class="label">Arquivos com ocorrências</div><div class="value">${fmtCompact(st.filesMatched)}</div><div class="detail">${pct.toLocaleString('pt-BR')}% dos verificados</div></div>
        <div class="tile"><div class="label">Ocorrências</div><div class="value">${fmtCompact(st.occurrences)}</div><div class="detail">somando nome e conteúdo</div></div>
        <div class="tile"><div class="label">Conteúdos lidos</div><div class="value">${fmtCompact(st.contentAnalyzed)}</div><div class="detail">${notRead ? `${fmtNum(st.contentEncrypted)} com senha · ${fmtNum(st.contentSkippedSize)} grandes · ${fmtNum(st.contentErrors)} com erro` : fmtBytes(st.bytesAnalyzed)}</div></div>
        <div class="tile"><div class="label">Erros de acesso ou leitura</div><div class="value">${fmtCompact(st.errors)}</div><div class="detail">${st.errors ? 'veja a aba Erros' : 'nenhum'}</div></div>`,
    );
    $('[data-error-count]').textContent = st.errors ? `(${fmtNum(st.errors)})` : '';
  };

  // ---------- Filtros ----------

  const fillSelect = (select, values, current, labelFn = (v) => v) => {
    const keep = select.querySelector('option[value=""]');
    const options = values.map((v) => {
      const option = document.createElement('option');
      option.value = v;
      option.textContent = labelFn(v);
      return option;
    });
    if (current && !values.includes(current)) {
      const option = document.createElement('option');
      option.value = current;
      option.textContent = labelFn(current);
      options.push(option);
    }
    select.replaceChildren(keep, ...options);
    select.value = current || '';
  };

  const drawFilterOptions = () => {
    const o = summary?.options || { terms: [], users: [], extensions: [] };
    fillSelect(filtersForm.elements.term, o.terms, filters.term);
    fillSelect(filtersForm.elements.user, o.users, filters.user);
    fillSelect(filtersForm.elements.extension, o.extensions, filters.extension);
  };

  // ---------- Gráficos (barras horizontais de uma série: cor única, valor na ponta) ----------

  const barChart = ({ key, title, subtitle, rows, max, filterKey, emptyText, tableHead, tableRow }) => {
    const view = chartView[key];
    const top = rows.slice(0, TOP);
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
          <button type="button" data-chart-view="chart" aria-pressed="${view === 'chart'}">Gráfico</button>
          <button type="button" data-chart-view="table" aria-pressed="${view === 'table'}">Tabela</button>
        </div>
      </figcaption>
      ${body}
    </figure>`;
  };

  const drawCharts = () => {
    if (!summary) return;
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
    paint(
      $('[data-charts]'),
      html`${barChart({
        key: 'terms',
        title: 'Termos encontrados',
        subtitle: 'Arquivos em que cada termo aparece. Clique para filtrar.',
        rows: terms,
        max: terms[0]?.value || 0,
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
        max: users[0]?.value || 0,
        filterKey: 'user',
        emptyText: 'Nenhum arquivo encontrado.',
        tableHead: html`<tr><th>Usuário</th><th>Fontes</th><th class="num">Arquivos</th><th class="num">Ocorrências</th></tr>`,
        tableRow: (r) => html`<tr><td>${r.raw.user}</td><td class="small">${Object.entries(r.raw.sources).map(([k, n]) => `${SOURCE_SHORT[k]}: ${n}`).join(', ')}</td><td class="num">${fmtNum(r.raw.files)}</td><td class="num">${fmtNum(r.raw.occurrences)}</td></tr>`,
      })}`,
    );
  };

  // ---------- Tabela de resultados ----------

  const sampleHtml = (s) => html`<div class="sample">${s.where ? html`<span class="where">${s.where}</span>` : ''}${s.before}<mark>${s.match}</mark>${s.after}</div>`;

  const detail = (r) => {
    const m = r.metadata || {};
    const a = r.audit;
    return html`<div class="detail-grid">
      <div>
        <h4>Arquivo</h4>
        <dl class="kv">
          <dt>Caminho</dt>
          <dd><span class="mono">${r.path}</span> <button type="button" class="btn small" data-action="copy" data-path="${r.path}">${icon('copy')} Copiar</button></dd>
          <dt>Tamanho</dt><dd>${fmtBytes(r.size)}</dd>
          <dt>Criado em</dt><dd>${fmtDateTime(r.created)}</dd>
          <dt>Modificado em</dt><dd>${fmtDateTime(r.modified)}</dd>
          <dt>Tipo</dt><dd>${(r.contentType || r.extension || '—').toString().toUpperCase()} · ${CONTENT_STATUS[r.contentStatus] || r.contentStatus || '—'}</dd>
          ${r.contentNote ? html`<dt>Observação</dt><dd>${r.contentNote}</dd>` : ''}
          ${m.title ? html`<dt>Título</dt><dd>${m.title}</dd>` : ''}
        </dl>
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
        ${r.matches.map(
          (mt) => html`<div class="match">
            <div class="match-head">
              <span class="chip"><b>${mt.term}</b></span>
              <span class="muted small">lista ${mt.list} · no ${LOCATION[mt.location]} · ${plural(mt.count, 'ocorrência', 'ocorrências')}${mt.truncated ? '+' : ''}</span>
            </div>
            ${mt.values?.length && mt.kind === 'regex' ? html`<div class="small"><span class="muted">Valores:</span> ${mt.values.join(' · ')}</div>` : ''}
            ${mt.samples.map(sampleHtml)}
          </div>`,
        )}
      </div>
    </div>`;
  };

  const drawResults = () => {
    const box = $('[data-results]');
    if (!results) {
      paint(box, html`<p class="loading">Carregando resultados…</p>`);
      return;
    }
    const filtered = results.total !== results.totalAll;
    const heading = html`<div class="card-head">
      <h2>${plural(results.total, 'arquivo', 'arquivos')}${filtered ? html` <span class="muted">de ${fmtNum(results.totalAll)}</span>` : ''}</h2>
      ${isActive(scan) ? html`<span class="muted small">atualizando enquanto a análise roda…</span>` : ''}
    </div>`;
    if (results.total === 0) {
      paint(
        box,
        html`${heading}<div class="empty">${filtered ? 'Nenhum arquivo corresponde aos filtros.' : isActive(scan) ? 'Nenhuma ocorrência encontrada até agora.' : 'Nenhum termo da lista foi encontrado nos arquivos analisados.'}</div>`,
      );
      return;
    }
    paint(
      box,
      html`${heading}
        <div class="table-wrap">
          <table class="data results">
            <thead><tr><th><span class="sr-only">Detalhes</span></th><th>Arquivo</th><th>Último usuário</th><th>Modificado em</th><th>Informação encontrada</th></tr></thead>
            <tbody>
              ${results.items.map((r) => {
                const open = expanded.has(r.id);
                return html`<tr data-id="${r.id}" aria-expanded="${open}">
                    <td><button type="button" class="icon-btn" data-action="toggle" aria-label="${open ? 'Ocultar' : 'Mostrar'} detalhes de ${r.name}" aria-expanded="${open}"><span class="row-toggle">${icon('chevron')}</span></button></td>
                    <td><div class="name">${r.name}</div><div class="path">${folderOf(r)}</div></td>
                    <td>${r.lastUser ? html`${r.lastUser}<div><span class="chip source">${SOURCE_SHORT[r.lastUserSource]}</span></div>` : html`<span class="muted">não identificado</span>`}</td>
                    <td class="nowrap">${fmtDateTime(r.modified)}</td>
                    <td><div class="chips">${r.matches.map((m) => html`<span class="chip"><b>${m.term}</b> ${fmtNum(m.count)}× · ${LOCATION[m.location]}</span>`)}</div></td>
                  </tr>
                  ${open ? html`<tr class="detail"><td colspan="5">${detail(r)}</td></tr>` : ''}`;
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
        ? html`<div class="empty">Nenhum erro de acesso ou leitura.</div>`
        : html`<div class="card-head"><h2>${plural(errors.total, 'erro', 'erros')}</h2>${errors.total > errors.items.length ? html`<span class="muted small">Mostrando ${fmtNum(errors.items.length)}. Exporte para Excel para ver todos.</span>` : ''}</div>
            <p class="muted small">Pastas ou arquivos que a conta do CLEAN não conseguiu abrir (permissão, arquivo em uso, caminho longo...). Eles não foram analisados.</p>
            <div class="table-wrap">
              <table class="data">
                <thead><tr><th>Caminho</th><th>Erro</th><th>Quando</th></tr></thead>
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
      if (results.page !== Number(filters.page || 1)) filters.page = results.page > 1 ? String(results.page) : '';
      drawFilterOptions();
      drawCharts();
      drawResults();
    } catch (err) {
      if (!stopped) toast(err.message, 'error');
    } finally {
      if (loading === request) busy.forEach((el) => el.classList.remove('is-loading'));
    }
  };

  const loadErrors = async () => {
    try {
      errors = await get(`/api/scans/${id}/errors?limit=500`);
      if (!stopped) drawErrors();
    } catch (err) {
      toast(err.message, 'error');
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
      if (tab === 'arquivos' && (finished || Date.now() - lastResults > 4000)) await loadResults();
      if (tab === 'erros' && (finished || scan.stats?.errors !== errors?.total)) await loadErrors();
    } catch {
      // tenta de novo no próximo ciclo
    }
    if (!stopped && isActive(scan)) timer = setTimeout(poll, 2000);
  };

  const applyFilters = (changes) => {
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
      if (tab === 'erros' && !errors) loadErrors();
      if (tab === 'registro') drawLog();
      if (tab === 'arquivos' && !results) loadResults();
      return;
    }
    const viewButton = event.target.closest('[data-chart-view]');
    if (viewButton) {
      chartView[viewButton.closest('[data-chart]').dataset.chart] = viewButton.dataset.chartView;
      drawCharts();
      return;
    }
    const bar = event.target.closest('.bar-row[data-filter-key]');
    if (bar) {
      const key = bar.dataset.filterKey;
      const value = bar.dataset.filterValue;
      applyFilters({ [key]: filters[key] === value ? '' : value });
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
      drawResults();
      root.querySelector(`tr[data-id="${rid}"] [data-action="toggle"]`)?.focus();
    } else if (action === 'page') {
      applyFilters({ page: el.dataset.page });
      $('[data-results]').scrollIntoView({ block: 'start', behavior: 'smooth' });
    } else if (action === 'clear-filters') {
      filtersForm.elements.q.value = '';
      filtersForm.elements.location.value = '';
      filtersForm.elements.sort.value = 'path';
      applyFilters({ q: '', term: '', user: '', location: '', extension: '', sort: '' });
    } else if (action === 'copy') {
      try {
        await copyText(el.dataset.path);
        toast('Caminho copiado.', 'success');
      } catch {
        toast('Não foi possível copiar. Selecione o caminho e copie manualmente.', 'error');
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
    if (name && name !== 'q' && FILTER_KEYS.includes(name)) applyFilters({ [name]: value });
  };
  const onSearch = debounce((value) => applyFilters({ q: value.trim() }), 350);
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
    clearTimeout(timer);
    root.removeEventListener('click', onClick);
  };
}
