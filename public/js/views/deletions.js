// Registro de exclusões: todas as exclusões feitas pelo CLEAN — automáticas (políticas de retenção,
// buscas por tipo e análises com exclusão automática), manuais pelo relatório e em lote —, inclusive
// as de relatórios que já foram excluídos. Filtros na URL, resumo, lista paginada e exportações.
import { get } from '../api.js';
import { html, render as paint, icon, fmtNum, fmtCompact, fmtBytes, fmtServerDateTime, redraw, debounce, plural } from '../ui.js';
import { replaceQuery } from '../nav.js';

const PAGE_SIZE = 50;
const TOP = 8; // linhas das tabelas "Por política" e "Por repositório"

// Parâmetros da URL (#/exclusoes?de=...) -> filtros da API. "finalidade" e "analise" não têm campo
// na tela (vêm de links, ex.: do relatório de uma análise) e aparecem como etiquetas removíveis.
const PARAMS = {
  de: 'from',
  ate: 'to',
  politica: 'scheduleId',
  repositorio: 'repository',
  situacao: 'status',
  origem: 'mode',
  tipo: 'kind',
  finalidade: 'purpose',
  analise: 'scanId',
  q: 'q',
};

const STATUS_OPTIONS = [
  ['deleted', 'Excluídos'],
  ['missing', 'Não encontrados'],
  ['changed,failed', 'Falhas ou alterados'],
  ['failed', 'Só as falhas'],
  ['changed', 'Só os alterados (mantidos)'],
];
const ORIGIN_OPTIONS = [
  ['retention', 'Automática (política de retenção)'],
  ['auto', 'Automática (análise ou busca por tipo)'],
  ['manual', 'Manual (relatório ou em lote)'],
];
const KIND_OPTIONS = [
  ['local', 'Pastas do Windows'],
  ['cloud', 'OneDrive e SharePoint'],
  ['onedrive', 'OneDrive'],
  ['sharepoint', 'SharePoint'],
  ['mail', 'E-mail'],
  ['teams', 'Teams'],
];
const PURPOSE_TEXT = { retention: 'Políticas de retenção', types: 'Buscas por tipo', terms: 'Análises por termos' };
const NO_POLICY = 'Sem política (análise avulsa ou manual)';

/** Situação em poucas palavras (o texto completo fica na dica): [classe, arquivo, mensagem]. */
const STATUS_BADGE = {
  deleted: ['completed', 'Excluído', 'Excluída'],
  missing: ['', 'Não encontrado', 'Não encontrada'],
  changed: ['cancelled', 'Mantido (alterado)', 'Mantida (alterada)'],
  failed: ['failed', 'Falha', 'Falha'],
};
const REPORT_BASE = { files: '#/analises', mail: '#/email/analises', teams: '#/teams/analises' };

const option = (value, label, current) => html`<option value="${value}" ${current === value ? 'selected' : ''}>${label}</option>`;

function statusBadge(n) {
  const [cls, file, message] = STATUS_BADGE[n.status] || ['', n.statusLabel, n.statusLabel];
  return html`<span class="badge ${cls}" title="${n.statusLabel}">${n.kind === 'mail' || n.kind === 'teams' ? message : file}</span>`;
}

/** A análise do item: link para o relatório ou o nome (relatório excluído). */
function scanLink(n) {
  if (!n.scanName && !n.scanId) return '';
  if (n.scanExists && REPORT_BASE[n.scanKind]) return html`<a href="${REPORT_BASE[n.scanKind]}/${n.scanId}">${n.scanName || 'abrir o relatório'}</a>`;
  return html`${n.scanName || n.scanId} <span class="muted">(relatório excluído)</span>`;
}

export async function render(root, { query, isCurrent = () => true }) {
  const filters = Object.fromEntries(Object.keys(PARAMS).map((key) => [key, query.get(key) || '']));
  filters.page = query.get('page') || '';
  let data = null; // página atual
  let summary = null; // resumo e opções dos filtros
  let failure = ''; // erro da última consulta (ex.: período inválido)
  let loading = null;
  let stopped = false;
  let lastAnnounced = '';

  const criteria = () => Object.keys(PARAMS).filter((key) => filters[key]);

  /** Filtros da tela -> parâmetros da API (extra: página e tamanho). */
  const apiQuery = (extra = {}) => {
    const out = new URLSearchParams();
    for (const [param, key] of Object.entries(PARAMS)) if (filters[param]) out.set(key, filters[param]);
    for (const [key, value] of Object.entries(extra)) if (value) out.set(key, value);
    return out.toString();
  };

  paint(
    root,
    html`<div class="page-head">
        <div>
          <h1>Registro de exclusões</h1>
          <div class="sub">Toda exclusão feita pelo CLEAN — automática (políticas de retenção, buscas por tipo e análises com exclusão automática), manual pelo relatório ou em lote — fica registrada aqui, mesmo depois que o relatório da análise é excluído.</div>
        </div>
        <div class="actions" data-exports></div>
      </div>
      <form class="filters" data-filters role="search">
        <label class="field"><span>De</span><input type="date" name="de" value="${filters.de}" /></label>
        <label class="field"><span>Até</span><input type="date" name="ate" value="${filters.ate}" /></label>
        <label class="field"><span>Política/agendamento</span><select name="politica"><option value="">Todas</option></select></label>
        <label class="field"><span>Repositório/conexão</span><select name="repositorio"><option value="">Todos</option></select></label>
        <label class="field"><span>Situação</span>
          <select name="situacao"><option value="">Todas</option>${STATUS_OPTIONS.map(([v, l]) => option(v, l, filters.situacao))}</select>
        </label>
        <label class="field"><span>Origem</span>
          <select name="origem"><option value="">Todas</option>${ORIGIN_OPTIONS.map(([v, l]) => option(v, l, filters.origem))}</select>
        </label>
        <label class="field"><span>Tipo</span>
          <select name="tipo"><option value="">Todos</option>${KIND_OPTIONS.map(([v, l]) => option(v, l, filters.tipo))}</select>
        </label>
        <label class="field grow"><span>Buscar</span><input type="search" name="q" value="${filters.q}" maxlength="500" placeholder="Arquivo, assunto, análise, política ou quem excluiu" /></label>
        <button type="button" class="btn" data-action="clear-filters">Limpar filtros</button>
      </form>
      <div data-scope></div>
      <section class="tiles" data-tiles aria-label="Números do registro de exclusões"></section>
      <div class="grid-2 deletions-groups" data-groups></div>
      <section class="card" data-results><p class="loading">Carregando…</p></section>
      <p class="sr-only" aria-live="polite" data-live></p>`,
  );
  const $ = (sel) => root.querySelector(sel);
  const form = $('[data-filters]');

  const announce = (message) => {
    if (message === lastAnnounced) return;
    lastAnnounced = message;
    $('[data-live]').textContent = message;
  };

  const syncUrl = () => replaceQuery(filters);

  // ---------- Desenho ----------

  const drawExports = () => {
    const qs = apiQuery();
    const suffix = qs ? `?${qs}` : '';
    paint(
      $('[data-exports]'),
      html`<span class="muted small nowrap">${criteria().length ? 'Exportar (com os filtros atuais):' : 'Exportar:'}</span>
        <a class="btn small" href="/api/deletions/export.xlsx${suffix}" download>${icon('download')} Excel</a>
        <a class="btn small" href="/api/deletions/export.csv${suffix}" download>CSV</a>
        <a class="btn small" href="/api/deletions/export.html${suffix}" download>HTML</a>`,
    );
  };

  /** Preenche uma lista de opções mantendo o valor atual (mesmo que ele não esteja mais no registro). */
  const fillSelect = (select, content, current, currentLabel) => {
    paint(select, content);
    if (current && ![...select.options].some((o) => o.value === current)) {
      const el = document.createElement('option');
      el.value = current;
      el.textContent = currentLabel || current;
      select.append(el);
    }
    select.value = current || '';
  };

  const drawOptions = () => {
    const o = summary?.options;
    if (!o) return;
    const groups = Object.entries(PURPOSE_TEXT)
      .map(([purpose, label]) => [label, o.schedules.filter((s) => (s.purpose || 'terms') === purpose)])
      .filter(([, list]) => list.length);
    const others = o.schedules.filter((s) => !PURPOSE_TEXT[s.purpose || 'terms']);
    const scheduleText = (s) => `${s.name || '(sem nome)'}${s.exists ? '' : s.purpose === 'retention' ? ' (excluída)' : ' (excluído)'}`;
    fillSelect(
      form.elements.politica,
      html`<option value="">Todas</option>
        <option value="none">${NO_POLICY}</option>
        ${groups.map(([label, list]) => html`<optgroup label="${label}">${list.map((s) => html`<option value="${s.id}">${scheduleText(s)}</option>`)}</optgroup>`)}
        ${others.map((s) => html`<option value="${s.id}">${scheduleText(s)}</option>`)}`,
      filters.politica,
    );
    fillSelect(
      form.elements.repositorio,
      html`<option value="">Todos</option>${o.repositories.map((r) => html`<option value="${r.id}">${r.name} (${r.kindLabel})</option>`)}`,
      filters.repositorio,
    );
  };

  // Filtros sem campo na tela (vindos de um link): etiquetas para removê-los.
  const drawScope = () => {
    const chips = [];
    if (filters.analise) {
      const name = data?.items?.find((n) => n.scanId === filters.analise)?.scanName || filters.analise;
      chips.push(['analise', `Análise: ${name}`]);
    }
    if (filters.finalidade) chips.push(['finalidade', `Finalidade: ${PURPOSE_TEXT[filters.finalidade] || filters.finalidade}`]);
    paint(
      $('[data-scope]'),
      chips.length
        ? html`<div class="deletions-scope chips">${chips.map(
            ([key, label]) => html`<span class="chip">${label} <button type="button" class="icon-btn" data-action="remove-filter" data-filter="${key}" aria-label="Remover o filtro ${label}" title="Remover o filtro">${icon('x')}</button></span>`,
          )}</div>`
        : '',
    );
  };

  const drawTiles = () => {
    const t = summary?.totals;
    if (!t) {
      paint($('[data-tiles]'), '');
      return;
    }
    paint(
      $('[data-tiles]'),
      html`<div class="tile"><div class="label">Excluídos</div><div class="value">${fmtCompact(t.uniqueDeleted)}</div><div class="detail">${plural(t.attempts, 'tentativa de exclusão', 'tentativas de exclusão')} no total</div></div>
        <div class="tile"><div class="label">Espaço liberado</div><div class="value">${fmtBytes(t.bytesDeleted)}</div><div class="detail">${t.bytesUnknown ? `${plural(t.bytesUnknown, 'item excluído', 'itens excluídos')} sem o tamanho registrado` : 'somando os itens excluídos'}</div></div>
        <div class="tile"><div class="label">Falhas ou alterados</div><div class="value">${fmtCompact(t.failed + t.changed)}</div><div class="detail">${plural(t.failed, 'falha', 'falhas')} · ${fmtNum(t.changed)} alterado${t.changed === 1 ? '' : 's'} (mantido${t.changed === 1 ? '' : 's'})${t.uniquePending ? ` · ${fmtNum(t.uniquePending)} ainda no lugar` : ''}</div></div>
        <div class="tile"><div class="label">Não encontrados</div><div class="value">${fmtCompact(t.missing)}</div><div class="detail">já excluídos ou movidos antes</div></div>`,
    );
  };

  /** Tabela pequena (por política ou por repositório); cada linha filtra a lista ao ser clicada. */
  const groupTable = ({ title, head, rows, param, valueOf, sub }) => {
    const body = rows.length
      ? html`<div class="table-wrap">
            <table class="data compact">
              <thead><tr><th>${head}</th><th class="num">Excluídos</th><th class="num">Espaço</th><th class="num">Falhas</th></tr></thead>
              <tbody>
                ${rows.slice(0, TOP).map((g) => {
                  const value = valueOf(g);
                  const active = Boolean(value) && filters[param] === value;
                  return html`<tr class="${active ? 'is-active' : ''}">
                    <td>${value
                      ? html`<button type="button" class="link-btn" data-filter-param="${param}" data-filter-value="${value}" aria-pressed="${active ? 'true' : 'false'}" title="${active ? 'Remover o filtro' : 'Filtrar a lista'}">${g.name}</button>`
                      : g.name}<div class="muted small">${sub(g)}</div></td>
                    <td class="num">${fmtNum(g.uniqueDeleted)}</td>
                    <td class="num nowrap">${fmtBytes(g.bytesDeleted)}</td>
                    <td class="num">${fmtNum(g.failed + g.changed)}</td>
                  </tr>`;
                })}
              </tbody>
            </table>
          </div>
          ${rows.length > TOP ? html`<p class="muted small">+ ${fmtNum(rows.length - TOP)} não exibidos — use o filtro acima ou exporte.</p>` : ''}`
      : html`<div class="empty">Nenhuma exclusão.</div>`;
    return html`<section class="card"><div class="card-head"><h2>${title}</h2></div>${body}</section>`;
  };

  const drawGroups = () => {
    if (!summary || !summary.totals.attempts) {
      paint($('[data-groups]'), '');
      return;
    }
    paint(
      $('[data-groups]'),
      html`${groupTable({
        title: 'Por política',
        head: 'Política/agendamento',
        rows: summary.byPolicy,
        param: 'politica',
        valueOf: (g) => g.key ?? 'none',
        sub: (g) => (g.key === null ? 'exclusões manuais e de análises avulsas' : `${g.purposeLabel}${g.exists ? '' : ' · excluída do cadastro'}`),
      })}
      ${groupTable({
        title: 'Por repositório',
        head: 'Repositório/conexão',
        rows: summary.byRepository,
        param: 'repositorio',
        valueOf: (g) => g.key,
        sub: (g) => g.kindLabel,
      })}`,
    );
  };

  const row = (n) => html`<tr>
    <td class="nowrap">${fmtServerDateTime(n.at, { weekday: false })}</td>
    <td>${statusBadge(n)}</td>
    <td class="deletion-item">
      <div class="path">${n.item || '—'}</div>
      ${n.folder || n.size !== null ? html`<div class="muted small">${n.folder ? `${n.kind === 'teams' ? 'Conversa' : 'Pasta'}: ${n.folder}` : ''}${n.folder && n.size !== null ? ' · ' : ''}${n.size !== null ? fmtBytes(n.size) : ''}</div>` : ''}
      ${n.detail ? html`<div class="small ${n.status === 'failed' ? 'danger-text' : 'muted'}">${n.detail}</div>` : ''}
      <div class="muted small">${scanLink(n)}</div>
    </td>
    <td>${n.repositoryName}<div class="muted small">${n.kindLabel}</div></td>
    <td>${n.scheduleName || html`<span class="muted">Sem política</span>`}${n.purpose ? html`<div class="muted small">${n.purposeLabel}</div>` : ''}</td>
    <td>${n.originLabel}${n.methodLabel ? html`<div class="muted small">${n.methodLabel}</div>` : ''}</td>
    <td class="small">${n.by || '—'}</td>
  </tr>`;

  const drawResults = () => {
    const box = $('[data-results]');
    if (failure) {
      paint(box, html`<div class="alert error">${icon('alert')}<div><b>Não foi possível consultar o registro.</b> ${failure}</div></div>`);
      return;
    }
    if (!data) {
      paint(box, html`<p class="loading">Carregando…</p>`);
      return;
    }
    const filtered = criteria().length > 0;
    if (data.total === 0) {
      paint(
        box,
        filtered
          ? html`<div class="empty"><p>Nenhuma exclusão com os filtros escolhidos.</p><button type="button" class="btn" data-action="clear-filters">Limpar filtros</button></div>`
          : html`<div class="empty">
              <p><b>Nenhuma exclusão registrada ainda.</b></p>
              <p>As exclusões feitas pelo CLEAN aparecem aqui: as automáticas (políticas de retenção, buscas por tipo e análises com exclusão automática), as manuais pelo relatório e as em lote.</p>
            </div>`,
      );
      return;
    }
    paint(
      box,
      html`<div class="card-head">
          <h2 tabindex="-1">${plural(data.total, 'exclusão', 'exclusões')}${filtered ? ' com os filtros' : ''}</h2>
          <span class="muted small">das mais recentes para as mais antigas · horários do servidor</span>
        </div>
        <div class="table-wrap">
          <table class="data deletions">
            <thead><tr><th>Quando</th><th>Situação</th><th>Item</th><th>Repositório</th><th>Política</th><th>Origem / como</th><th>Por</th></tr></thead>
            <tbody>${data.items.map(row)}</tbody>
          </table>
        </div>
        ${data.pages > 1
          ? html`<div class="pager">
              <span class="muted small">Página ${fmtNum(data.page)} de ${fmtNum(data.pages)}</span>
              <div class="inline">
                <button type="button" class="btn small" data-action="page" data-page="${data.page - 1}" ${data.page <= 1 ? 'disabled' : ''}>Anterior</button>
                <button type="button" class="btn small" data-action="page" data-page="${data.page + 1}" ${data.page >= data.pages ? 'disabled' : ''}>Próxima</button>
              </div>
            </div>`
          : ''}`,
    );
  };

  const drawAll = () =>
    redraw(root, () => {
      drawExports();
      drawOptions();
      drawScope();
      drawTiles();
      drawGroups();
      drawResults();
    });

  // ---------- Carregamento ----------

  const load = async () => {
    const busy = [$('[data-groups]'), $('[data-results]')];
    busy.forEach((el) => el.classList.add('is-loading'));
    const request = Promise.all([get(`/api/deletions?${apiQuery({ page: filters.page, pageSize: PAGE_SIZE })}`), get(`/api/deletions/summary?${apiQuery()}`)]);
    loading = request;
    // O foco no título da lista vale só para a carga pedida pela mudança de página (e não para outra
    // carga que termine antes, ex.: a de uma busca digitada enquanto isso).
    const wantFocus = focusResults;
    focusResults = false;
    try {
      const [page, sum] = await request;
      // Outra tela aberta enquanto esperava: nada é desenhado nem vai para o endereço dela.
      if (stopped || !isCurrent() || loading !== request) return;
      data = page;
      summary = sum;
      failure = '';
      // A página pedida não existe mais (ex.: menos exclusões com o novo filtro): fica a última.
      if (data.page !== Number(filters.page || 1)) {
        filters.page = data.page > 1 ? String(data.page) : '';
        syncUrl();
      }
      drawAll();
      // Mudança de página: o foco vai para o título da lista (depois de redesenhada).
      if (wantFocus) $('[data-results] h2')?.focus({ preventScroll: true });
      announce(`${plural(data.total, 'exclusão encontrada', 'exclusões encontradas')}.`);
    } catch (err) {
      if (stopped || !isCurrent() || loading !== request) return;
      // Sem números antigos na tela: eles seriam de outro recorte.
      failure = err.message;
      data = null;
      summary = null;
      redraw(root, () => {
        drawTiles();
        drawGroups();
        drawResults();
      });
      announce(`Erro: ${err.message}`);
    } finally {
      if (loading === request) busy.forEach((el) => el.classList.remove('is-loading'));
    }
  };

  let focusResults = false;

  const applyFilters = (changes) => {
    if (stopped || !isCurrent()) return;
    Object.assign(filters, changes);
    if (!('page' in changes)) filters.page = '';
    syncUrl();
    drawExports();
    drawScope();
    load();
  };

  // ---------- Eventos ----------

  const onClick = (event) => {
    const filterButton = event.target.closest('[data-filter-param]');
    if (filterButton) {
      const param = filterButton.dataset.filterParam;
      const value = filterButton.dataset.filterValue;
      const next = filters[param] === value ? '' : value;
      if (form.elements[param]) form.elements[param].value = next;
      applyFilters({ [param]: next });
      return;
    }
    const el = event.target.closest('[data-action]');
    if (!el) return;
    const action = el.dataset.action;
    if (action === 'page') {
      focusResults = true;
      applyFilters({ page: el.dataset.page > 1 ? el.dataset.page : '' });
      $('[data-results]').scrollIntoView({ block: 'start' });
      $('[data-results] h2')?.focus({ preventScroll: true });
    } else if (action === 'clear-filters') {
      onSearch.cancel();
      Object.values(onDate).forEach((fn) => fn.cancel());
      for (const field of form.elements) if (field.name) field.value = '';
      applyFilters(Object.fromEntries(Object.keys(PARAMS).map((key) => [key, ''])));
    } else if (action === 'remove-filter') {
      applyFilters({ [el.dataset.filter]: '' });
    }
  };

  // Datas: aplicadas depois de uma pausa e só completas (ao digitar o ano, o campo passa por valores
  // como 0002-10-01, que não devem virar consultas nem ir para o endereço). Outras datas recusadas
  // pelo servidor (ex.: antes de 1970) seguem, para o motivo aparecer na tela.
  const completeDate = (value) => value === '' || (/^\d{4}-\d{2}-\d{2}$/.test(value) && !value.startsWith('0'));
  const onDate = Object.fromEntries(
    ['de', 'ate'].map((name) => [
      name,
      debounce((value) => {
        if (!stopped && completeDate(value) && value !== filters[name]) applyFilters({ [name]: value });
      }, 600),
    ]),
  );
  const onChange = (event) => {
    const { name, value } = event.target;
    if (onDate[name]) onDate[name](value);
    else if (name && name !== 'q' && Object.hasOwn(PARAMS, name)) applyFilters({ [name]: value });
  };
  const onSearch = debounce((value) => {
    if (!stopped) applyFilters({ q: value.trim() });
  }, 350);
  const onInput = (event) => {
    if (event.target.name === 'q') onSearch(event.target.value);
  };
  const onSubmit = (event) => event.preventDefault();

  root.addEventListener('click', onClick);
  form.addEventListener('change', onChange);
  form.addEventListener('input', onInput);
  form.addEventListener('submit', onSubmit);

  drawExports();
  drawScope();
  await load();

  return () => {
    stopped = true;
    onSearch.cancel();
    Object.values(onDate).forEach((fn) => fn.cancel());
    root.removeEventListener('click', onClick);
  };
}
