// Visualizador do Microsoft Teams ao vivo (somente leitura): lê as conversas de um usuário —
// chats (1:1, em grupo e de reunião) e os canais das equipes de que ele participa — falando com o
// Microsoft Graph na hora, sob demanda. Nada é gravado; os botões "Atualizar" e "Carregar mais"
// buscam novas páginas quando o usuário pede (sem atualização automática).
import { get } from '../api.js';
import { html, render as paint, icon, toast, fmtDateTime, plural, debounce } from '../ui.js';
import { replaceQuery } from '../nav.js';

const CHAT_TYPES = { oneOnOne: 'Conversa', group: 'Grupo', meeting: 'Reunião' };

export async function render(root, { query }) {
  const sources = (await get('/api/mail-sources')).filter((s) => s.type === 'graph');

  if (sources.length === 0) {
    paint(
      root,
      html`<div class="page-head"><div><h1>Microsoft Teams ao vivo</h1></div></div>
        <div class="alert info">${icon('info')}<div>
          Para ler o Teams ao vivo é preciso ao menos uma conexão <b>Microsoft 365</b> (o mesmo aplicativo do e-mail, com as permissões do Teams consentidas pelo administrador).
          <div class="inline page-actions"><a class="btn small" href="#/email/caixas">Cadastrar conexão Microsoft 365</a></div>
        </div></div>`,
    );
    return null;
  }

  // Estado do visualizador (em memória; a conexão e o usuário ficam na URL para reabrir/compartilhar).
  const state = {
    source: null, // conexão escolhida
    user: null, // usuário resolvido { id, name, mail, upn, address }
    convs: null, // { chats:[], chatsNext, teams:[] }
    convMap: new Map(), // chave → descritor da conversa
    selectedKey: null,
    msgs: [], // mensagens carregadas da conversa selecionada (ordem cronológica: a mais antiga primeiro)
    msgsNext: null,
    loadingMsgs: false,
    loadSeq: 0, // identifica a carga de mensagens mais recente (troca de conversa cancela as anteriores)
    replies: new Map(), // messageId → { open, items, next, loading, loaded }
    filter: '', // filtro rápido das mensagens carregadas da conversa aberta
    search: { active: false, q: '', loading: false, results: [], truncated: false, seq: 0 }, // busca geral
  };
  let stopped = false;
  const alive = () => !stopped;

  const prefillSource = query.get('conexao') || (sources.length === 1 ? sources[0].id : '');
  const prefillUser = query.get('usuario') || '';

  // ---------- Tela de abertura (escolher conexão e usuário) ----------

  function drawSetup(message) {
    paint(
      root,
      html`<div class="page-head">
          <div>
            <h1>Microsoft Teams ao vivo</h1>
            <div class="sub">Somente leitura das conversas de um usuário — os chats e os canais das equipes de que ele participa — direto do Microsoft 365, sob demanda. Nada é gravado.</div>
          </div>
        </div>
        <form class="card" data-open novalidate>
          <div class="form-grid">
            <label class="field">
              <span>Conexão Microsoft 365</span>
              <select name="source">
                ${sources.map((s) => html`<option value="${s.id}" ${s.id === prefillSource ? 'selected' : ''}>${s.name}</option>`)}
              </select>
              <small>Um aplicativo (segredo ou certificado) lê qualquer usuário; uma conta conectada vê só as próprias conversas.</small>
            </label>
            <label class="field">
              <span>E-mail do usuário</span>
              <input type="email" name="address" autocomplete="off" spellcheck="false" placeholder="pessoa@empresa.com" value="${prefillUser}" />
              <small>Quem terá as conversas lidas. Use o e-mail ou o nome de logon (UPN).</small>
            </label>
          </div>
          ${message ? html`<div class="alert error">${icon('alert')}<div>${message}</div></div>` : ''}
          <div class="inline page-actions">
            <button type="submit" class="btn primary" data-submit>${icon('search')} Abrir conversas</button>
          </div>
        </form>`,
    );
  }

  // ---------- Abertura: resolve o usuário e carrega as conversas ----------

  async function open(sourceId, address) {
    const submit = root.querySelector('[data-submit]');
    if (submit) submit.disabled = true;
    try {
      const source = sources.find((s) => s.id === sourceId);
      const { user } = await get(`/api/teams-live/${encodeURIComponent(sourceId)}/user?address=${encodeURIComponent(address)}`);
      if (!alive()) return;
      const convs = await get(`/api/teams-live/${encodeURIComponent(sourceId)}/conversations?userId=${encodeURIComponent(user.id)}`);
      if (!alive()) return;
      state.source = source;
      state.user = user;
      state.convs = convs;
      state.selectedKey = null;
      state.msgs = [];
      state.msgsNext = null;
      state.replies = new Map();
      state.filter = '';
      state.search = { active: false, q: '', loading: false, results: [], truncated: false, seq: 0 };
      indexConversations();
      replaceQuery({ conexao: sourceId, usuario: user.address || address });
      drawViewer();
    } catch (err) {
      if (!alive()) return;
      drawSetup(err.message);
    }
  }

  /** Monta o mapa chave → descritor para a barra lateral e a seleção. */
  function indexConversations() {
    state.convMap = new Map();
    for (const c of state.convs.chats || []) state.convMap.set(chatKey(c.id), { kind: 'chat', chatId: c.id, label: c.label, chatType: c.chatType, members: c.members });
    for (const t of state.convs.teams || []) {
      for (const ch of t.channels || []) {
        state.convMap.set(channelKey(t.id, ch.id), { kind: 'channel', teamId: t.id, channelId: ch.id, teamName: t.name, channelName: ch.name, membershipType: ch.membershipType, label: `${t.name} › ${ch.name}` });
      }
    }
  }

  // ---------- Visualizador (duas colunas: conversas | mensagens) ----------

  function drawViewer() {
    const u = state.user;
    paint(
      root,
      html`<div class="page-head">
          <div>
            <h1>Teams ao vivo · ${u.name || u.address}</h1>
            <div class="sub">${u.address} · ${state.source.name}</div>
          </div>
          <div class="actions">
            <form class="inline live-search" data-search novalidate>
              <input type="search" name="q" data-search-input aria-label="Buscar em todas as conversas" placeholder="Buscar em todas as conversas" value="${state.search.q}" />
              <button type="submit" class="btn">${icon('search')} Buscar</button>
            </form>
            <button class="btn" data-act="refresh-convs">${icon('refresh')} Recarregar</button>
            <button class="btn" data-act="reset">${icon('user')} Trocar usuário</button>
          </div>
        </div>
        <div class="live">
          <aside class="live-side" data-side></aside>
          <section class="live-main" data-main></section>
        </div>`,
    );
    drawSide();
    drawMain();
  }

  function drawSide() {
    const { chats = [], chatsNext, teams = [] } = state.convs;
    const sideEl = root.querySelector('[data-side]');
    const item = (key, label, sub) => html`<button class="conv-item ${key === state.selectedKey ? 'is-selected' : ''}" data-act="open" data-key="${key}" aria-current="${key === state.selectedKey ? 'true' : 'false'}">
        <span class="conv-label">${label}</span>${sub ? html`<span class="conv-sub">${sub}</span>` : ''}
      </button>`;
    paint(
      sideEl,
      html`<div class="conv-group">Chats${chats.length ? html` <span class="muted small">(${chats.length})</span>` : ''}</div>
        ${chats.length === 0 ? html`<p class="muted small conv-empty">Nenhum chat.</p>` : chats.map((c) => item(chatKey(c.id), c.label, CHAT_TYPES[c.chatType] || 'Conversa'))}
        ${chatsNext ? html`<button class="btn small conv-more" data-act="more-chats">Carregar mais chats</button>` : ''}
        ${teams.map(
          (t) => html`<div class="conv-group">${t.name}</div>
            ${t.error
              ? html`<p class="muted small conv-empty">Não foi possível listar os canais.</p>`
              : (t.channels || []).length === 0
                ? html`<p class="muted small conv-empty">Sem canais.</p>`
                : (t.channels || []).map((ch) => item(channelKey(t.id, ch.id), ch.name, ch.membershipType !== 'standard' ? channelKind(ch.membershipType) : ''))}`,
        )}
        ${teams.length === 0 && chats.length === 0 ? html`<p class="empty">Nenhuma conversa encontrada para este usuário.</p>` : ''}`,
    );
  }

  function drawMain() {
    if (state.search.active) return drawSearch();
    const mainEl = root.querySelector('[data-main]');
    const scroll = mainEl.scrollTop;
    const conv = state.selectedKey ? state.convMap.get(state.selectedKey) : null;
    if (!conv) {
      paint(mainEl, html`<div class="empty">Selecione uma conversa à esquerda para ver as mensagens.</div>`);
      return;
    }
    // O cabeçalho (com o filtro) é estável; só o corpo (data-body) é repintado ao filtrar, para o
    // campo de filtro não perder o foco enquanto se digita.
    paint(
      mainEl,
      html`<div class="live-head">
          <div class="live-title">${conv.label}${conv.kind === 'chat' ? html` <span class="chip">${CHAT_TYPES[conv.chatType] || 'Conversa'}</span>` : conv.membershipType && conv.membershipType !== 'standard' ? html` <span class="chip">${channelKind(conv.membershipType)}</span>` : ''}</div>
          <div class="inline">
            <input type="search" data-filter aria-label="Filtrar nesta conversa" placeholder="Filtrar nesta conversa" value="${state.filter}" />
            <button class="btn small" data-act="refresh">${icon('refresh')} Atualizar</button>
          </div>
        </div>
        ${conv.kind === 'chat' && (conv.members || []).length ? html`<div class="muted small live-members">${(conv.members || []).map((m) => m.displayName || m.email).filter(Boolean).join(', ')}</div>` : ''}
        <div class="live-body" data-body></div>`,
    );
    drawBody();
    mainEl.scrollTop = scroll;
  }

  /** Corpo da conversa: as mensagens em ordem cronológica (a mais antiga no topo), já com o filtro. */
  function drawBody() {
    const bodyEl = root.querySelector('[data-body]');
    if (!bodyEl) return;
    const mainEl = root.querySelector('[data-main]');
    const scroll = mainEl ? mainEl.scrollTop : 0;
    const conv = state.convMap.get(state.selectedKey);
    const term = state.filter.trim();
    const shown = term ? state.msgs.filter((m) => fold(`${m.from} ${m.subject} ${m.text}`).includes(fold(term))) : state.msgs;
    // Ainda há mensagens mais antigas além do que foi carregado (conversa muito longa): botão no topo.
    const older =
      state.msgsNext && !term
        ? html`<div class="live-older"><button class="btn small" data-act="older" ${state.loadingMsgs ? 'disabled' : ''}>${state.loadingMsgs ? 'Carregando…' : 'Carregar mensagens mais antigas'}</button></div>`
        : '';
    paint(
      bodyEl,
      html`${state.loadingMsgs && state.msgs.length === 0
        ? html`<p class="loading">Carregando a conversa…</p>`
        : html`${term
                ? html`<p class="muted small live-order">${plural(shown.length, 'mensagem encontrada', 'mensagens encontradas')} de ${state.msgs.length} carregada(s).</p>`
                : state.msgs.length
                  ? html`<p class="muted small live-order">${plural(state.msgs.length, 'mensagem', 'mensagens')}${state.msgsNext ? ' carregadas (há mais antigas)' : ''} · da mais antiga (topo) para a mais recente.</p>`
                  : ''}
              ${older}
              ${shown.length
                ? html`<div class="messages">${shown.map((m) => messageCard(m, conv, term))}</div>`
                : term
                  ? html`<div class="empty">Nenhuma mensagem carregada corresponde a “${term}”. Use a busca em todas as conversas para procurar fora desta conversa.</div>`
                  : html`<div class="empty">${state.msgsNext ? 'Nenhuma mensagem recente exibível — use “Carregar mensagens mais antigas”.' : 'Nenhuma mensagem nesta conversa.'}</div>`}`}`,
    );
    if (mainEl) mainEl.scrollTop = scroll;
  }

  /** Painel de resultados da busca geral. */
  function drawSearch() {
    const mainEl = root.querySelector('[data-main]');
    const s = state.search;
    paint(
      mainEl,
      html`<div class="live-head">
          <div class="live-title">Busca: “${s.q}”</div>
          <button class="btn small" data-act="search-clear">Voltar às conversas</button>
        </div>
        ${s.loading
          ? html`<p class="loading">Procurando nas conversas do usuário…</p>`
          : html`<p class="muted small live-order">${plural(s.results.length, 'resultado', 'resultados')}${s.truncated ? ' · busca limitada às mensagens recentes (para uma busca completa, use a Análise do Teams)' : ''}.</p>
              ${s.results.length === 0
                ? html`<div class="empty">Nada encontrado para “${s.q}”.</div>`
                : html`<div class="messages">${s.results.map((m) => searchResultCard(m))}</div>`}`}`,
    );
  }

  /** Cartão de uma mensagem; nos canais, inclui o botão de respostas e as respostas abertas. */
  function messageCard(m, conv, term = '') {
    const rep = state.replies.get(m.id);
    const canReplies = conv.kind === 'channel' && m.hasReplies;
    return html`<article class="msg">
      <div class="msg-head">
        <span class="msg-from">${m.from || 'Desconhecido'}</span>
        <span class="muted small">${fmtDateTime(m.date)}${m.edited ? html` · <span title="Editada em ${fmtDateTime(m.edited)}">editada</span>` : ''}</span>
      </div>
      ${quoteBlock(m)}
      ${m.subject ? html`<div class="msg-subject">${highlight(m.subject, term)}</div>` : ''}
      ${m.text ? html`<div class="msg-body">${highlight(m.text, term)}</div>` : html`<div class="msg-body muted"><em>(sem texto)</em></div>`}
      ${images(conv, m, null)}
      ${attachments(m)}
      <div class="msg-foot">
        ${m.webUrl ? html`<a class="msg-link" href="${m.webUrl}" target="_blank" rel="noopener noreferrer">Abrir no Teams</a>` : ''}
        ${canReplies ? html`<button class="link-btn" data-act="replies" data-id="${m.id}">${rep?.open ? 'Ocultar respostas' : repliesLabel(m)}</button>` : ''}
      </div>
      ${canReplies && rep?.open ? repliesBlock(m, rep, term) : ''}
    </article>`;
  }

  /** Cartão de um resultado da busca geral (com o link para abrir a conversa). */
  function searchResultCard(m) {
    const key = m.conv.kind === 'chat' ? chatKey(m.conv.chatId) : channelKey(m.conv.teamId, m.conv.channelId);
    return html`<article class="msg">
      <div class="msg-head">
        <span class="msg-from">${m.from || 'Desconhecido'}</span>
        <span class="muted small">${fmtDateTime(m.date)}</span>
      </div>
      <div class="result-conv">${icon(m.conv.kind === 'chat' ? 'user' : 'list')} ${m.conv.label}</div>
      ${quoteBlock(m)}
      ${m.subject ? html`<div class="msg-subject">${highlight(m.subject, state.search.q)}</div>` : ''}
      ${m.text ? html`<div class="msg-body">${highlight(m.text, state.search.q)}</div>` : html`<div class="msg-body muted"><em>(sem texto)</em></div>`}
      ${images(m.conv, m, null)}
      <div class="msg-foot"><button class="link-btn" data-act="open-result" data-key="${key}">Abrir conversa</button></div>
    </article>`;
  }

  function repliesBlock(m, rep, term = '') {
    if (rep.loading && (rep.items || []).length === 0) return html`<div class="replies"><p class="loading">Carregando respostas…</p></div>`;
    if ((rep.items || []).length === 0) return html`<div class="replies"><p class="muted small">Sem respostas.</p></div>`;
    const conv = state.convMap.get(state.selectedKey);
    return html`<div class="replies">
      ${rep.next ? html`<button class="btn small" data-act="replies-more" data-id="${m.id}" ${rep.loading ? 'disabled' : ''}>${rep.loading ? 'Carregando…' : 'Carregar respostas mais antigas'}</button>` : ''}
      ${rep.items.map(
        (r) => html`<article class="msg reply">
          <div class="msg-head"><span class="msg-from">${r.from || 'Desconhecido'}</span><span class="muted small">${fmtDateTime(r.date)}${r.edited ? html` · editada` : ''}</span></div>
          ${quoteBlock(r)}
          ${r.text ? html`<div class="msg-body">${highlight(r.text, term)}</div>` : html`<div class="msg-body muted"><em>(sem texto)</em></div>`}
          ${images(conv, r, m.id)}
          ${attachments(r)}
        </article>`,
      )}
    </div>`;
  }

  /** Bloco de citação: deixa claro que a mensagem é resposta a outra (remetente e prévia). */
  function quoteBlock(m) {
    if (!m.quote) return '';
    const q = m.quote;
    return html`<div class="msg-quote">
      <div class="msg-quote-head">${icon('reply')} Em resposta a ${q.sender ? html`<b>${q.sender}</b>` : 'uma mensagem'}</div>
      ${q.preview ? html`<div class="msg-quote-text">${q.preview}</div>` : ''}
    </div>`;
  }

  /** Imagens embutidas (hosted content): miniaturas servidas pelo proxy do servidor. */
  function images(conv, m, replyTo) {
    const imgs = m.images || [];
    if (!conv || imgs.length === 0) return '';
    return html`<div class="msg-images">${imgs.map((im) => {
      const url = imageUrl(conv, m.id, replyTo, im.hostedId);
      return html`<a class="msg-image" href="${url}" target="_blank" rel="noopener noreferrer"><img loading="lazy" src="${url}" alt="Imagem da conversa" /></a>`;
    })}</div>`;
  }

  function imageUrl(conv, messageId, replyTo, hostedId) {
    const p = new URLSearchParams({ messageId, hostedId });
    if (conv.kind === 'chat') {
      p.set('kind', 'chat');
      p.set('chatId', conv.chatId);
    } else {
      p.set('kind', 'channel');
      p.set('teamId', conv.teamId);
      p.set('channelId', conv.channelId);
      if (replyTo) p.set('replyTo', replyTo);
    }
    return `/api/teams-live/${encodeURIComponent(state.source.id)}/image?${p.toString()}`;
  }

  function attachments(m) {
    const atts = m.attachments || [];
    if (atts.length === 0) return '';
    return html`<div class="msg-atts">${atts.map(
      (a) =>
        html`<span class="chip">${icon('file')} ${a.url ? html`<a href="${a.url}" target="_blank" rel="noopener noreferrer">${a.name}</a>` : a.name}</span>`,
    )}</div>`;
  }

  // ---------- Carregamento das mensagens ----------

  // Teto de lotes do histórico (cada lote traz várias páginas do servidor). ~9000 mensagens no total;
  // além disso, o botão "Carregar mensagens mais antigas" continua a leitura.
  const MAX_BATCHES = 6;

  /**
   * reset=true: abre a conversa e carrega todo o histórico (em lotes, com progresso), em ordem
   * cronológica, rolando para a mensagem mais recente. reset=false: carrega um lote de mensagens
   * mais antigas e as acrescenta no topo, mantendo a posição de leitura.
   */
  async function loadMessages(reset) {
    const conv = state.convMap.get(state.selectedKey);
    if (!conv) return;
    // Cada carga recebe um número; trocar de conversa (ou recarregar) emite um número maior e cancela
    // as anteriores — assim uma resposta atrasada nunca trava o painel nem pinta a conversa errada.
    const seq = ++state.loadSeq;
    state.loadingMsgs = true;
    const main0 = root.querySelector('[data-main]');
    const prevH = !reset && main0 ? main0.scrollHeight : 0;
    const prevTop = !reset && main0 ? main0.scrollTop : 0;
    if (reset) {
      state.msgs = [];
      state.msgsNext = null;
      state.replies = new Map();
      state.filter = '';
      drawMain();
    } else {
      drawBody();
    }
    try {
      let next = reset ? null : state.msgsNext;
      const loaded = []; // acumulado em ordem do Graph (mais recentes primeiro)
      const batches = reset ? MAX_BATCHES : 1;
      for (let b = 0; b < batches; b++) {
        const page = await fetchHistory(conv, next);
        if (!alive() || seq !== state.loadSeq) return; // superada por outra conversa/recarga
        loaded.push(...page.items);
        next = page.next || null;
        if (!next || b + 1 >= batches) break;
        const el = root.querySelector('[data-body]');
        if (reset && el) paint(el, html`<p class="loading">Carregando o histórico… ${plural(loaded.length, 'mensagem', 'mensagens')}</p>`);
      }
      if (!alive() || seq !== state.loadSeq) return;
      const chrono = loaded.reverse(); // da mais antiga para a mais recente
      state.msgs = reset ? chrono : chrono.concat(state.msgs);
      state.msgsNext = next;
    } catch (err) {
      if (alive() && seq === state.loadSeq) toast(err.message, 'error');
    } finally {
      if (alive() && seq === state.loadSeq) {
        state.loadingMsgs = false;
        drawBody();
        const el = root.querySelector('[data-main]');
        if (el) el.scrollTop = reset ? el.scrollHeight : prevTop + (el.scrollHeight - prevH);
      }
    }
  }

  function fetchHistory(conv, next) {
    const base = `/api/teams-live/${encodeURIComponent(state.source.id)}/history`;
    const q =
      conv.kind === 'chat'
        ? `kind=chat&chatId=${encodeURIComponent(conv.chatId)}`
        : `kind=channel&teamId=${encodeURIComponent(conv.teamId)}&channelId=${encodeURIComponent(conv.channelId)}`;
    return get(`${base}?${q}${next ? `&next=${encodeURIComponent(next)}` : ''}`);
  }

  async function loadReplies(messageId, more) {
    const conv = state.convMap.get(state.selectedKey);
    if (!conv || conv.kind !== 'channel') return;
    let rep = state.replies.get(messageId);
    if (!rep) {
      rep = { open: true, items: [], next: null, loading: false, loaded: false };
      state.replies.set(messageId, rep);
    }
    if (rep.loading) return;
    rep.loading = true;
    drawMain();
    try {
      const base = `/api/teams-live/${encodeURIComponent(state.source.id)}/replies`;
      const q = `teamId=${encodeURIComponent(conv.teamId)}&channelId=${encodeURIComponent(conv.channelId)}&messageId=${encodeURIComponent(messageId)}`;
      const page = await get(`${base}?${q}${more && rep.next ? `&next=${encodeURIComponent(rep.next)}` : ''}`);
      if (!alive() || state.convMap.get(state.selectedKey) !== conv) return;
      // Em ordem cronológica (a mais antiga primeiro); "mais antigas" entram no topo.
      const chrono = [...page.items].reverse();
      rep.items = more ? chrono.concat(rep.items) : chrono;
      rep.next = page.next || null;
      rep.loaded = true;
    } catch (err) {
      if (alive()) toast(err.message, 'error');
    } finally {
      rep.loading = false;
      if (alive() && state.convMap.get(state.selectedKey) === conv) drawMain();
    }
  }

  // ---------- Eventos ----------

  async function onClick(event) {
    const el = event.target.closest('[data-act]');
    if (!el) return;
    const act = el.dataset.act;
    if (act === 'reset') {
      state.source = null;
      state.user = null;
      state.convs = null;
      state.convMap = new Map();
      state.selectedKey = null;
      state.msgs = [];
      state.msgsNext = null;
      state.loadSeq++; // cancela cargas de mensagens em voo
      state.filter = '';
      state.search = { active: false, q: '', loading: false, results: [], truncated: false, seq: 0 };
      replaceQuery({});
      drawSetup();
      return;
    }
    if (act === 'open' || act === 'open-result') {
      const key = el.dataset.key;
      if (act === 'open-result') ensureConv(key);
      const sameConv = state.selectedKey === key;
      const wasSearch = state.search.active;
      state.search.active = false;
      if (sameConv && !wasSearch) return; // já aberta e não vínhamos da busca
      state.selectedKey = key;
      drawSide();
      if (sameConv) drawMain(); // volta da busca para a conversa já carregada
      else loadMessages(true); // conversa diferente: carrega
      return;
    }
    if (act === 'search-clear') {
      state.search.active = false;
      state.search.q = '';
      drawViewer();
      return;
    }
    if (act === 'refresh') return loadMessages(true);
    if (act === 'older') return loadMessages(false);
    if (act === 'more-chats') return loadMoreChats(el);
    if (act === 'refresh-convs') return refreshConversations(el);
    if (act === 'replies') {
      const id = el.dataset.id;
      const rep = state.replies.get(id);
      if (rep?.open) {
        rep.open = false;
        drawMain();
      } else if (rep?.loaded) {
        rep.open = true;
        drawMain();
      } else {
        loadReplies(id, false);
      }
      return;
    }
    if (act === 'replies-more') return loadReplies(el.dataset.id, true);
  }

  async function loadMoreChats(button) {
    button.disabled = true;
    try {
      const page = await get(`/api/teams-live/${encodeURIComponent(state.source.id)}/chats?userId=${encodeURIComponent(state.user.id)}&next=${encodeURIComponent(state.convs.chatsNext)}`);
      if (!alive() || !state.source) return;
      state.convs.chats = state.convs.chats.concat(page.items);
      state.convs.chatsNext = page.next || null;
      indexConversations();
      drawSide();
    } catch (err) {
      if (alive()) {
        toast(err.message, 'error');
        button.disabled = false;
      }
    }
  }

  async function refreshConversations(button) {
    button.disabled = true;
    try {
      const convs = await get(`/api/teams-live/${encodeURIComponent(state.source.id)}/conversations?userId=${encodeURIComponent(state.user.id)}`);
      if (!alive() || !state.source) return;
      state.convs = convs;
      indexConversations();
      // Mantém a seleção se a conversa ainda existir; senão, limpa o painel.
      if (state.selectedKey && !state.convMap.has(state.selectedKey)) {
        state.selectedKey = null;
        state.msgs = [];
        state.msgsNext = null;
      }
      drawSide();
      drawMain();
    } catch (err) {
      if (alive()) toast(err.message, 'error');
    } finally {
      if (alive()) button.disabled = false;
    }
  }

  /** Garante um descritor da conversa no mapa (ao abrir um resultado de busca fora da lista carregada). */
  function ensureConv(key) {
    if (state.convMap.has(key)) return;
    const m = state.search.results.find((r) => (r.conv.kind === 'chat' ? chatKey(r.conv.chatId) : channelKey(r.conv.teamId, r.conv.channelId)) === key);
    if (!m) return;
    const c = m.conv;
    state.convMap.set(key, c.kind === 'chat' ? { kind: 'chat', chatId: c.chatId, label: c.label, chatType: c.chatType || 'group', members: [] } : { kind: 'channel', teamId: c.teamId, channelId: c.channelId, label: c.label });
  }

  // Busca geral das mensagens recentes nos chats e canais do usuário.
  async function runSearch(q) {
    const term = String(q || '').trim();
    if (term.length < 2) return toast('Digite ao menos 2 caracteres para a busca.', 'error');
    const seq = ++state.search.seq;
    state.search = { active: true, q: term, loading: true, results: [], truncated: false, seq };
    drawMain();
    try {
      const data = await get(`/api/teams-live/${encodeURIComponent(state.source.id)}/search?userId=${encodeURIComponent(state.user.id)}&q=${encodeURIComponent(term)}`);
      if (!alive() || seq !== state.search.seq) return;
      state.search.results = data.matches || [];
      state.search.truncated = Boolean(data.truncated);
    } catch (err) {
      if (alive() && seq === state.search.seq) {
        state.search.active = false;
        toast(err.message, 'error');
        drawViewer();
        return;
      }
    } finally {
      if (alive() && seq === state.search.seq) {
        state.search.loading = false;
        drawMain();
      }
    }
  }

  function onSubmit(event) {
    const form = event.target;
    if (form.matches('[data-open]')) {
      event.preventDefault();
      const sourceId = form.elements.source.value;
      const address = String(form.elements.address.value || '').trim();
      if (!address) return toast('Informe o e-mail do usuário.', 'error');
      open(sourceId, address);
    } else if (form.matches('[data-search]')) {
      event.preventDefault();
      runSearch(form.elements.q.value);
    }
  }

  // Filtro rápido na conversa aberta: atualiza o estado na hora e repinta só o corpo (com atraso),
  // para o campo de filtro não perder o foco.
  const redrawBody = debounce(() => {
    if (alive() && !state.search.active) drawBody();
  }, 120);
  function onInput(event) {
    if (event.target.matches('[data-filter]')) {
      state.filter = event.target.value;
      redrawBody();
    }
  }

  // Delegação para toda a tela (os botões e formulários do cabeçalho ficam fora de .live). Os ouvintes
  // ficam no contêiner da tela e sobrevivem às repinturas internas.
  root.addEventListener('click', onClick);
  root.addEventListener('submit', onSubmit);
  root.addEventListener('input', onInput);

  // Início: tela de abertura. Se a URL já traz conexão e usuário (reabrir/compartilhar), abre direto
  // uma vez — "Trocar usuário" depois volta para a tela de abertura sem reabrir sozinho.
  drawSetup();
  if (prefillUser && prefillSource) open(prefillSource, prefillUser);

  return () => {
    stopped = true;
    redrawBody.cancel();
    root.removeEventListener('click', onClick);
    root.removeEventListener('submit', onSubmit);
    root.removeEventListener('input', onInput);
  };
}

// ---------- Auxiliares ----------

const chatKey = (id) => `c:${id}`;
const channelKey = (teamId, channelId) => `h:${teamId}/${channelId}`;

function channelKind(membershipType) {
  if (membershipType === 'private') return 'privado';
  if (membershipType === 'shared') return 'compartilhado';
  return membershipType;
}

function repliesLabel(m) {
  const n = Number(m.replyCount);
  return n > 0 ? `Ver respostas (${n})` : 'Ver respostas';
}

/** Minúsculas e sem acentos, para comparar na busca/filtro. */
function fold(value) {
  return String(value || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '');
}

/**
 * Realça as ocorrências do termo no texto (comparação sem diferenciar maiúsculas). O texto é escapado
 * pelo template `html`; só as marcas <mark> são HTML.
 */
function highlight(text, term) {
  const t = String(term || '').trim();
  if (!t) return text;
  const re = new RegExp(`(${t.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')})`, 'gi');
  const parts = String(text).split(re);
  return html`${parts.map((part, i) => (i % 2 === 1 ? html`<mark>${part}</mark>` : part))}`;
}
