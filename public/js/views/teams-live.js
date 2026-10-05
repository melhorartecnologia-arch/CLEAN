// Visualizador do Microsoft Teams ao vivo (somente leitura): lê as conversas de um usuário —
// chats (1:1, em grupo e de reunião) e os canais das equipes de que ele participa — falando com o
// Microsoft Graph na hora, sob demanda. Nada é gravado; os botões "Atualizar" e "Carregar mais"
// buscam novas páginas quando o usuário pede (sem atualização automática).
import { get } from '../api.js';
import { html, render as paint, icon, toast, fmtDateTime } from '../ui.js';
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
    msgs: [], // mensagens carregadas da conversa selecionada (mais recentes primeiro)
    msgsNext: null,
    loadingMsgs: false,
    loadSeq: 0, // identifica a carga de mensagens mais recente (troca de conversa cancela as anteriores)
    replies: new Map(), // messageId → { open, items, next, loading, loaded }
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
    const form = root.querySelector('[data-open]');
    form.addEventListener('submit', (event) => {
      event.preventDefault();
      const sourceId = form.elements.source.value;
      const address = String(form.elements.address.value || '').trim();
      if (!address) return toast('Informe o e-mail do usuário.', 'error');
      open(sourceId, address);
    });
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
            <button class="btn" data-act="refresh-convs">${icon('refresh')} Recarregar conversas</button>
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
    const mainEl = root.querySelector('[data-main]');
    const scroll = mainEl.scrollTop;
    const conv = state.selectedKey ? state.convMap.get(state.selectedKey) : null;
    if (!conv) {
      paint(mainEl, html`<div class="empty">Selecione uma conversa à esquerda para ver as mensagens.</div>`);
      return;
    }
    paint(
      mainEl,
      html`<div class="live-head">
          <div class="live-title">${conv.label}${conv.kind === 'chat' ? html` <span class="chip">${CHAT_TYPES[conv.chatType] || 'Conversa'}</span>` : conv.membershipType && conv.membershipType !== 'standard' ? html` <span class="chip">${channelKind(conv.membershipType)}</span>` : ''}</div>
          <button class="btn small" data-act="refresh">${icon('refresh')} Atualizar</button>
        </div>
        ${conv.kind === 'chat' && (conv.members || []).length ? html`<div class="muted small live-members">${(conv.members || []).map((m) => m.displayName || m.email).filter(Boolean).join(', ')}</div>` : ''}
        ${state.loadingMsgs && state.msgs.length === 0
          ? html`<p class="loading">Carregando mensagens…</p>`
          : state.msgs.length === 0 && !state.msgsNext
            ? html`<div class="empty">Nenhuma mensagem nesta conversa.</div>`
            : html`${state.msgs.length
                  ? html`<div class="messages">
                      <p class="muted small live-order">As mais recentes primeiro.</p>
                      ${state.msgs.map((m) => messageCard(m, conv))}
                    </div>`
                  : html`<div class="empty">Nenhuma mensagem exibível nesta página (apenas mensagens de sistema). Carregue as mais antigas.</div>`}
                ${state.msgsNext ? html`<div class="live-foot"><button class="btn" data-act="older" ${state.loadingMsgs ? 'disabled' : ''}>${state.loadingMsgs ? 'Carregando…' : 'Carregar mais antigas'}</button></div>` : ''}`}`,
    );
    mainEl.scrollTop = scroll;
  }

  /** Cartão de uma mensagem; nos canais, inclui o botão de respostas e as respostas abertas. */
  function messageCard(m, conv) {
    const rep = state.replies.get(m.id);
    const canReplies = conv.kind === 'channel' && m.hasReplies;
    return html`<article class="msg">
      <div class="msg-head">
        <span class="msg-from">${m.from || 'Desconhecido'}</span>
        <span class="muted small">${fmtDateTime(m.date)}${m.edited ? html` · <span title="Editada em ${fmtDateTime(m.edited)}">editada</span>` : ''}</span>
      </div>
      ${m.subject ? html`<div class="msg-subject">${m.subject}</div>` : ''}
      ${m.text ? html`<div class="msg-body">${m.text}</div>` : html`<div class="msg-body muted"><em>(sem texto)</em></div>`}
      ${attachments(m)}
      <div class="msg-foot">
        ${m.webUrl ? html`<a class="msg-link" href="${m.webUrl}" target="_blank" rel="noopener noreferrer">Abrir no Teams</a>` : ''}
        ${canReplies ? html`<button class="link-btn" data-act="replies" data-id="${m.id}">${rep?.open ? 'Ocultar respostas' : repliesLabel(m)}</button>` : ''}
      </div>
      ${canReplies && rep?.open ? repliesBlock(m, rep) : ''}
    </article>`;
  }

  function repliesBlock(m, rep) {
    if (rep.loading && (rep.items || []).length === 0) return html`<div class="replies"><p class="loading">Carregando respostas…</p></div>`;
    if ((rep.items || []).length === 0) return html`<div class="replies"><p class="muted small">Sem respostas.</p></div>`;
    return html`<div class="replies">
      ${rep.items.map(
        (r) => html`<article class="msg reply">
          <div class="msg-head"><span class="msg-from">${r.from || 'Desconhecido'}</span><span class="muted small">${fmtDateTime(r.date)}${r.edited ? html` · editada` : ''}</span></div>
          ${r.text ? html`<div class="msg-body">${r.text}</div>` : html`<div class="msg-body muted"><em>(sem texto)</em></div>`}
          ${attachments(r)}
        </article>`,
      )}
      ${rep.next ? html`<button class="btn small" data-act="replies-more" data-id="${m.id}" ${rep.loading ? 'disabled' : ''}>${rep.loading ? 'Carregando…' : 'Carregar mais respostas'}</button>` : ''}
    </div>`;
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

  async function loadMessages(reset) {
    const conv = state.convMap.get(state.selectedKey);
    if (!conv) return;
    // Cada carga recebe um número; trocar de conversa (ou recarregar) emite um número maior e cancela
    // as anteriores — assim uma resposta atrasada nunca trava o painel nem pinta a conversa errada.
    const seq = ++state.loadSeq;
    state.loadingMsgs = true;
    if (reset) {
      state.msgs = [];
      state.msgsNext = null;
      state.replies = new Map();
    }
    drawMain();
    try {
      const page = await fetchMessages(conv, reset ? null : state.msgsNext);
      if (!alive() || seq !== state.loadSeq) return; // superada por uma carga mais recente
      state.msgs = reset ? page.items : state.msgs.concat(page.items);
      state.msgsNext = page.next || null;
    } catch (err) {
      if (alive() && seq === state.loadSeq) toast(err.message, 'error');
    } finally {
      if (alive() && seq === state.loadSeq) {
        state.loadingMsgs = false;
        drawMain();
      }
    }
  }

  function fetchMessages(conv, next) {
    const base = `/api/teams-live/${encodeURIComponent(state.source.id)}/messages`;
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
      rep.items = more ? rep.items.concat(page.items) : page.items;
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
      replaceQuery({});
      drawSetup();
      return;
    }
    if (act === 'open') {
      if (state.selectedKey === el.dataset.key) return;
      state.selectedKey = el.dataset.key;
      drawSide();
      loadMessages(true);
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
      if (!alive()) return;
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
      if (!alive()) return;
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

  // Delegação de cliques para toda a tela (os botões "Trocar usuário"/"Recarregar conversas" ficam
  // no cabeçalho, fora de .live). Fica no contêiner da tela e sobrevive às repinturas internas.
  root.addEventListener('click', onClick);

  // Início: tela de abertura. Se a URL já traz conexão e usuário (reabrir/compartilhar), abre direto
  // uma vez — "Trocar usuário" depois volta para a tela de abertura sem reabrir sozinho.
  drawSetup();
  if (prefillUser && prefillSource) open(prefillSource, prefillUser);

  return () => {
    stopped = true;
    root.removeEventListener('click', onClick);
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
