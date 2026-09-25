// Roteador da interface (rotas no fragmento da URL: #/analises/123?termo=...).
import { get } from './api.js';
import { html, render, icon } from './ui.js';
import * as dashboard from './views/dashboard.js';
import * as scans from './views/scans.js';
import * as scanNew from './views/scan-new.js';
import * as report from './views/report.js';
import * as repositories from './views/repositories.js';
import * as lists from './views/lists.js';
import * as listEdit from './views/list-edit.js';

const ROUTES = [
  [/^\/$/, dashboard, 'painel'],
  [/^\/analises$/, scans, 'analises'],
  [/^\/analises\/nova$/, scanNew, 'analises'],
  [/^\/analises\/([\w-]+)$/, report, 'analises'],
  [/^\/repositorios$/, repositories, 'repositorios'],
  [/^\/listas$/, lists, 'listas'],
  [/^\/listas\/nova$/, listEdit, 'listas'],
  [/^\/listas\/([\w-]+)$/, listEdit, 'listas'],
];

const view = document.getElementById('view');
const ctx = { info: null };
let cleanup = null;
let navigation = 0;
let currentUrl = location.href;
let leaveGuard = null; // função async que decide se é possível sair da tela atual

async function route() {
  // Tela com alterações não salvas: pergunta antes de sair (e volta a URL se o usuário desistir).
  if (leaveGuard && location.href !== currentUrl) {
    const guard = leaveGuard;
    leaveGuard = null;
    if (!(await guard())) {
      leaveGuard = guard;
      history.replaceState(null, '', currentUrl);
      return;
    }
  }
  currentUrl = location.href;
  const token = ++navigation;
  const hash = location.hash.replace(/^#/, '') || '/';
  const [path, query = ''] = hash.split('?');
  const match = ROUTES.map(([re, mod, nav]) => [re.exec(path), mod, nav]).find(([m]) => m);
  if (cleanup) {
    try {
      cleanup();
    } catch {
      // ignora
    }
    cleanup = null;
  }
  leaveGuard = null;
  document.querySelectorAll('.nav a').forEach((a) => {
    if (match && a.dataset.nav === match[2]) a.setAttribute('aria-current', 'page');
    else a.removeAttribute('aria-current');
  });
  document.getElementById('tooltip').hidden = true;
  // Cada navegação desenha em um contêiner novo: respostas atrasadas da tela anterior vão para um
  // elemento que já saiu da página e não aparecem por cima da tela atual.
  const container = document.createElement('div');
  view.replaceChildren(container);
  if (!match) {
    render(container, html`<div class="empty"><h1>Página não encontrada</h1><p><a href="#/">Voltar ao painel</a></p></div>`);
    return;
  }
  const [m, mod] = match;
  render(container, html`<p class="loading">Carregando…</p>`);
  const isCurrent = () => token === navigation;
  try {
    if (!ctx.info) ctx.info = await get('/api/info');
    const result = await mod.render(container, {
      params: m.slice(1),
      query: new URLSearchParams(query),
      ctx,
      isCurrent,
      setLeaveGuard: (fn) => {
        if (isCurrent()) leaveGuard = fn;
      },
    });
    if (isCurrent()) cleanup = typeof result === 'function' ? result : null;
    else if (typeof result === 'function') result();
    // Não tira o foco de quem já começou a usar a tela (ex.: digitando na busca).
    if (isCurrent() && !container.contains(document.activeElement)) view.focus({ preventScroll: true });
  } catch (err) {
    if (!isCurrent()) return;
    render(container, html`<div class="alert error">${icon('alert')}<div><b>Não foi possível carregar a página.</b><br />${err.message}</div></div>`);
  }
}

async function showServerInfo() {
  try {
    ctx.info ||= await get('/api/info');
    const i = ctx.info;
    const platform = i.platform === 'win32' ? 'Windows' : i.platform;
    render(
      document.getElementById('server-info'),
      html`<div>Servidor: <b>${i.hostname}</b></div>
        <div>Conta: ${i.user} · ${platform}</div>
        <div>Versão ${i.version}</div>`,
    );
  } catch {
    // opcional
  }
}

window.addEventListener('hashchange', () => {
  route();
  window.scrollTo(0, 0);
});
route().then(showServerInfo);
