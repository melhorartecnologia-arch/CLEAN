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

async function route() {
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
  document.querySelectorAll('.nav a').forEach((a) => {
    if (match && a.dataset.nav === match[2]) a.setAttribute('aria-current', 'page');
    else a.removeAttribute('aria-current');
  });
  document.getElementById('tooltip').hidden = true;
  if (!match) {
    render(view, html`<div class="empty"><h1>Página não encontrada</h1><p><a href="#/">Voltar ao painel</a></p></div>`);
    return;
  }
  const [m, mod] = match;
  render(view, html`<p class="loading">Carregando…</p>`);
  try {
    if (!ctx.info) ctx.info = await get('/api/info');
    const result = await mod.render(view, { params: m.slice(1), query: new URLSearchParams(query), ctx, isCurrent: () => token === navigation });
    if (token === navigation) cleanup = typeof result === 'function' ? result : null;
    else if (typeof result === 'function') result();
    view.focus({ preventScroll: true });
  } catch (err) {
    if (token !== navigation) return;
    render(view, html`<div class="alert error">${icon('alert')}<div><b>Não foi possível carregar a página.</b><br />${err.message}</div></div>`);
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
