// Roteador da interface (rotas no fragmento da URL: #/analises/123?termo=...).
import { get } from './api.js';
import { html, render, icon, setServerTimeZone } from './ui.js';
import { setActiveNav } from './nav.js';
import * as dashboard from './views/dashboard.js';
import * as scans from './views/scans.js';
import * as scanNew from './views/scan-new.js';
import * as report from './views/report.js';
import * as repositories from './views/repositories.js';
import * as lists from './views/lists.js';
import * as listEdit from './views/list-edit.js';
import * as mailSources from './views/mail-sources.js';
import * as mailScanNew from './views/mail-scan-new.js';
import * as schedules from './views/schedules.js';
import * as scheduleEdit from './views/schedule-edit.js';
import * as retention from './views/retention.js';
import * as retentionEdit from './views/retention-edit.js';

// [caminho, tela, item do menu, propriedades extras da tela]
const ROUTES = [
  [/^\/$/, dashboard, 'painel'],
  [/^\/analises$/, scans, 'analises', { kind: 'files' }],
  [/^\/analises\/nova$/, scanNew, 'analises'],
  [/^\/analises\/([\w-]+)$/, report, 'analises'],
  [/^\/repositorios$/, repositories, 'repositorios'],
  [/^\/email\/caixas$/, mailSources, 'email-caixas'],
  [/^\/email\/analises$/, scans, 'email-analises', { kind: 'mail' }],
  [/^\/email\/analises\/nova$/, mailScanNew, 'email-analises'],
  [/^\/email\/analises\/([\w-]+)$/, report, 'email-analises'],
  [/^\/listas$/, lists, 'listas'],
  [/^\/listas\/nova$/, listEdit, 'listas'],
  [/^\/listas\/([\w-]+)$/, listEdit, 'listas'],
  [/^\/agendamentos$/, schedules, 'agendamentos'],
  [/^\/agendamentos\/novo$/, scheduleEdit, 'agendamentos'],
  [/^\/agendamentos\/([\w-]+)$/, scheduleEdit, 'agendamentos'],
  [/^\/retencao$/, retention, 'retencao'],
  [/^\/retencao\/nova$/, retentionEdit, 'retencao'],
  [/^\/retencao\/([\w-]+)$/, retentionEdit, 'retencao'],
];

/** Informações do servidor (versão, fuso, opções padrão), lidas uma vez. */
async function loadInfo() {
  if (!ctx.info) {
    ctx.info = await get('/api/info');
    setServerTimeZone(ctx.info.timeZone);
  }
  return ctx.info;
}

const view = document.getElementById('view');
const ctx = { info: null };
let cleanup = null;
let navigation = 0;
let currentUrl = location.href;
let leaveGuard = null; // função async que decide se é possível sair da tela atual
let guarding = false; // a pergunta do leaveGuard está aberta

async function route() {
  // Tela com alterações não salvas: pergunta antes de sair (e volta a URL se o usuário desistir).
  if (leaveGuard && location.href !== currentUrl) {
    const guard = leaveGuard;
    leaveGuard = null;
    const started = navigation;
    guarding = true;
    let leave;
    try {
      leave = await guard();
    } finally {
      guarding = false;
    }
    // Outra navegação aconteceu com a pergunta aberta (ex.: "Voltar"): ela decide a tela.
    if (navigation !== started) return;
    if (!leave) {
      leaveGuard = guard;
      history.replaceState(null, '', currentUrl);
      return;
    }
  }
  currentUrl = location.href;
  // Um link dentro de um diálogo (ex.: o relatório no Histórico) leva a outra tela: o diálogo fecha
  // (menos a pergunta sobre alterações não salvas, que continua esperando a resposta).
  const modal = document.getElementById('modal');
  if (modal?.open && !guarding) modal.close();
  const token = ++navigation;
  const hash = location.hash.replace(/^#/, '') || '/';
  const [path, query = ''] = hash.split('?');
  const match = ROUTES.map(([re, mod, nav, props]) => [re.exec(path), mod, nav, props]).find(([m]) => m);
  if (cleanup) {
    try {
      cleanup();
    } catch {
      // ignora
    }
    cleanup = null;
  }
  leaveGuard = null;
  setActiveNav(match?.[2]);
  document.getElementById('tooltip').hidden = true;
  // Cada navegação desenha em um contêiner novo: respostas atrasadas da tela anterior vão para um
  // elemento que já saiu da página e não aparecem por cima da tela atual.
  const container = document.createElement('div');
  view.replaceChildren(container);
  if (!match) {
    render(container, html`<div class="empty"><h1>Página não encontrada</h1><p><a href="#/">Voltar ao painel</a></p></div>`);
    return;
  }
  const [m, mod, , props = {}] = match;
  render(container, html`<p class="loading">Carregando…</p>`);
  const isCurrent = () => token === navigation;
  try {
    await loadInfo();
    const result = await mod.render(container, {
      params: m.slice(1),
      props,
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
    const i = await loadInfo();
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
