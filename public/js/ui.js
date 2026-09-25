// Utilitários da interface: HTML com escape automático, ícones, formatação, diálogos e avisos.

export class SafeHtml {
  constructor(value) {
    this.value = value;
  }
  toString() {
    return this.value;
  }
}

const ESCAPES = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' };

export function esc(value) {
  return String(value ?? '').replace(/[&<>"']/g, (c) => ESCAPES[c]);
}

function part(value) {
  if (value === null || value === undefined || value === false) return '';
  if (value instanceof SafeHtml) return value.value;
  if (Array.isArray(value)) return value.map(part).join('');
  return esc(value);
}

/** Template de HTML: tudo que é interpolado é escapado, exceto SafeHtml e listas de SafeHtml. */
export function html(strings, ...values) {
  let out = '';
  strings.forEach((s, i) => {
    out += s;
    if (i < values.length) out += part(values[i]);
  });
  return new SafeHtml(out);
}

export const raw = (value) => new SafeHtml(String(value));

/** Larguras de barras via CSSOM (a política de segurança bloqueia atributos style). */
export function applyWidths(root) {
  root.querySelectorAll('[data-w]').forEach((el) => {
    el.style.width = `${el.dataset.w}%`;
  });
}

export function render(el, content) {
  el.innerHTML = part(content);
  applyWidths(el);
}

// ---------- Ícones ----------

const ICONS = {
  plus: '<path d="M12 5v14M5 12h14"/>',
  x: '<path d="M18 6 6 18M6 6l12 12"/>',
  trash: '<path d="M4 7h16M9 7V4h6v3M6 7l1 13h10l1-13"/>',
  edit: '<path d="M4 20h4L20 8l-4-4L4 16v4zM13.5 6.5l4 4"/>',
  copy: '<rect x="9" y="9" width="11" height="11" rx="2"/><path d="M5 15V6a2 2 0 0 1 2-2h9"/>',
  chevron: '<path d="m9 6 6 6-6 6"/>',
  download: '<path d="M12 4v11m0 0-4-4m4 4 4-4M5 20h14"/>',
  upload: '<path d="M12 16V5m0 0-4 4m4-4 4 4M5 20h14"/>',
  search: '<circle cx="11" cy="11" r="7"/><path d="m20 20-3.5-3.5"/>',
  folder: '<path d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/>',
  list: '<path d="M9 6h11M9 12h11M9 18h11M4.5 6h.01M4.5 12h.01M4.5 18h.01"/>',
  play: '<path d="M7 4v16l13-8z"/>',
  stop: '<rect x="6" y="6" width="12" height="12" rx="1.5"/>',
  check: '<path d="m5 12 5 5 9-10"/>',
  alert: '<path d="M12 3 2 20h20L12 3zM12 10v4M12 17h.01"/>',
  info: '<circle cx="12" cy="12" r="9"/><path d="M12 11v6M12 7.5h.01"/>',
  refresh: '<path d="M20 11a8 8 0 1 0-2.3 5.7M20 4v7h-7"/>',
  file: '<path d="M6 3h8l5 5v13H6zM14 3v5h5"/>',
  user: '<circle cx="12" cy="8" r="4"/><path d="M4 21a8 8 0 0 1 16 0"/>',
  flask: '<path d="M9 3h6M10 3v6L4 19a1.5 1.5 0 0 0 1.3 2h13.4a1.5 1.5 0 0 0 1.3-2L14 9V3"/>',
  clock: '<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/>',
  pause: '<rect x="6" y="5" width="4" height="14" rx="1"/><rect x="14" y="5" width="4" height="14" rx="1"/>',
  history: '<path d="M3 12a9 9 0 1 0 3-6.7L3 8M3 3v5h5M12 7v5l3 2"/>',
};

export function icon(name, label = '') {
  const a11y = label ? `role="img" aria-label="${esc(label)}"` : 'aria-hidden="true"';
  return raw(
    `<svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" ${a11y}>${ICONS[name] || ''}</svg>`,
  );
}

// ---------- Formatação ----------

const numberFormat = new Intl.NumberFormat('pt-BR');
const compactFormat = new Intl.NumberFormat('pt-BR', { notation: 'compact', maximumFractionDigits: 1 });

export const fmtNum = (n) => numberFormat.format(Number(n) || 0);
export const fmtCompact = (n) => (Math.abs(Number(n) || 0) < 10000 ? fmtNum(n) : compactFormat.format(Number(n)));

export function fmtDateTime(iso) {
  if (!iso) return '—';
  const d = new Date(iso);
  return Number.isNaN(d.getTime()) ? '—' : d.toLocaleString('pt-BR', { dateStyle: 'short', timeStyle: 'short' });
}

let serverTimeZone = '';

/** Fuso do servidor: os horários dos agendamentos são mostrados nele (e não no do navegador). */
export function setServerTimeZone(timeZone) {
  serverTimeZone = timeZone || '';
}

/** Data e hora no fuso do servidor, com o dia da semana (ex.: "sex., 25/09/2026, 22:00"). */
export function fmtServerDateTime(iso, { weekday = true } = {}) {
  if (!iso) return '—';
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return '—';
  const options = { ...(weekday ? { weekday: 'short' } : {}), day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' };
  try {
    return d.toLocaleString('pt-BR', { ...options, timeZone: serverTimeZone || undefined });
  } catch {
    return d.toLocaleString('pt-BR', options); // fuso desconhecido pelo navegador
  }
}

export function fmtDate(iso) {
  if (!iso) return '—';
  const d = new Date(iso);
  return Number.isNaN(d.getTime()) ? '—' : d.toLocaleDateString('pt-BR');
}

export function fmtBytes(bytes) {
  let value = Number(bytes) || 0;
  const units = ['B', 'KB', 'MB', 'GB', 'TB'];
  let i = 0;
  while (value >= 1024 && i < units.length - 1) {
    value /= 1024;
    i++;
  }
  return `${value.toLocaleString('pt-BR', { maximumFractionDigits: i ? 1 : 0 })} ${units[i]}`;
}

export function fmtDuration(ms) {
  const s = Math.max(0, Math.round(ms / 1000));
  if (s < 60) return `${s}s`;
  const m = Math.floor(s / 60);
  if (m < 60) return `${m}min ${String(s % 60).padStart(2, '0')}s`;
  const h = Math.floor(m / 60);
  return `${h}h ${String(m % 60).padStart(2, '0')}min`;
}

export function plural(n, one, many) {
  return `${fmtNum(n)} ${Number(n) === 1 ? one : many}`;
}

export const SCAN_STATUS = {
  queued: 'Na fila',
  running: 'Em andamento',
  completed: 'Concluída',
  cancelled: 'Cancelada',
  failed: 'Falhou',
  interrupted: 'Interrompida',
};

export function statusBadge(status) {
  return html`<span class="badge ${status}">${SCAN_STATUS[status] || status}</span>`;
}

// ---------- Avisos, diálogos e dicas ----------

export function toast(message, kind = 'info') {
  const box = document.getElementById('toasts');
  const el = document.createElement('div');
  el.className = `toast ${kind}`;
  el.setAttribute('role', kind === 'error' ? 'alert' : 'status');
  el.textContent = message;
  box.append(el);
  setTimeout(() => el.remove(), kind === 'error' ? 8000 : 4500);
}

/**
 * Abre um formulário em diálogo modal. onSubmit(form) pode lançar erro (exibido no diálogo) ou
 * devolver false para manter o diálogo aberto.
 */
export function openDialog({ title, body, submitLabel = 'Salvar', cancelLabel = 'Cancelar', wide = false, danger = false, onOpen, onSubmit }) {
  const dialog = document.getElementById('modal');
  dialog.className = wide ? 'wide' : '';
  // Leitores de tela anunciam o título e o texto ao abrir (importante nas confirmações de exclusão).
  dialog.setAttribute('aria-labelledby', 'modal-title');
  dialog.setAttribute('aria-describedby', 'modal-body');
  render(
    dialog,
    html`<form class="dialog" novalidate>
      <header>
        <h2 id="modal-title">${title}</h2>
        <button type="button" class="icon-btn" data-close aria-label="Fechar">${icon('x')}</button>
      </header>
      <div class="dialog-body" id="modal-body">${body}</div>
      <div class="dialog-error" hidden></div>
      <footer>
        ${cancelLabel ? html`<button type="button" class="btn" data-close>${cancelLabel}</button>` : ''}
        <button type="submit" class="btn ${danger ? 'danger' : 'primary'}">${submitLabel}</button>
      </footer>
    </form>`,
  );
  const form = dialog.querySelector('form');
  const errorBox = dialog.querySelector('.dialog-error');
  return new Promise((resolve) => {
    let result = null;
    const close = () => dialog.close();
    dialog.querySelectorAll('[data-close]').forEach((b) => b.addEventListener('click', close));
    dialog.addEventListener(
      'close',
      () => {
        resolve(result);
      },
      { once: true },
    );
    form.addEventListener('submit', async (event) => {
      event.preventDefault();
      errorBox.hidden = true;
      const submit = form.querySelector('[type="submit"]');
      submit.disabled = true;
      try {
        const value = onSubmit ? await onSubmit(form) : true;
        if (value !== false) {
          result = value ?? true;
          close();
        }
      } catch (err) {
        render(errorBox, html`<div class="alert error">${icon('alert')}<div>${err.message}</div></div>`);
        errorBox.hidden = false;
      } finally {
        submit.disabled = false;
      }
    });
    dialog.showModal();
    if (onOpen) onOpen(form);
    // Foco no primeiro campo; nas confirmações (sem campos), no botão que não faz nada.
    const first = form.querySelector('input, select, textarea') || form.querySelector('footer [data-close]');
    if (first) first.focus();
  });
}

export async function confirmDialog(message, { title = 'Confirmar', confirmLabel = 'Confirmar', danger = true } = {}) {
  const result = await openDialog({ title, body: html`<p>${message}</p>`, submitLabel: confirmLabel, danger });
  return Boolean(result);
}

/** Dicas flutuantes para elementos com data-tip-value / data-tip-label (texto não confiável: textContent). */
export function bindTooltips(root) {
  const tip = document.getElementById('tooltip');
  const show = (el, x, y) => {
    tip.replaceChildren();
    const strong = document.createElement('strong');
    strong.textContent = el.dataset.tipValue || '';
    const span = document.createElement('span');
    span.textContent = el.dataset.tipLabel || '';
    tip.append(strong, span);
    tip.hidden = false;
    const rect = tip.getBoundingClientRect();
    const left = Math.min(x + 14, window.innerWidth - rect.width - 8);
    const top = y + 16 + rect.height > window.innerHeight ? y - rect.height - 10 : y + 16;
    tip.style.left = `${Math.max(8, left)}px`;
    tip.style.top = `${Math.max(8, top)}px`;
  };
  const hide = () => {
    tip.hidden = true;
  };
  root.addEventListener('pointermove', (e) => {
    const el = e.target.closest('[data-tip-value]');
    if (el && root.contains(el)) show(el, e.clientX, e.clientY);
    else hide();
  });
  root.addEventListener('pointerleave', hide);
  root.addEventListener('focusin', (e) => {
    const el = e.target.closest('[data-tip-value]');
    if (!el) return;
    const r = el.getBoundingClientRect();
    show(el, r.left + r.width / 2, r.bottom - 8);
  });
  root.addEventListener('focusout', hide);
}

/** Copia texto (com alternativa para páginas servidas por HTTP na rede, onde a API de área de transferência não existe). */
export async function copyText(value) {
  if (navigator.clipboard && window.isSecureContext) {
    await navigator.clipboard.writeText(value);
    return;
  }
  const area = document.createElement('textarea');
  area.value = value;
  area.setAttribute('readonly', '');
  area.className = 'sr-only';
  document.body.append(area);
  area.select();
  document.execCommand('copy');
  area.remove();
}

/** Adia a chamada até `ms` sem novas chamadas. A função devolvida tem .cancel(). */
export function debounce(fn, ms = 300) {
  let timer;
  const debounced = (...args) => {
    clearTimeout(timer);
    timer = setTimeout(() => fn(...args), ms);
  };
  debounced.cancel = () => clearTimeout(timer);
  return debounced;
}

function focusKey(el, root) {
  if (!el || el === document.body || !root.contains(el)) return null;
  const attrs = ['data-action', 'data-page', 'data-filter-key', 'data-filter-value', 'data-chart-view', 'data-tab', 'name'];
  let selector = attrs
    .filter((a) => el.hasAttribute(a))
    .map((a) => `[${a}="${CSS.escape(el.getAttribute(a))}"]`)
    .join('');
  if (!selector) return null;
  const row = el.closest('tr[data-id]');
  if (row) selector = `tr[data-id="${CSS.escape(row.dataset.id)}"] ${selector}`;
  const chart = el.closest('[data-chart]');
  if (chart) selector = `[data-chart="${CSS.escape(chart.dataset.chart)}"] ${selector}`;
  return selector;
}

/** Redesenha um bloco devolvendo o foco ao mesmo controle (ex.: durante atualizações automáticas). */
export function redraw(root, draw) {
  const key = focusKey(document.activeElement, root);
  draw();
  if (key) root.querySelector(key)?.focus({ preventScroll: true });
}

export function formData(form) {
  return Object.fromEntries(new FormData(form).entries());
}
