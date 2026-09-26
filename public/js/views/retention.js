// Políticas de retenção: listam (e, se configuradas para isso, excluem) os arquivos e as mensagens
// mais antigos que a idade máxima de cada política. Execução manual ("Simular agora" e "Executar
// agora") ou agendada.
import { get, post, del } from '../api.js';
import { html, render as paint, icon, toast, confirmDialog, openDialog, fmtNum, fmtServerDateTime, statusBadge, redraw, plural } from '../ui.js';
import { zoneNote } from '../schedule-form.js';
import { deletionModeText, deletionsText } from '../retention.js';

const KIND = { files: 'Arquivos', mail: 'E-mail' };
const TRIGGER = { schedule: 'No horário', 'catch-up': 'Atrasada', manual: 'Manual' };
const STATE = {
  active: html`<span class="badge completed">Ativa</span>`,
  paused: html`<span class="badge cancelled">Pausada</span>`,
  finished: html`<span class="badge">Encerrada</span>`,
  manual: html`<span class="badge">Manual</span>`,
};
const reportLink = (p, scanId) => `${p.kind === 'mail' ? '#/email/analises' : '#/analises'}/${scanId}`;
const words = (p) => (p.kind === 'mail' ? { o: 'a', items: 'mensagens' } : { o: 'o', items: 'arquivos' });

/** Forma e limite da exclusão: "exclusão definitiva, até 1.000 exclusões por execução". */
function deletionText(p) {
  const r = p.retention || {};
  const mode = deletionModeText(p.kind, r, (p.targets || []).map((t) => t.type || 'local'));
  return `${mode}, ${r.maxDeletions ? `até ${deletionsText(r.maxDeletions)} por execução` : 'sem limite por execução'}`;
}

/** Resultado de uma execução do histórico (situação da análise, pulada, falhou ou perdida). */
function runResult(p, h) {
  if (!h) return html`<span class="muted">—</span>`;
  if (h.status === 'skipped') return html`<span class="badge cancelled">Pulada</span>`;
  if (h.status === 'missed') return html`<span class="badge interrupted">Perdida</span>`;
  if (h.status === 'failed') return html`<span class="badge failed">Não iniciada</span>`;
  const status = h.scanStatus || h.outcome?.status;
  const badge = status && status !== 'removed' ? statusBadge(status) : html`<span class="badge">Iniciada</span>`;
  return h.reportExists ? html`<a href="${reportLink(p, h.scanId)}" class="result-link" data-action="report">${badge}<span class="sr-only"> (abrir o relatório)</span></a>` : badge;
}

function outcomeText(p, h) {
  const out = h.outcome;
  if (!out || out.status === 'removed') return '';
  const { o } = words(p);
  const parts = [p.kind === 'mail' ? plural(out.matched, 'mensagem expirada', 'mensagens expiradas') : plural(out.matched, 'arquivo expirado', 'arquivos expirados')];
  if (out.deleted) parts.push(plural(out.deleted, `excluíd${o}`, `excluíd${o}s`));
  if (out.errors) parts.push(`${fmtNum(out.errors)} ${out.errors === 1 ? 'erro' : 'erros'}`);
  return parts.join(' · ');
}

export async function render(root, { ctx }) {
  let policies = [];
  let timer = null;
  let stopped = false;

  const row = (p) => {
    const last = p.lastRun;
    const deleting = p.action === 'delete';
    return html`<tr data-id="${p.id}">
      <td>
        <a href="#/retencao/${p.id}" data-action="open"><b>${p.name}</b></a>
        <span class="kind-badge">${KIND[p.kind]}</span>
        ${deleting ? html`<span class="chip danger">exclui ${p.kind === 'mail' ? 'as expiradas' : 'os expirados'}</span>` : html`<span class="chip">só lista (simulação)</span>`}
        <div class="muted small">${p.targets.map((t) => t.name).join(', ')}</div>
        ${p.problems.map((problem) => html`<div class="small danger-text">${icon('alert')} ${problem}</div>`)}
      </td>
      <td>
        ${p.retentionText}
        ${deleting ? html`<div class="muted small">${deletionText(p)}</div>` : ''}
      </td>
      <td>
        ${p.description}
        <div class="small nowrap">${p.state === 'active' ? html`Próxima: <b>${fmtServerDateTime(p.nextRunAt)}</b>` : STATE[p.state]}</div>
        ${p.running ? html`<div class="small"><span class="badge running">Em andamento</span></div>` : ''}
      </td>
      <td>
        ${last
          ? html`${runResult(p, last)}${last.simulated ? html` <span class="chip">simulação</span>` : ''}
              <div class="muted small nowrap">${fmtServerDateTime(last.at, { weekday: false })}${last.trigger !== 'schedule' ? ` · ${TRIGGER[last.trigger].toLowerCase()}` : ''}</div>
              ${outcomeText(p, last) ? html`<div class="muted small">${outcomeText(p, last)}</div>` : ''}
              ${last.status === 'failed' || last.status === 'skipped' || last.status === 'missed' ? html`<div class="small muted clamp">${last.message}</div>` : ''}`
          : html`<span class="muted">Ainda não executada</span>`}
      </td>
      <td class="actions">
        <button class="icon-btn" data-action="simulate" aria-label="Simular agora: ${p.name}" title="Simular agora (nada é excluído)" ${p.running ? 'disabled' : ''}>${icon('flask')}</button>
        ${deleting ? html`<button class="icon-btn danger" data-action="run" aria-label="Executar e excluir agora: ${p.name}" title="Executar e excluir agora" ${p.running ? 'disabled' : ''}>${icon('play')}</button>` : ''}
        ${p.rule
          ? p.enabled
            ? html`<button class="icon-btn" data-action="pause" aria-label="Pausar ${p.name}" title="Pausar">${icon('pause')}</button>`
            : html`<button class="icon-btn" data-action="resume" aria-label="Retomar ${p.name}" title="Retomar">${icon('refresh')}</button>`
          : ''}
        <button class="icon-btn" data-action="history" aria-label="Histórico de ${p.name}" title="Histórico">${icon('history')}</button>
        <a class="icon-btn" href="#/retencao/${p.id}" data-action="edit" aria-label="Editar ${p.name}" title="Editar">${icon('edit')}</a>
        <button class="icon-btn danger" data-action="delete" aria-label="Excluir ${p.name}" title="Excluir">${icon('trash')}</button>
      </td>
    </tr>`;
  };

  paint(
    root,
    html`<div class="page-head">
        <div>
          <h1>Políticas de retenção</h1>
          <div class="sub">Eliminam os arquivos e as mensagens mais antigos que a idade máxima de cada política, pela data da última modificação, do último acesso, da criação ou do recebimento.</div>
        </div>
        <div class="actions">
          <a class="btn primary" href="#/retencao/nova">${icon('plus')} Política de arquivos</a>
          <a class="btn primary" href="#/retencao/nova?tipo=email">${icon('plus')} Política de e-mail</a>
        </div>
      </div>
      <div class="alert info">${icon('info')}<div>
        <b>Simule antes de excluir:</b> "Simular agora" gera o relatório com o que a política excluiria, sem excluir nada.
        ${zoneNote(ctx.info)} As políticas agendadas só são executadas com o CLEAN em execução (instale-o como serviço do Windows).
      </div></div>
      <section class="card" data-list></section>`,
  );
  const list = root.querySelector('[data-list]');

  const draw = () => {
    const scroll = list.querySelector('.table-wrap')?.scrollLeft || 0;
    redraw(list, () =>
      paint(
        list,
        policies.length === 0
          ? html`<div class="empty">
              <p>Nenhuma política ainda. Crie uma para eliminar, por exemplo, os arquivos sem uso há mais de 5 anos ou os e-mails com mais de 10 anos.</p>
              <div class="inline">
                <a class="btn primary" href="#/retencao/nova" data-action="new-files">${icon('plus')} Política de arquivos</a>
                <a class="btn" href="#/retencao/nova?tipo=email" data-action="new-mail">${icon('plus')} Política de e-mail</a>
              </div>
            </div>`
          : html`<div class="table-wrap">
              <table class="data schedules">
                <thead>
                  <tr><th>Política</th><th>Regra</th><th>Quando</th><th>Última execução</th><th><span class="sr-only">Ações</span></th></tr>
                </thead>
                <tbody>${policies.map(row)}</tbody>
              </table>
            </div>`,
      ),
    );
    const wrap = list.querySelector('.table-wrap');
    if (wrap) wrap.scrollLeft = scroll;
  };

  const dialogOpen = () => Boolean(document.getElementById('modal')?.open);

  const refresh = async ({ force = false } = {}) => {
    clearTimeout(timer);
    try {
      if (!force && dialogOpen()) return;
      const latest = await get('/api/schedules?purpose=retention');
      if (stopped) return;
      policies = latest;
      draw();
    } finally {
      if (!stopped) {
        clearTimeout(timer);
        const soon = policies.some((p) => p.running || (p.nextRunAt && Date.parse(p.nextRunAt) - Date.now() < 90000));
        timer = setTimeout(() => refresh().catch(() => {}), soon ? 4000 : 30000);
      }
    }
  };

  const showHistory = async (p) => {
    const full = await get(`/api/schedules/${p.id}`);
    const history = full.history || [];
    await openDialog({
      title: `Histórico – ${p.name}`,
      wide: true,
      describe: false,
      submitLabel: 'Fechar',
      cancelLabel: '',
      body: history.length
        ? html`<p class="muted small">${history.length === 1 ? 'A única execução até agora.' : `As últimas ${fmtNum(history.length)} execuções (as mais recentes primeiro).`} ${zoneNote(ctx.info)}</p>
            <div class="table-wrap">
              <table class="data">
                <thead><tr><th>Quando</th><th>Origem</th><th>Resultado</th><th>Detalhes</th></tr></thead>
                <tbody>
                  ${history.map(
                    (h) => html`<tr>
                      <td class="nowrap">${fmtServerDateTime(h.at, { weekday: false })}${h.plannedFor && Math.abs(Date.parse(h.plannedFor) - Date.parse(h.at)) > 120000 ? html`<div class="muted small">previsto: ${fmtServerDateTime(h.plannedFor, { weekday: false })}</div>` : ''}</td>
                      <td>${TRIGGER[h.trigger] || h.trigger}${h.simulated ? ' (simulação)' : ''}${h.by ? html`<div class="muted small">${h.by}</div>` : ''}</td>
                      <td>${runResult(full, h)}${h.status === 'started' && !h.reportExists ? html`<div class="muted small">relatório excluído</div>` : ''}</td>
                      <td class="small">${outcomeText(full, h) ? html`<div>${outcomeText(full, h)}</div>` : ''}${h.status !== 'started' && h.message ? html`<div>${h.message}</div>` : ''}</td>
                    </tr>`,
                  )}
                </tbody>
              </table>
            </div>`
        : html`<p>A política ainda não foi executada.</p>`,
    });
  };

  const onClick = async (event) => {
    const button = event.target.closest('button[data-action]');
    if (!button) return;
    const p = policies.find((x) => x.id === button.closest('tr')?.dataset.id);
    if (!p) return;
    const action = button.dataset.action;
    const { o, items } = words(p);
    try {
      if (action === 'history') return await showHistory(p);
      if (action === 'simulate' || action === 'run') {
        const simulate = action === 'simulate';
        const message = simulate
          ? `Simular agora a política "${p.name}"? O relatório lista ${o}s ${items} que ${p.action === 'delete' ? 'seriam excluíd' : 'estão expirad'}${o}s; nada é excluído.`
          : `Executar agora a política "${p.name}"? ${o === 'a' ? 'As' : 'Os'} ${items} expirad${o}s serão EXCLUÍD${o === 'a' ? 'A' : 'O'}S: ${deletionText(p)}. Na dúvida, simule antes.`;
        const ok = await confirmDialog(message, { title: simulate ? 'Simular agora' : 'Executar e excluir agora', confirmLabel: simulate ? 'Simular' : 'Executar e excluir', danger: !simulate });
        if (!ok) return;
        const { scan } = await post(`/api/schedules/${p.id}/run`, simulate ? { simulate: true } : { confirm: true });
        toast(simulate ? 'Simulação iniciada.' : 'Execução iniciada.', 'success');
        location.hash = reportLink(p, scan.id);
        return;
      }
      if (action === 'pause' || action === 'resume') {
        if (action === 'pause') {
          await post(`/api/schedules/${p.id}/pause`);
          toast(`Política "${p.name}" pausada.`);
        } else {
          const updated = await post(`/api/schedules/${p.id}/resume`);
          toast(updated.nextRunAt ? `Política retomada. Próxima execução: ${fmtServerDateTime(updated.nextRunAt)}.` : 'Política retomada, mas a regra não tem mais execuções futuras.', updated.nextRunAt ? 'success' : 'info');
        }
        await refresh({ force: true });
        list.querySelector(`tr[data-id="${CSS.escape(p.id)}"] [data-action="${action === 'pause' ? 'resume' : 'pause'}"]`)?.focus();
        return;
      }
      if (action === 'delete') {
        if (!(await confirmDialog(`Excluir a política "${p.name}"? Os relatórios já gerados são mantidos.`, { confirmLabel: 'Excluir' }))) return;
        await del(`/api/schedules/${p.id}`);
        toast('Política excluída.', 'success');
      }
      await refresh({ force: true });
    } catch (err) {
      toast(err.message, 'error');
      await refresh({ force: true }).catch(() => {});
    }
  };

  try {
    await refresh({ force: true });
  } catch (err) {
    stopped = true;
    clearTimeout(timer);
    throw err;
  }
  root.addEventListener('click', onClick);
  return () => {
    stopped = true;
    clearTimeout(timer);
    root.removeEventListener('click', onClick);
  };
}
