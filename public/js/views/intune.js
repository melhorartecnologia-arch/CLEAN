// Reinício agendado dos iPhone/iPad gerenciados no Intune (Microsoft Graph): lista os agendamentos,
// permite executar agora (ou simular), pausar/retomar e ver o histórico com o resultado por aparelho.
import { get, post, del } from '../api.js';
import { html, render as paint, icon, toast, confirmDialog, openDialog, fmtNum, fmtServerDateTime, redraw } from '../ui.js';
import { zoneNote } from '../schedule-form.js';

const TRIGGER = { schedule: 'No horário', 'catch-up': 'Atrasada', manual: 'Manual' };
const STATE = {
  active: html`<span class="badge completed">Ativo</span>`,
  paused: html`<span class="badge cancelled">Pausado</span>`,
  finished: html`<span class="badge">Encerrado</span>`,
};
const RUN_BADGE = {
  ok: html`<span class="badge completed">Concluído</span>`,
  partial: html`<span class="badge interrupted">Parcial</span>`,
  failed: html`<span class="badge failed">Falhou</span>`,
  skipped: html`<span class="badge cancelled">Pulada</span>`,
  missed: html`<span class="badge interrupted">Perdida</span>`,
};
const DEVICE_BADGE = {
  ok: html`<span class="badge completed">Reinício enviado</span>`,
  failed: html`<span class="badge failed">Falha</span>`,
  skipped: html`<span class="badge cancelled">Ignorado</span>`,
};

const runBadge = (status) => RUN_BADGE[status] || html`<span class="badge">${status}</span>`;

/** Resumo de um escopo (todos os aparelhos ou uma lista). */
function scopeText(job) {
  const s = job.scope || {};
  if (s.mode === 'list') return `${fmtNum((s.names || []).length)} aparelho(s) na lista`;
  return 'Todos os iPhone/iPad gerenciados';
}

export async function render(root, { ctx }) {
  let jobs = [];
  let timer = null;
  let stopped = false;

  const row = (j) => {
    const last = j.lastRun;
    return html`<tr data-id="${j.id}">
      <td>
        <a href="#/intune/${j.id}" data-action="open"><b>${j.name}</b></a>
        <div class="muted small">${j.source.name}${j.source.missing ? ' (conexão excluída)' : ''} · ${scopeText(j)}</div>
        ${j.problems.map((p) => html`<div class="small danger-text">${icon('alert')} ${p}</div>`)}
      </td>
      <td>
        ${j.description}
        ${j.remaining !== null && j.state === 'active' ? html`<div class="muted small">${j.remaining === 1 ? 'Falta 1 execução' : `Faltam ${fmtNum(j.remaining)} execuções`}</div>` : ''}
      </td>
      <td class="nowrap">${j.state === 'active' ? fmtServerDateTime(j.nextRunAt) : STATE[j.state]}</td>
      <td>
        ${last
          ? html`${runBadge(last.status)}
              <div class="muted small nowrap">${fmtServerDateTime(last.at, { weekday: false })}${last.trigger !== 'schedule' ? ` · ${(TRIGGER[last.trigger] || last.trigger).toLowerCase()}` : ''}</div>
              ${last.message ? html`<div class="small muted clamp">${last.message}</div>` : ''}`
          : html`<span class="muted">Ainda não executado</span>`}
      </td>
      <td class="actions">
        <button class="icon-btn" data-action="simulate" aria-label="Simular ${j.name}" title="Simular (só lista os aparelhos)">${icon('flask')}</button>
        <button class="icon-btn" data-action="run" aria-label="Reiniciar agora: ${j.name}" title="Reiniciar agora">${icon('play')}</button>
        ${j.enabled
          ? html`<button class="icon-btn" data-action="pause" aria-label="Pausar ${j.name}" title="Pausar">${icon('pause')}</button>`
          : html`<button class="icon-btn" data-action="resume" aria-label="Retomar ${j.name}" title="Retomar">${icon('refresh')}</button>`}
        <button class="icon-btn" data-action="history" aria-label="Histórico de ${j.name}" title="Histórico">${icon('history')}</button>
        <a class="icon-btn" href="#/intune/${j.id}" data-action="edit" aria-label="Editar ${j.name}" title="Editar">${icon('edit')}</a>
        <button class="icon-btn danger" data-action="delete" aria-label="Excluir ${j.name}" title="Excluir">${icon('trash')}</button>
      </td>
    </tr>`;
  };

  paint(
    root,
    html`<div class="page-head">
        <div>
          <h1>Reinício de iPhones (Intune)</h1>
          <div class="sub">Reinicia os aparelhos iPhone/iPad gerenciados no Microsoft Intune, de forma agendada e recorrente, pela API Microsoft Graph — com o registro de sucesso ou falha por aparelho.</div>
        </div>
        <div class="actions">
          <a class="btn primary" href="#/intune/novo">${icon('plus')} Novo agendamento</a>
        </div>
      </div>
      <div class="alert info">${icon('info')}<div>${zoneNote(ctx.info)} O reinício só acontece com o CLEAN em execução nesses horários. A conexão do Microsoft 365 usada precisa das permissões <span class="mono">DeviceManagementManagedDevices.Read.All</span> e <span class="mono">DeviceManagementManagedDevices.PrivilegedOperations.All</span>. O comando de reinício é <b>enfileirado pelo Intune</b> e entregue na próxima comunicação do aparelho: um aparelho <b>offline</b> (sem sincronizar com o Intune) só reinicia quando voltar a se conectar — o resultado de cada execução destaca esses casos.</div></div>
      <section class="card" data-list></section>`,
  );
  const list = root.querySelector('[data-list]');

  const draw = () => {
    const scroll = list.querySelector('.table-wrap')?.scrollLeft || 0;
    redraw(list, () =>
      paint(
        list,
        jobs.length === 0
          ? html`<div class="empty">
              <p>Nenhum agendamento de reinício ainda. Crie um para reiniciar os iPhones automaticamente, por exemplo toda madrugada.</p>
              <div class="inline"><a class="btn primary" href="#/intune/novo">${icon('plus')} Novo agendamento</a></div>
            </div>`
          : html`<div class="table-wrap">
              <table class="data">
                <thead>
                  <tr><th>Agendamento</th><th>Quando</th><th>Próxima execução</th><th>Última execução</th><th><span class="sr-only">Ações</span></th></tr>
                </thead>
                <tbody>${jobs.map(row)}</tbody>
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
      const latest = await get('/api/intune');
      if (stopped) return;
      jobs = latest;
      draw();
    } finally {
      if (!stopped) {
        clearTimeout(timer);
        const soon = jobs.some((j) => j.nextRunAt && Date.parse(j.nextRunAt) - Date.now() < 90000);
        timer = setTimeout(() => refresh().catch(() => {}), soon ? 5000 : 30000);
      }
    }
  };

  /** Última sincronização do aparelho com o Intune, com aviso de provável offline. */
  const syncCell = (d) =>
    html`${d.lastSync ? fmtServerDateTime(d.lastSync, { weekday: false }) : html`<span class="muted">nunca</span>`}${d.stale ? html`<div class="small warn-text">${icon('alert')} provavelmente offline</div>` : ''}`;

  /** Tabela do resultado por aparelho de uma execução. */
  const devicesTable = (run) =>
    (run.devices || []).length
      ? html`<div class="table-wrap">
          <table class="data compact">
            <thead><tr><th>Aparelho</th><th>Usuário</th><th>Sistema</th><th>Última sincronização</th><th>Resultado</th></tr></thead>
            <tbody>
              ${run.devices.map(
                (d) => html`<tr>
                  <td><b>${d.name}</b>${d.serialNumber ? html`<div class="mono small muted">${d.serialNumber}</div>` : ''}</td>
                  <td>${d.userName || d.user || html`<span class="muted">—</span>`}</td>
                  <td class="nowrap">${d.os || ''}${d.osVersion ? ` ${d.osVersion}` : ''}</td>
                  <td class="nowrap">${syncCell(d)}</td>
                  <td>${DEVICE_BADGE[d.status] || d.status}${d.error ? html`<div class="small danger-text clamp">${d.error}</div>` : ''}</td>
                </tr>`,
              )}
            </tbody>
          </table>
          ${run.devicesTruncated ? html`<p class="muted small">Só os primeiros ${fmtNum(run.devices.length)} aparelhos são mostrados.</p>` : ''}
        </div>`
      : html`<p class="muted small">Nenhum aparelho no resultado desta execução.</p>`;

  /** Detalhes de uma execução (resumo + resultado por aparelho). */
  const showRun = (job, run) =>
    openDialog({
      title: `Execução – ${job.name}`,
      wide: true,
      describe: false,
      submitLabel: 'Fechar',
      cancelLabel: '',
      body: html`<div class="run-summary">
          ${runBadge(run.status)}
          <span class="muted small">${fmtServerDateTime(run.at)} · ${(TRIGGER[run.trigger] || run.trigger)}${run.by ? ` · ${run.by}` : ''}${run.simulate ? ' · simulação' : ''}</span>
        </div>
        <p>${run.message || ''}</p>
        ${run.total ? html`<p class="muted small">${fmtNum(run.ok || 0)} reiniciado(s), ${fmtNum(run.failed || 0)} com falha, ${fmtNum(run.skipped || 0)} ignorado(s) de ${fmtNum(run.total)}.</p>` : ''}
        ${devicesTable(run)}`,
    });

  const showHistory = async (job) => {
    const full = await get(`/api/intune/${job.id}`);
    const history = full.history || [];
    // Os aparelhos de cada execução ficam em seções recolhíveis (sem abrir outra janela sobre esta).
    await openDialog({
      title: `Histórico – ${job.name}`,
      wide: true,
      describe: false,
      submitLabel: 'Fechar',
      cancelLabel: '',
      body: history.length
        ? html`<p class="muted small">${history.length === 1 ? 'A única execução até agora.' : `As últimas ${fmtNum(history.length)} execuções (as mais recentes primeiro).`} ${zoneNote(ctx.info)}</p>
            <div class="history-list">
              ${history.map(
                (h) => html`<div class="history-item">
                  <div class="history-head">
                    ${runBadge(h.status)}
                    <span class="muted small">${fmtServerDateTime(h.at, { weekday: false })} · ${(TRIGGER[h.trigger] || h.trigger)}${h.by ? ` · ${h.by}` : ''}${h.simulate ? ' · simulação' : ''}</span>
                  </div>
                  <div class="small">${h.message || ''}</div>
                  ${(h.devices || []).length ? html`<details class="run-devices"><summary>Ver ${fmtNum(h.devices.length)} aparelho(s)</summary>${devicesTable(h)}</details>` : ''}
                </div>`,
              )}
            </div>`
        : html`<p>O agendamento ainda não foi executado.</p>`,
    });
  };

  const onClick = async (event) => {
    const button = event.target.closest('button[data-action]');
    if (!button) return;
    const j = jobs.find((x) => x.id === button.closest('tr')?.dataset.id);
    if (!j) return;
    const action = button.dataset.action;
    try {
      if (action === 'history') return await showHistory(j);
      if (action === 'simulate') {
        const { run } = await post(`/api/intune/${j.id}/run`, { simulate: true });
        toast('Simulação concluída.', 'success');
        await refresh({ force: true });
        return await showRun(j, run);
      }
      if (action === 'run') {
        const message = `Reiniciar agora os aparelhos do agendamento "${j.name}"? Os iPhones/iPads do escopo serão reiniciados pelo Intune. A próxima execução programada não muda.`;
        if (!(await confirmDialog(message, { title: 'Reiniciar agora', confirmLabel: 'Reiniciar agora', danger: true }))) return;
        const { run } = await post(`/api/intune/${j.id}/run`, { confirm: true });
        toast(run.message || 'Reinício enviado.', run.status === 'failed' ? 'error' : 'success');
        await refresh({ force: true });
        return await showRun(j, run);
      }
      if (action === 'pause' || action === 'resume') {
        if (action === 'pause') {
          await post(`/api/intune/${j.id}/pause`);
          toast(`Agendamento "${j.name}" pausado.`);
        } else {
          const updated = await post(`/api/intune/${j.id}/resume`);
          toast(updated.nextRunAt ? `Agendamento retomado. Próxima execução: ${fmtServerDateTime(updated.nextRunAt)}.` : 'Agendamento retomado, mas a regra não tem mais execuções futuras.', updated.nextRunAt ? 'success' : 'info');
        }
        await refresh({ force: true });
        list.querySelector(`tr[data-id="${CSS.escape(j.id)}"] [data-action="${action === 'pause' ? 'resume' : 'pause'}"]`)?.focus();
        return;
      }
      if (action === 'delete') {
        if (!(await confirmDialog(`Excluir o agendamento de reinício "${j.name}"? O histórico dele é descartado.`, { confirmLabel: 'Excluir' }))) return;
        await del(`/api/intune/${j.id}`);
        toast('Agendamento excluído.', 'success');
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
