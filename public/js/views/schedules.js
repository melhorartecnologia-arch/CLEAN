// Agendamentos: análises executadas automaticamente pelas regras de recorrência.
import { get, post, del } from '../api.js';
import { html, render as paint, icon, toast, confirmDialog, openDialog, fmtNum, fmtServerDateTime, statusBadge, redraw } from '../ui.js';
import { zoneNote } from '../schedule-form.js';

const KIND = { files: 'Arquivos', mail: 'E-mail' };
const TRIGGER = { schedule: 'No horário', 'catch-up': 'Atrasada', manual: 'Manual' };
const STATE = {
  active: html`<span class="badge completed">Ativo</span>`,
  paused: html`<span class="badge cancelled">Pausado</span>`,
  finished: html`<span class="badge">Encerrado</span>`,
};
const reportLink = (s, scanId) => `${s.kind === 'mail' ? '#/email/analises' : '#/analises'}/${scanId}`;

function periodText(s) {
  const p = s.period || {};
  const items = s.kind === 'mail' ? 'mensagens recebidas' : 'arquivos alterados';
  if (p.type === 'days') return `Somente ${items} nos últimos ${fmtNum(p.days)} dias`;
  if (p.type === 'since-last') return `Incremental (${items} desde a execução anterior)${p.fullEvery ? `; completa a cada ${fmtNum(p.fullEvery)} execuções` : ''}`;
  return s.kind === 'mail' ? 'Todas as mensagens' : 'Todos os arquivos';
}

/** Resultado de uma execução do histórico (situação da análise, pulada, falhou ou perdida). */
function runResult(s, h) {
  if (!h) return html`<span class="muted">—</span>`;
  if (h.status === 'skipped') return html`<span class="badge cancelled">Pulada</span>`;
  if (h.status === 'missed') return html`<span class="badge interrupted">Perdida</span>`;
  if (h.status === 'failed') return html`<span class="badge failed">Não iniciada</span>`;
  const status = h.scanStatus || h.outcome?.status;
  const badge = status && status !== 'removed' ? statusBadge(status) : html`<span class="badge">Iniciada</span>`;
  return h.reportExists ? html`<a href="${reportLink(s, h.scanId)}" class="result-link" data-action="report" aria-label="Abrir o relatório">${badge}</a>` : badge;
}

function outcomeText(s, h) {
  const o = h.outcome;
  if (!o || o.status === 'removed') return '';
  const parts = [`${fmtNum(o.matched)} ${s.kind === 'mail' ? (o.matched === 1 ? 'mensagem encontrada' : 'mensagens encontradas') : o.matched === 1 ? 'arquivo encontrado' : 'arquivos encontrados'}`];
  if (o.deleted) parts.push(`${fmtNum(o.deleted)} excluíd${s.kind === 'mail' ? 'a' : 'o'}${o.deleted === 1 ? '' : 's'}`);
  if (o.errors) parts.push(`${fmtNum(o.errors)} ${o.errors === 1 ? 'erro' : 'erros'}`);
  return parts.join(' · ');
}

export async function render(root, { ctx }) {
  let schedules = [];
  let timer = null;
  let stopped = false;

  const row = (s) => {
    const last = s.lastRun;
    return html`<tr data-id="${s.id}">
      <td>
        <a href="#/agendamentos/${s.id}" data-action="open"><b>${s.name}</b></a>
        <span class="kind-badge">${KIND[s.kind]}</span>
        ${s.action === 'delete' ? html`<span class="chip danger">exclusão automática</span>` : ''}
        <div class="muted small">${s.targets.map((t) => t.name).join(', ')} · ${s.lists.map((l) => l.name).join(', ')}</div>
        ${s.problems.map((p) => html`<div class="small danger-text">${icon('alert')} ${p}</div>`)}
      </td>
      <td>
        ${s.description}
        <div class="muted small">${periodText(s)}</div>
      </td>
      <td class="nowrap">
        ${s.state === 'active' ? html`${fmtServerDateTime(s.nextRunAt)}` : STATE[s.state]}
        ${s.running ? html`<div class="small"><span class="badge running">Em andamento</span></div>` : ''}
      </td>
      <td>
        ${last ? html`${runResult(s, last)}<div class="muted small nowrap">${fmtServerDateTime(last.at, { weekday: false })}${last.trigger !== 'schedule' ? ` · ${TRIGGER[last.trigger].toLowerCase()}` : ''}</div>${last.status === 'failed' || last.status === 'skipped' || last.status === 'missed' ? html`<div class="small muted clamp">${last.message}</div>` : ''}` : html`<span class="muted">Ainda não executado</span>`}
      </td>
      <td class="actions">
        <button class="icon-btn" data-action="run" aria-label="Executar agora: ${s.name}" title="Executar agora" ${s.running ? 'disabled' : ''}>${icon('play')}</button>
        ${s.enabled
          ? html`<button class="icon-btn" data-action="pause" aria-label="Pausar ${s.name}" title="Pausar">${icon('pause')}</button>`
          : html`<button class="icon-btn" data-action="resume" aria-label="Retomar ${s.name}" title="Retomar">${icon('refresh')}</button>`}
        <button class="icon-btn" data-action="history" aria-label="Histórico de ${s.name}" title="Histórico">${icon('history')}</button>
        <a class="icon-btn" href="#/agendamentos/${s.id}" data-action="edit" aria-label="Editar ${s.name}" title="Editar">${icon('edit')}</a>
        <button class="icon-btn danger" data-action="delete" aria-label="Excluir ${s.name}" title="Excluir">${icon('trash')}</button>
      </td>
    </tr>`;
  };

  // A página é desenhada uma vez; as atualizações redesenham só a lista, mantendo o foco e a
  // rolagem horizontal da tabela.
  paint(
    root,
    html`<div class="page-head">
        <div>
          <h1>Agendamentos</h1>
          <div class="sub">Análises executadas automaticamente nos dias e horários definidos, com os mesmos locais, listas e opções de uma análise comum.</div>
        </div>
        <div class="actions">
          <a class="btn primary" href="#/agendamentos/novo">${icon('plus')} Agendar arquivos</a>
          <a class="btn primary" href="#/agendamentos/novo?tipo=email">${icon('plus')} Agendar e-mails</a>
        </div>
      </div>
      <div class="alert info">${icon('info')}<div>${zoneNote(ctx.info)} Os agendamentos só são executados com o CLEAN em execução (instale-o como serviço do Windows). Se a execução anterior de um agendamento ainda estiver em andamento, a nova é pulada.</div></div>
      <section class="card" data-list></section>`,
  );
  const list = root.querySelector('[data-list]');

  const draw = () => {
    const scroll = list.querySelector('.table-wrap')?.scrollLeft || 0;
    redraw(list, () =>
      paint(
        list,
        schedules.length === 0
          ? html`<div class="empty">
              <p>Nenhum agendamento ainda. Agende uma análise para rodar sozinha, por exemplo toda noite ou uma vez por semana.</p>
              <div class="inline">
                <a class="btn primary" href="#/agendamentos/novo">${icon('plus')} Agendar arquivos</a>
                <a class="btn" href="#/agendamentos/novo?tipo=email">${icon('plus')} Agendar e-mails</a>
              </div>
            </div>`
          : html`<div class="table-wrap">
              <table class="data schedules">
                <thead>
                  <tr><th>Agendamento</th><th>Quando</th><th>Próxima execução</th><th>Última execução</th><th><span class="sr-only">Ações</span></th></tr>
                </thead>
                <tbody>${schedules.map(row)}</tbody>
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
      // Com um diálogo aberto (histórico, confirmação), a lista espera: o botão que o abriu continua
      // o mesmo e recebe o foco de volta ao fechar.
      if (!force && dialogOpen()) return;
      const latest = await get('/api/schedules');
      if (stopped) return;
      schedules = latest;
      draw();
    } finally {
      // Mais frequente enquanto alguma execução estiver em andamento ou prestes a começar; uma falha
      // (servidor reiniciando, rede) não interrompe as atualizações.
      if (!stopped) {
        clearTimeout(timer); // uma única atualização programada, mesmo com duas em andamento
        const soon = schedules.some((s) => s.running || (s.nextRunAt && Date.parse(s.nextRunAt) - Date.now() < 90000));
        timer = setTimeout(() => refresh().catch(() => {}), soon ? 4000 : 30000);
      }
    }
  };

  const showHistory = async (s) => {
    const full = await get(`/api/schedules/${s.id}`);
    const history = full.history || [];
    await openDialog({
      title: `Histórico – ${s.name}`,
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
                      <td>${TRIGGER[h.trigger] || h.trigger}${h.by ? html`<div class="muted small">${h.by}</div>` : ''}</td>
                      <td>${runResult(full, h)}${h.status === 'started' && !h.reportExists ? html`<div class="muted small">relatório excluído</div>` : ''}</td>
                      <td class="small">${h.status === 'started' ? html`<div>${h.full === false ? (h.from ? `Desde ${fmtServerDateTime(h.from, { weekday: false })}` : '') : 'Análise completa'}</div>` : ''}${outcomeText(full, h) ? html`<div>${outcomeText(full, h)}</div>` : ''}${h.status !== 'started' && h.message ? html`<div>${h.message}</div>` : ''}</td>
                    </tr>`,
                  )}
                </tbody>
              </table>
            </div>`
        : html`<p>O agendamento ainda não foi executado.</p>`,
    });
  };

  const onClick = async (event) => {
    const button = event.target.closest('button[data-action]');
    if (!button) return;
    const s = schedules.find((x) => x.id === button.closest('tr')?.dataset.id);
    if (!s) return;
    const action = button.dataset.action;
    try {
      if (action === 'history') return await showHistory(s);
      if (action === 'run') {
        const deleting = s.action === 'delete';
        const message = deleting
          ? `Executar agora o agendamento "${s.name}"? Os itens encontrados serão EXCLUÍDOS, como definido no agendamento. A próxima execução programada não muda.`
          : `Executar agora o agendamento "${s.name}"? A próxima execução programada não muda.`;
        if (!(await confirmDialog(message, { title: 'Executar agora', confirmLabel: deleting ? 'Executar e excluir' : 'Executar agora', danger: deleting }))) return;
        const { scan } = await post(`/api/schedules/${s.id}/run`, { confirm: true });
        toast('Análise iniciada.', 'success');
        location.hash = reportLink(s, scan.id);
        return;
      }
      if (action === 'pause') {
        await post(`/api/schedules/${s.id}/pause`);
        toast(`Agendamento "${s.name}" pausado.`);
      }
      if (action === 'resume') {
        const updated = await post(`/api/schedules/${s.id}/resume`);
        toast(updated.nextRunAt ? `Agendamento retomado. Próxima execução: ${fmtServerDateTime(updated.nextRunAt)}.` : 'Agendamento retomado, mas a regra não tem mais execuções futuras.', updated.nextRunAt ? 'success' : 'info');
      }
      if (action === 'delete') {
        if (!(await confirmDialog(`Excluir o agendamento "${s.name}"? Os relatórios já gerados são mantidos.`, { confirmLabel: 'Excluir' }))) return;
        await del(`/api/schedules/${s.id}`);
        toast('Agendamento excluído.', 'success');
      }
      await refresh({ force: true });
    } catch (err) {
      toast(err.message, 'error');
      await refresh({ force: true }).catch(() => {});
    }
  };

  await refresh({ force: true });
  root.addEventListener('click', onClick);
  return () => {
    stopped = true;
    clearTimeout(timer);
    root.removeEventListener('click', onClick);
  };
}
