// Lista de análises, com progresso das que estão em andamento.
import { get, post, del } from '../api.js';
import { html, render as paint, icon, toast, confirmDialog, fmtNum, fmtDateTime, fmtDuration, statusBadge } from '../ui.js';

const active = (s) => s.status === 'running' || s.status === 'queued';

function duration(scan) {
  if (!scan.startedAt) return '—';
  const end = scan.finishedAt ? new Date(scan.finishedAt) : new Date();
  return fmtDuration(end - new Date(scan.startedAt));
}

export async function render(root) {
  let scans = [];
  let timer = null;
  let stopped = false;

  const draw = () =>
    paint(
      root,
      html`<div class="page-head">
          <div>
            <h1>Análises</h1>
            <div class="sub">Cada análise percorre os repositórios escolhidos e gera um relatório com os arquivos em que algum termo foi encontrado.</div>
          </div>
          <div class="actions"><a class="btn primary" href="#/analises/nova">${icon('play')} Nova análise</a></div>
        </div>
        <section class="card">
          ${scans.length === 0
            ? html`<div class="empty"><p>Nenhuma análise realizada ainda.</p><a class="btn primary" href="#/analises/nova">${icon('play')} Iniciar a primeira</a></div>`
            : html`<div class="table-wrap">
                <table class="data">
                  <thead>
                    <tr>
                      <th>Análise</th><th>Situação</th><th>Início</th><th>Duração</th>
                      <th class="num">Arquivos verificados</th><th class="num">Com ocorrências</th><th class="num">Erros</th>
                      <th><span class="sr-only">Ações</span></th>
                    </tr>
                  </thead>
                  <tbody>
                    ${scans.map(
                      (s) => html`<tr>
                        <td>
                          <a href="#/analises/${s.id}"><b>${s.name}</b></a>
                          <div class="muted small">${(s.summary?.repositories || []).map((r) => r.name).join(', ')} · ${(s.summary?.lists || []).map((l) => l.name).join(', ')}</div>
                          ${active(s) ? html`<div class="progress-line" role="progressbar" aria-label="Análise em andamento"></div>` : ''}
                        </td>
                        <td>${statusBadge(s.status)}</td>
                        <td class="nowrap">${fmtDateTime(s.startedAt || s.createdAt)}</td>
                        <td class="nowrap">${duration(s)}</td>
                        <td class="num">${fmtNum(s.stats?.filesSeen)}</td>
                        <td class="num">${fmtNum(s.stats?.filesMatched)}</td>
                        <td class="num">${fmtNum(s.stats?.errors)}</td>
                        <td class="actions">
                          <a class="icon-btn" href="#/analises/${s.id}" aria-label="Abrir relatório de ${s.name}" title="Abrir relatório">${icon('file')}</a>
                          ${active(s)
                            ? html`<button class="icon-btn danger" data-action="cancel" data-id="${s.id}" aria-label="Cancelar ${s.name}" title="Cancelar">${icon('stop')}</button>`
                            : html`<button class="icon-btn danger" data-action="delete" data-id="${s.id}" aria-label="Excluir ${s.name}" title="Excluir">${icon('trash')}</button>`}
                        </td>
                      </tr>`,
                    )}
                  </tbody>
                </table>
              </div>`}
        </section>`,
    );

  const refresh = async () => {
    const latest = await get('/api/scans');
    if (stopped) return;
    scans = latest;
    draw();
    clearTimeout(timer);
    if (scans.some(active)) timer = setTimeout(() => refresh().catch(() => {}), 2000);
  };

  const onClick = async (event) => {
    const button = event.target.closest('button[data-action]');
    if (!button) return;
    const scan = scans.find((s) => s.id === button.dataset.id);
    try {
      if (button.dataset.action === 'cancel') {
        if (!(await confirmDialog(`Cancelar a análise "${scan.name}"? Os resultados encontrados até agora serão mantidos.`, { confirmLabel: 'Cancelar análise' }))) return;
        await post(`/api/scans/${scan.id}/cancel`);
        toast('Cancelamento solicitado.');
      }
      if (button.dataset.action === 'delete') {
        if (!(await confirmDialog(`Excluir a análise "${scan.name}" e o seu relatório?`, { confirmLabel: 'Excluir' }))) return;
        await del(`/api/scans/${scan.id}`);
        toast('Análise excluída.', 'success');
      }
      await refresh();
    } catch (err) {
      toast(err.message, 'error');
    }
  };

  await refresh();
  root.addEventListener('click', onClick);
  return () => {
    stopped = true;
    clearTimeout(timer);
    root.removeEventListener('click', onClick);
  };
}
