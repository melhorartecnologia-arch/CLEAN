// Lista de análises (de arquivos ou de e-mail), com progresso das que estão em andamento.
import { get, post, del } from '../api.js';
import { html, render as paint, icon, toast, confirmDialog, fmtNum, fmtDateTime, fmtDuration, statusBadge } from '../ui.js';
import { describeRetention } from '../retention.js';

const active = (s) => s.status === 'running' || s.status === 'queued';

function duration(scan) {
  if (!scan.startedAt) return '—';
  const end = scan.finishedAt ? new Date(scan.finishedAt) : new Date();
  return fmtDuration(end - new Date(scan.startedAt));
}

const KINDS = {
  files: {
    title: 'Análises de arquivos',
    sub: 'Cada análise percorre os repositórios escolhidos e gera um relatório com os arquivos em que algum termo foi encontrado.',
    base: '#/analises',
    empty: 'Nenhuma análise de arquivos realizada ainda.',
    where: (s) => (s.summary?.repositories || []).map((r) => r.name).join(', '),
    columns: ['Arquivos verificados', 'Com ocorrências'],
    values: (s) => [s.stats?.filesSeen, s.stats?.filesMatched],
  },
  mail: {
    title: 'Análises de e-mail',
    sub: 'Cada análise percorre as caixas de e-mail escolhidas (corpo, assunto e anexos das mensagens) e gera um relatório com as mensagens em que algum termo foi encontrado.',
    base: '#/email/analises',
    empty: 'Nenhuma análise de e-mail realizada ainda.',
    where: (s) => (s.summary?.sources || []).map((r) => r.name).join(', '),
    columns: ['Mensagens verificadas', 'Com ocorrências'],
    values: (s) => [s.stats?.messagesSeen, s.stats?.messagesMatched],
  },
};

export async function render(root, { props = {} }) {
  const kind = props.kind === 'mail' ? 'mail' : 'files';
  const K = KINDS[kind];
  let scans = [];
  let timer = null;
  let stopped = false;

  const draw = () =>
    paint(
      root,
      html`<div class="page-head">
          <div>
            <h1>${K.title}</h1>
            <div class="sub">${K.sub}</div>
          </div>
          <div class="actions"><a class="btn primary" href="${K.base}/nova">${icon('play')} Nova análise</a></div>
        </div>
        <section class="card">
          ${scans.length === 0
            ? html`<div class="empty"><p>${K.empty}</p><a class="btn primary" href="${K.base}/nova">${icon('play')} Iniciar a primeira</a></div>`
            : html`<div class="table-wrap">
                <table class="data">
                  <thead>
                    <tr>
                      <th>Análise</th><th>Situação</th><th>Início</th><th>Duração</th>
                      <th class="num">${K.columns[0]}</th><th class="num">${K.columns[1]}</th><th class="num">Erros</th>
                      <th><span class="sr-only">Ações</span></th>
                    </tr>
                  </thead>
                  <tbody>
                    ${scans.map((s) => {
                      const [seen, matched] = K.values(s);
                      return html`<tr>
                        <td>
                          <a href="${K.base}/${s.id}"><b>${s.name}</b></a>${s.retention
                            ? html` <span class="chip" title="${s.scheduleId ? `Execução da política de retenção "${s.scheduleName}"` : 'Execução de uma política de retenção'}">${icon('clock')} retenção</span>`
                            : s.scheduleId
                              ? html` <span class="chip" title="Iniciada pelo agendamento &quot;${s.scheduleName}&quot;">${icon('clock')} agendada</span>`
                              : ''}
                          <div class="muted small">${K.where(s)} · ${s.retention ? describeRetention(s.retention, kind) : (s.summary?.lists || []).map((l) => l.name).join(', ')}</div>
                          ${active(s) ? html`<div class="progress-line" role="progressbar" aria-label="Análise em andamento"></div>` : ''}
                        </td>
                        <td>${statusBadge(s.status)}</td>
                        <td class="nowrap">${fmtDateTime(s.startedAt || s.createdAt)}</td>
                        <td class="nowrap">${duration(s)}</td>
                        <td class="num">${fmtNum(seen)}</td>
                        <td class="num">${fmtNum(matched)}${s.retention ? html`<div class="muted small">expirad${kind === 'mail' ? 'as' : 'os'}</div>` : ''}</td>
                        <td class="num">${fmtNum(s.stats?.errors)}</td>
                        <td class="actions">
                          <a class="icon-btn" href="${K.base}/${s.id}" aria-label="Abrir relatório de ${s.name}" title="Abrir relatório">${icon('file')}</a>
                          ${active(s)
                            ? html`<button class="icon-btn danger" data-action="cancel" data-id="${s.id}" aria-label="Cancelar ${s.name}" title="Cancelar">${icon('stop')}</button>`
                            : html`<button class="icon-btn danger" data-action="delete" data-id="${s.id}" aria-label="Excluir ${s.name}" title="Excluir">${icon('trash')}</button>`}
                        </td>
                      </tr>`;
                    })}
                  </tbody>
                </table>
              </div>`}
        </section>`,
    );

  const refresh = async () => {
    const latest = await get(`/api/scans?kind=${kind}`);
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
