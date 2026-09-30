// Listagens de e-mail (somente leitura): catálogo de contas do domínio e listagem de mensagens por
// caixa. Lista as listagens já feitas, com o progresso das que estão em andamento.
import { get, post, del } from '../api.js';
import { html, render as paint, icon, toast, confirmDialog, fmtNum, fmtDateTime, fmtDuration, statusBadge } from '../ui.js';

const active = (s) => s.status === 'running' || s.status === 'queued';
const TYPE_LABELS = { graph: 'Microsoft 365', gmail: 'Google Workspace', imap: 'IMAP' };
const KIND = {
  directory: { chip: 'contas do domínio', column: 'Contas', value: (s) => s.stats?.accounts },
  messages: { chip: 'mensagens', column: 'Mensagens', value: (s) => s.stats?.messagesSeen },
};

function duration(scan) {
  if (!scan.startedAt) return '—';
  const end = scan.finishedAt ? new Date(scan.finishedAt) : new Date();
  return fmtDuration(end - new Date(scan.startedAt));
}

export async function render(root) {
  const base = '#/email/listagens';
  let listings = [];
  let timer = null;
  let stopped = false;

  const draw = () =>
    paint(
      root,
      html`<div class="page-head">
          <div>
            <h1>Listagens de e-mail</h1>
            <div class="sub">Inventários somente leitura: o catálogo de contas registradas no domínio e a lista de mensagens de cada caixa, com todos os dados de cada uma (sem procurar termos e sem excluir nada).</div>
          </div>
          <div class="actions">
            <a class="btn" href="${base}/nova?tipo=directory">${icon('mail')} Contas do domínio</a>
            <a class="btn primary" href="${base}/nova?tipo=messages">${icon('mail')} Mensagens por caixa</a>
          </div>
        </div>
        <section class="card">
          ${listings.length === 0
            ? html`<div class="empty"><p>Nenhuma listagem feita ainda.</p><div class="inline"><a class="btn" href="${base}/nova?tipo=directory">Listar contas do domínio</a><a class="btn primary" href="${base}/nova?tipo=messages">Listar mensagens por caixa</a></div></div>`
            : html`<div class="table-wrap">
                <table class="data">
                  <thead>
                    <tr>
                      <th>Listagem</th><th>Situação</th><th>Início</th><th>Duração</th>
                      <th class="num">Itens</th><th class="num">Erros</th>
                      <th><span class="sr-only">Ações</span></th>
                    </tr>
                  </thead>
                  <tbody>
                    ${listings.map((s) => {
                      const K = KIND[s.listing?.kind] || KIND.messages;
                      const where = (s.summary?.sources || []).map((x) => `${x.name} (${TYPE_LABELS[x.type] || x.type})`).join(', ');
                      return html`<tr>
                        <td>
                          <a href="${base}/${s.id}"><b>${s.name}</b></a> <span class="chip">${K.chip}</span>
                          <div class="muted small">${where}</div>
                          ${active(s) ? html`<div class="progress-line" role="progressbar" aria-label="Listagem em andamento"></div>` : ''}
                        </td>
                        <td>${statusBadge(s.status)}</td>
                        <td class="nowrap">${fmtDateTime(s.startedAt || s.createdAt)}</td>
                        <td class="nowrap">${duration(s)}</td>
                        <td class="num">${fmtNum(K.value(s))}<div class="muted small">${K.column.toLowerCase()}</div></td>
                        <td class="num">${fmtNum(s.stats?.errors)}</td>
                        <td class="actions">
                          <a class="icon-btn" href="${base}/${s.id}" aria-label="Abrir ${s.name}" title="Abrir">${icon('file')}</a>
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
    const latest = await get('/api/scans?kind=mail&listing=only');
    if (stopped) return;
    listings = latest;
    draw();
    clearTimeout(timer);
    if (listings.some(active)) timer = setTimeout(() => refresh().catch(() => {}), 2000);
  };

  const onClick = async (event) => {
    const button = event.target.closest('button[data-action]');
    if (!button) return;
    const scan = listings.find((s) => s.id === button.dataset.id);
    try {
      if (button.dataset.action === 'cancel') {
        if (!(await confirmDialog(`Cancelar a listagem "${scan.name}"? O que já foi listado será mantido.`, { confirmLabel: 'Cancelar listagem' }))) return;
        await post(`/api/scans/${scan.id}/cancel`);
        toast('Cancelamento solicitado.');
      }
      if (button.dataset.action === 'delete') {
        if (!(await confirmDialog(`Excluir a listagem "${scan.name}" e o seu relatório?`, { confirmLabel: 'Excluir' }))) return;
        await del(`/api/scans/${scan.id}`);
        toast('Listagem excluída.', 'success');
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
