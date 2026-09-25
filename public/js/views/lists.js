// Listas de referência: conjuntos de termos procurados nos nomes e conteúdos.
import { get, del, post } from '../api.js';
import { html, render as paint, icon, confirmDialog, toast, fmtNum, fmtDateTime } from '../ui.js';

export async function render(root) {
  let lists = [];

  const draw = () =>
    paint(
      root,
      html`<div class="page-head">
          <div>
            <h1>Listas de referência</h1>
            <div class="sub">Termos procurados no nome e no conteúdo dos arquivos: palavras, nomes, códigos ou expressões regulares.</div>
          </div>
          <div class="actions"><a class="btn primary" href="#/listas/nova">${icon('plus')} Nova lista</a></div>
        </div>
        <section class="card">
          ${lists.length === 0
            ? html`<div class="empty">
                <p>Nenhuma lista cadastrada.</p>
                <a class="btn primary" href="#/listas/nova">${icon('plus')} Criar a primeira lista</a>
              </div>`
            : html`<div class="table-wrap">
                <table class="data">
                  <thead><tr><th>Nome</th><th>Descrição</th><th class="num">Termos</th><th>Atualizada em</th><th><span class="sr-only">Ações</span></th></tr></thead>
                  <tbody>
                    ${lists.map(
                      (l) => html`<tr>
                        <td><a href="#/listas/${l.id}"><b>${l.name}</b></a></td>
                        <td class="small">${l.description || html`<span class="muted">—</span>`}</td>
                        <td class="num">${fmtNum(l.termCount)}</td>
                        <td class="nowrap">${fmtDateTime(l.updatedAt)}</td>
                        <td class="actions">
                          <a class="icon-btn" href="#/listas/${l.id}" aria-label="Editar ${l.name}" title="Editar">${icon('edit')}</a>
                          <button class="icon-btn" data-action="copy" data-id="${l.id}" aria-label="Duplicar ${l.name}" title="Duplicar">${icon('copy')}</button>
                          <button class="icon-btn danger" data-action="delete" data-id="${l.id}" aria-label="Excluir ${l.name}" title="Excluir">${icon('trash')}</button>
                        </td>
                      </tr>`,
                    )}
                  </tbody>
                </table>
              </div>`}
        </section>`,
    );

  const refresh = async () => {
    lists = await get('/api/lists');
    draw();
  };

  const onClick = async (event) => {
    const button = event.target.closest('button[data-action]');
    if (!button) return;
    const list = lists.find((l) => l.id === button.dataset.id);
    try {
      if (button.dataset.action === 'copy') {
        const full = await get(`/api/lists/${list.id}`);
        await post('/api/lists', { name: `${full.name} (cópia)`, description: full.description, terms: full.terms.map(({ id, ...t }) => t) });
        toast('Lista duplicada.', 'success');
        await refresh();
      }
      if (button.dataset.action === 'delete') {
        const ok = await confirmDialog(`Excluir a lista "${list.name}" com ${fmtNum(list.termCount)} termo(s)? As análises já feitas continuam disponíveis.`, {
          confirmLabel: 'Excluir',
        });
        if (!ok) return;
        await del(`/api/lists/${list.id}`);
        toast('Lista excluída.', 'success');
        await refresh();
      }
    } catch (err) {
      toast(err.message, 'error');
    }
  };

  await refresh();
  root.addEventListener('click', onClick);
  return () => root.removeEventListener('click', onClick);
}
