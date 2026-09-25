// Formulário de nova análise.
import { get, post } from '../api.js';
import { html, render as paint, icon, toast, fmtNum, plural } from '../ui.js';
import { go } from '../nav.js';

const CLOUD_LABELS = { onedrive: 'OneDrive', sharepoint: 'SharePoint' };

/** Forma de exclusão de cada repositório: pastas do Windows, sempre definitiva; na nuvem, a do cadastro. */
const deletionChip = (r) => (CLOUD_LABELS[r.type] && r.deleteMode !== 'permanent' ? 'exclusão para a lixeira' : 'exclusão definitiva');

export async function render(root, { ctx }) {
  const [repos, lists] = await Promise.all([get('/api/repositories'), get('/api/lists')]);
  const d = ctx.info.defaults || {};
  const usable = lists.filter((l) => l.termCount > 0);

  if (repos.length === 0 || usable.length === 0) {
    paint(
      root,
      html`<div class="page-head"><div><h1>Nova análise</h1></div></div>
        <div class="alert info">${icon('info')}<div>
          Para iniciar uma análise é preciso ter ao menos um repositório e uma lista de referência com termos.
          <div class="inline page-actions">
            ${repos.length === 0 ? html`<a class="btn small" href="#/repositorios">Cadastrar repositório</a>` : ''}
            ${usable.length === 0 ? html`<a class="btn small" href="#/listas/nova">Criar lista de referência</a>` : ''}
          </div>
        </div></div>`,
    );
    return null;
  }

  paint(
    root,
    html`<div class="page-head">
        <div>
          <h1>Nova análise</h1>
          <div class="sub">Escolha onde procurar, o que procurar e como.</div>
        </div>
      </div>
      <form class="card" data-form novalidate>
        <div class="form-grid">
          <label class="field full">
            <span>Nome da análise (opcional)</span>
            <input type="text" name="name" maxlength="200" placeholder="Ex.: Varredura LGPD – setembro" />
          </label>

          <fieldset>
            <legend>Repositórios</legend>
            <div class="choice-list">
              ${repos.map(
                (r) => html`<label class="check">
                  <input type="checkbox" name="repositoryIds" value="${r.id}" ${repos.length === 1 ? 'checked' : ''} />
                  <span><b>${r.name}</b>${CLOUD_LABELS[r.type] ? html` <span class="chip">${CLOUD_LABELS[r.type]}</span>` : ''}${r.allowDelete ? html` <span class="chip danger">${deletionChip(r)}</span>` : ''}<br /><span class="${CLOUD_LABELS[r.type] ? 'muted small' : 'mono muted'}">${r.path}</span></span>
                </label>`,
              )}
            </div>
          </fieldset>

          <fieldset>
            <legend>Listas de referência</legend>
            <div class="choice-list">
              ${usable.map(
                (l) => html`<label class="check">
                  <input type="checkbox" name="listIds" value="${l.id}" ${usable.length === 1 ? 'checked' : ''} />
                  <span><b>${l.name}</b><br /><span class="muted small">${plural(l.termCount, 'termo', 'termos')}</span></span>
                </label>`,
              )}
            </div>
          </fieldset>

          <fieldset class="full">
            <legend>O que verificar</legend>
            <div class="form-grid">
              <div class="field">
                <label class="check"><input type="checkbox" name="checkName" ${d.checkName !== false ? 'checked' : ''} /><span><b>Nome dos arquivos</b></span></label>
                <select name="nameTarget" aria-label="Parte do nome verificada">
                  <option value="file">Somente o nome do arquivo</option>
                  <option value="path">Caminho completo (inclui os nomes das pastas)</option>
                </select>
              </div>
              <div class="field">
                <label class="check"><input type="checkbox" name="checkContent" ${d.checkContent !== false ? 'checked' : ''} /><span><b>Conteúdo dos arquivos</b></span></label>
                <small>Word, Excel, PowerPoint (novos e 97-2003), PDF, OpenDocument, RTF, e-mails .msg, textos, CSV, HTML e nomes dentro de .zip.</small>
              </div>
            </div>
          </fieldset>

          <fieldset class="full">
            <legend>Filtros e desempenho</legend>
            <div class="form-grid">
              <label class="field">
                <span>Somente arquivos modificados a partir de</span>
                <input type="date" name="modifiedAfter" />
                <small>Em branco: todos os arquivos.</small>
              </label>
              <label class="field">
                <span>Tamanho máximo para ler o conteúdo (MB)</span>
                <input type="number" name="maxFileSizeMB" min="1" max="2048" value="${d.maxFileSizeMB || 50}" />
                <small>Arquivos maiores têm só o nome verificado (textos longos: apenas o início).</small>
              </label>
              <label class="field">
                <span>Arquivos processados em paralelo</span>
                <input type="number" name="concurrency" min="1" max="16" value="${d.concurrency || 4}" />
                <small>Aumente para servidores rápidos; diminua para não sobrecarregar a rede.</small>
              </label>
              <label class="check">
                <input type="checkbox" name="resolveOwner" ${d.resolveOwner !== false ? 'checked' : ''} />
                <span><b>Identificar o proprietário do arquivo (NTFS)</b><br /><small class="muted">Pastas do Windows: usado quando não há log de auditoria nem metadados do documento. No OneDrive e no SharePoint, o último usuário vem do Microsoft 365.</small></span>
              </label>
            </div>
          </fieldset>
          <fieldset class="full">
            <legend>O que fazer com os arquivos encontrados</legend>
            <label class="check">
              <input type="radio" name="action" value="analyze" checked />
              <span><b>Somente analisar</b><br /><small class="muted">Gera o relatório. Depois, se quiser, exclua item a item pelo relatório.</small></span>
            </label>
            <label class="check">
              <input type="radio" name="action" value="delete" />
              <span><b>Analisar e excluir automaticamente</b><br /><small class="muted">Todo arquivo em que algum termo for encontrado é excluído, sem confirmação item a item (arquivos alterados depois de lidos são mantidos). Só para repositórios com "Permitir exclusão".</small></span>
            </label>
            <div class="alert error" data-delete-confirm hidden>
              ${icon('alert')}
              <div>
                <b>Exclusão sem volta nas pastas do Windows:</b> arquivos excluídos pela rede não vão para a Lixeira. No OneDrive e no SharePoint, vale a forma definida em cada repositório (para a lixeira ou definitiva). Confira as listas de referência antes de continuar: tudo o que for encontrado será excluído.
                <p data-path-warning hidden><b>Atenção:</b> com "Caminho completo", um termo no nome de uma pasta faz todos os arquivos dela (e das subpastas) serem encontrados — e excluídos.</p>
                <label class="field"><span>Digite EXCLUIR para confirmar</span><input type="text" name="confirmDelete" autocomplete="off" spellcheck="false" /></label>
              </div>
            </div>
          </fieldset>
        </div>
        <div class="inline page-actions">
          <button type="submit" class="btn primary" data-submit>${icon('play')} Iniciar análise</button>
          <a class="btn" href="#/analises">Cancelar</a>
        </div>
      </form>`,
  );

  const form = root.querySelector('[data-form]');
  const submit = form.querySelector('[data-submit]');
  const onChange = (event) => {
    if (!['action', 'nameTarget', 'checkName'].includes(event.target.name)) return;
    const deleting = form.elements.action.value === 'delete';
    form.querySelector('[data-delete-confirm]').hidden = !deleting;
    form.querySelector('[data-path-warning]').hidden = !(form.elements.checkName.checked && form.elements.nameTarget.value === 'path');
    submit.className = `btn ${deleting ? 'danger' : 'primary'}`;
    paint(submit, html`${icon('play')} ${deleting ? 'Iniciar análise e exclusão' : 'Iniciar análise'}`);
  };
  form.addEventListener('change', onChange);
  const onSubmit = async (event) => {
    event.preventDefault();
    const f = new FormData(form);
    const deleting = f.get('action') === 'delete';
    const body = {
      name: f.get('name'),
      repositoryIds: f.getAll('repositoryIds'),
      listIds: f.getAll('listIds'),
      options: {
        checkName: f.get('checkName') === 'on',
        nameTarget: f.get('nameTarget'),
        checkContent: f.get('checkContent') === 'on',
        modifiedAfter: f.get('modifiedAfter') ? `${f.get('modifiedAfter')}T00:00:00` : null,
        maxFileSizeMB: Number(f.get('maxFileSizeMB')),
        concurrency: Number(f.get('concurrency')),
        resolveOwner: f.get('resolveOwner') === 'on',
        deleteMatches: deleting,
      },
      confirmDelete: deleting ? String(f.get('confirmDelete') || '') : '',
    };
    if (body.repositoryIds.length === 0) return toast('Selecione ao menos um repositório.', 'error');
    if (body.listIds.length === 0) return toast('Selecione ao menos uma lista de referência.', 'error');
    if (deleting) {
      const blocked = repos.filter((r) => body.repositoryIds.includes(r.id) && !r.allowDelete).map((r) => r.name);
      if (blocked.length) return toast(`A exclusão não está permitida em: ${blocked.join(', ')}. Ative em Repositórios ou escolha "Somente analisar".`, 'error');
      if (body.confirmDelete.trim().toUpperCase() !== 'EXCLUIR') return toast('Digite EXCLUIR para confirmar a exclusão.', 'error');
    }
    const button = submit;
    button.disabled = true;
    try {
      const scan = await post('/api/scans', body);
      toast(`Análise iniciada (${fmtNum(scan.summary.termCount)} termos).`, 'success');
      go(`/analises/${scan.id}`);
    } catch (err) {
      toast(err.message, 'error');
      button.disabled = false;
    }
  };
  form.addEventListener('submit', onSubmit);
  return () => {
    form.removeEventListener('submit', onSubmit);
    form.removeEventListener('change', onChange);
  };
}
