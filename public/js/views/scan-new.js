// Formulário de nova análise de arquivos: agora ou agendada (também edita agendamentos).
import { get, post, put } from '../api.js';
import { html, render as paint, icon, toast, fmtNum, plural, fmtServerDateTime } from '../ui.js';
import { go } from '../nav.js';
import { scheduleSection, bindSchedule, readSchedule, scheduling } from '../schedule-form.js';

const CLOUD_LABELS = { onedrive: 'OneDrive', sharepoint: 'SharePoint' };

/** Forma de exclusão de cada repositório: pastas do Windows, sempre definitiva; na nuvem, a do cadastro. */
const deletionChip = (r) => (CLOUD_LABELS[r.type] && r.deleteMode !== 'permanent' ? 'exclusão para a lixeira' : 'exclusão definitiva');

/**
 * props.scheduling: tela só de agendamento (vinda de Agendamentos); props.schedule: agendamento em
 * edição.
 */
export async function render(root, { ctx, props = {} }) {
  const [repos, lists] = await Promise.all([get('/api/repositories'), get('/api/lists')]);
  const schedule = props.schedule || null;
  const fixed = Boolean(props.scheduling || schedule);
  const d = schedule?.options || ctx.info.defaults || {};
  const usable = lists.filter((l) => l.termCount > 0);
  const title = schedule ? 'Editar agendamento' : fixed ? 'Novo agendamento de arquivos' : 'Nova análise';
  const back = fixed ? '#/agendamentos' : '#/analises';

  if (repos.length === 0 || usable.length === 0) {
    paint(
      root,
      html`<div class="page-head"><div><h1>${title}</h1></div></div>
        <div class="alert info">${icon('info')}<div>
          Para ${fixed ? 'agendar' : 'iniciar'} uma análise é preciso ter ao menos um repositório e uma lista de referência com termos.
          <div class="inline page-actions">
            ${repos.length === 0 ? html`<a class="btn small" href="#/repositorios">Cadastrar repositório</a>` : ''}
            ${usable.length === 0 ? html`<a class="btn small" href="#/listas/nova">Criar lista de referência</a>` : ''}
          </div>
        </div></div>`,
    );
    return null;
  }

  const chosen = (list, id) => (schedule ? (list || []).includes(id) : null);
  const section = scheduleSection({ kind: 'files', schedule, fixed, info: ctx.info });
  paint(
    root,
    html`<div class="page-head">
        <div>
          <h1>${title}</h1>
          <div class="sub">${fixed ? 'Escolha onde procurar, o que procurar, como e quando.' : 'Escolha onde procurar, o que procurar e como.'}</div>
        </div>
      </div>
      <form class="card" data-form novalidate>
        <div class="form-grid">
          <label class="field full">
            <span data-name-label>Nome da análise (opcional)</span>
            <input type="text" name="name" maxlength="200" value="${schedule?.name || ''}" placeholder="Ex.: Varredura LGPD – setembro" />
          </label>
          ${fixed ? section : ''}

          <fieldset>
            <legend>Repositórios</legend>
            <div class="choice-list">
              ${repos.map(
                (r) => html`<label class="check">
                  <input type="checkbox" name="repositoryIds" value="${r.id}" ${(chosen(schedule?.targetIds, r.id) ?? repos.length === 1) ? 'checked' : ''} />
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
                  <input type="checkbox" name="listIds" value="${l.id}" ${(chosen(schedule?.listIds, l.id) ?? usable.length === 1) ? 'checked' : ''} />
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
                  <option value="path" ${d.nameTarget === 'path' ? 'selected' : ''}>Caminho completo (inclui os nomes das pastas)</option>
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
              <label class="field" data-now-only>
                <span>Somente arquivos alterados a partir de</span>
                <input type="date" name="modifiedAfter" />
                <small>Em branco: todos os arquivos. Vale a data mais recente entre a modificação e a chegada ao repositório (arquivo copiado ou movido para ele).</small>
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
              <input type="radio" name="action" value="analyze" ${schedule?.action === 'delete' ? '' : 'checked'} />
              <span><b>Somente analisar</b><br /><small class="muted">Gera o relatório. Depois, se quiser, exclua item a item pelo relatório.</small></span>
            </label>
            <label class="check">
              <input type="radio" name="action" value="delete" ${schedule?.action === 'delete' ? 'checked' : ''} />
              <span><b>Analisar e excluir automaticamente</b><br /><small class="muted">Todo arquivo em que algum termo for encontrado é excluído, sem confirmação item a item (arquivos alterados depois de lidos são mantidos). Só para repositórios com "Permitir exclusão".</small></span>
            </label>
            <div class="alert error" data-delete-confirm hidden>
              ${icon('alert')}
              <div>
                <b>Exclusão sem volta nas pastas do Windows:</b> arquivos excluídos pela rede não vão para a Lixeira. No OneDrive e no SharePoint, vale a forma definida em cada repositório (para a lixeira ou definitiva). Confira as listas de referência antes de continuar: tudo o que for encontrado será excluído.
                <p data-path-warning hidden><b>Atenção:</b> com "Caminho completo", um termo no nome de uma pasta faz todos os arquivos dela (e das subpastas) serem encontrados — e excluídos.</p>
                <p data-schedule-only hidden><b>Agendamento:</b> a exclusão vale para todas as execuções, sem nova confirmação. Ela é conferida em cada execução: se um repositório deixar de permitir a exclusão, mudar de caminho (ou de contas e sites) ou passar a excluir de forma definitiva, as execuções falham até o agendamento ser salvo e confirmado de novo.</p>
                <label class="field"><span>Digite EXCLUIR para confirmar${schedule?.action === 'delete' ? ' (de novo, a cada vez que o agendamento é salvo)' : ''}</span><input type="text" name="confirmDelete" autocomplete="off" spellcheck="false" /></label>
              </div>
            </div>
          </fieldset>
          ${fixed ? '' : section}
        </div>
        <div class="inline page-actions">
          <button type="submit" class="btn primary" data-submit>${icon('play')} Iniciar análise</button>
          <a class="btn" href="${back}">Cancelar</a>
        </div>
      </form>`,
  );

  const form = root.querySelector('[data-form]');
  const submit = form.querySelector('[data-submit]');
  const sync = () => {
    const deleting = form.elements.action.value === 'delete';
    const later = scheduling(form);
    form.querySelector('[data-delete-confirm]').hidden = !deleting;
    form.querySelector('[data-path-warning]').hidden = !(form.elements.checkName.checked && form.elements.nameTarget.value === 'path');
    form.querySelector('[data-schedule-only]').hidden = !later;
    form.querySelectorAll('[data-now-only]').forEach((el) => {
      el.hidden = later;
    });
    form.querySelector('[data-name-label]').textContent = later ? 'Nome do agendamento' : 'Nome da análise (opcional)';
    form.elements.name.placeholder = later ? 'Ex.: Varredura LGPD semanal' : 'Ex.: Varredura LGPD – setembro';
    submit.className = `btn ${deleting ? 'danger' : 'primary'}`;
    const label = later ? (deleting ? 'Salvar agendamento com exclusão' : 'Salvar agendamento') : deleting ? 'Iniciar análise e exclusão' : 'Iniciar análise';
    paint(submit, html`${icon(later ? 'clock' : 'play')} ${label}`);
  };
  const onChange = (event) => {
    if (['action', 'nameTarget', 'checkName'].includes(event.target.name)) sync();
  };
  form.addEventListener('change', onChange);
  const unbind = bindSchedule(form, { onModeChange: sync });
  const onSubmit = async (event) => {
    event.preventDefault();
    const f = new FormData(form);
    const deleting = f.get('action') === 'delete';
    const later = scheduling(form);
    const body = {
      kind: 'files',
      name: String(f.get('name') || '').trim(),
      repositoryIds: f.getAll('repositoryIds'),
      listIds: f.getAll('listIds'),
      options: {
        checkName: f.get('checkName') === 'on',
        nameTarget: f.get('nameTarget'),
        checkContent: f.get('checkContent') === 'on',
        modifiedAfter: !later && f.get('modifiedAfter') ? `${f.get('modifiedAfter')}T00:00:00` : null,
        maxFileSizeMB: Number(f.get('maxFileSizeMB')),
        concurrency: Number(f.get('concurrency')),
        resolveOwner: f.get('resolveOwner') === 'on',
        deleteMatches: deleting,
      },
      confirmDelete: deleting ? String(f.get('confirmDelete') || '') : '',
    };
    if (later && !body.name) {
      form.elements.name.focus();
      return toast('Dê um nome ao agendamento.', 'error');
    }
    if (body.repositoryIds.length === 0) return toast('Selecione ao menos um repositório.', 'error');
    if (body.listIds.length === 0) return toast('Selecione ao menos uma lista de referência.', 'error');
    if (deleting) {
      const blocked = repos.filter((r) => body.repositoryIds.includes(r.id) && !r.allowDelete).map((r) => r.name);
      if (blocked.length) return toast(`A exclusão não está permitida em: ${blocked.join(', ')}. Ative em Repositórios ou escolha "Somente analisar".`, 'error');
      if (body.confirmDelete.trim().toUpperCase() !== 'EXCLUIR') return toast('Digite EXCLUIR para confirmar a exclusão.', 'error');
    }
    submit.disabled = true;
    try {
      if (later) {
        const saved = await (schedule ? put(`/api/schedules/${schedule.id}`, { ...body, ...readSchedule(form) }) : post('/api/schedules', { ...body, ...readSchedule(form) }));
        const next = saved.nextRunAt ? ` Próxima execução: ${fmtServerDateTime(saved.nextRunAt)}.` : saved.state === 'paused' ? ' Ele está pausado.' : '';
        toast(`Agendamento salvo.${next}`, 'success');
        go('/agendamentos');
        return;
      }
      const scan = await post('/api/scans', body);
      toast(`Análise iniciada (${fmtNum(scan.summary.termCount)} termos).`, 'success');
      go(`/analises/${scan.id}`);
    } catch (err) {
      toast(err.message, 'error');
      submit.disabled = false;
    }
  };
  form.addEventListener('submit', onSubmit);
  return () => {
    unbind();
    form.removeEventListener('submit', onSubmit);
    form.removeEventListener('change', onChange);
  };
}
