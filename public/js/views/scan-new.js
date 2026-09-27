// Formulário de nova análise de arquivos: agora ou agendada (também edita agendamentos). Procura os
// termos das listas de referência ou, na busca por tipo, os arquivos das categorias e extensões
// escolhidas (ex.: vídeos, músicas, executáveis).
import { get, post, put } from '../api.js';
import { html, render as paint, icon, toast, fmtNum, plural, fmtServerDateTime } from '../ui.js';
import { go } from '../nav.js';
import { scheduleSection, bindSchedule, readSchedule, scheduling } from '../schedule-form.js';

const CLOUD_LABELS = { onedrive: 'OneDrive', sharepoint: 'SharePoint' };
const SHOWN_EXTENSIONS = 6; // extensões de exemplo em cada categoria (todas em "Ver as extensões")
const MAX_LIMIT = 10_000_000;

/** Forma de exclusão de cada repositório: pastas do Windows, sempre definitiva; na nuvem, a do cadastro. */
const deletionChip = (r) => (CLOUD_LABELS[r.type] && r.deleteMode !== 'permanent' ? 'exclusão para a lixeira' : 'exclusão definitiva');

/** ".mp4, .mov, .avi, .mkv, .wmv, .flv e mais 17". */
function sampleExtensions(category) {
  const shown = category.extensions.slice(0, SHOWN_EXTENSIONS).join(', ');
  const rest = category.extensions.length - SHOWN_EXTENSIONS;
  return rest > 0 ? `${shown} e mais ${rest}` : shown;
}

/** Textos que mudam com o que se procura (termos ou tipos de arquivo). */
const MODE_TEXT = {
  terms: {
    analyze: ['Somente analisar', 'Gera o relatório. Depois, se quiser, exclua pelo relatório: item a item, os selecionados ou todos os filtrados.'],
    delete: [
      'Analisar e excluir automaticamente',
      'Todo arquivo em que algum termo for encontrado é excluído, sem confirmação item a item (arquivos alterados depois de lidos são mantidos). Só para repositórios com "Permitir exclusão".',
    ],
    check: 'Confira as listas de referência antes de continuar: tudo o que for encontrado será excluído.',
    namePlaceholder: ['Ex.: Varredura LGPD – setembro', 'Ex.: Varredura LGPD semanal'],
    start: ['Iniciar análise', 'Iniciar análise e exclusão'],
  },
  types: {
    analyze: ['Somente procurar', 'Gera o relatório para revisão. Depois, exclua pelo relatório: item a item, os selecionados ou todos os filtrados.'],
    delete: [
      'Procurar e excluir automaticamente',
      'Todo arquivo dos tipos escolhidos é excluído durante a busca, sem confirmação item a item, até o limite de cada execução. Só para repositórios com "Permitir exclusão".',
    ],
    check: 'Confira os tipos escolhidos antes de continuar: todo arquivo encontrado será excluído.',
    namePlaceholder: ['Ex.: Vídeos e músicas nos compartilhamentos', 'Ex.: Limpeza semanal de vídeos e músicas'],
    start: ['Iniciar busca', 'Iniciar busca e exclusão'],
  },
};

/**
 * props.scheduling: tela só de agendamento (vinda de Agendamentos); props.schedule: agendamento em
 * edição. ?busca=tipos: abre com a busca por tipo de arquivo.
 */
export async function render(root, { ctx, props = {}, query = new URLSearchParams() }) {
  const [repos, lists] = await Promise.all([get('/api/repositories'), get('/api/lists')]);
  const schedule = props.schedule || null;
  const fixed = Boolean(props.scheduling || schedule);
  const d = schedule?.options || ctx.info.defaults || {};
  const usable = lists.filter((l) => l.termCount > 0);
  const catalog = ctx.info.fileTypes || { categories: [], defaultMaxDeletions: 1000 };
  const t = schedule?.fileTypes || {};
  const title = schedule ? 'Editar agendamento' : fixed ? 'Novo agendamento de arquivos' : 'Nova análise';
  const back = fixed ? '#/agendamentos' : '#/analises';
  // O que procurar: o do agendamento em edição; numa nova análise, os tipos de arquivo se pedido
  // (?busca=tipos) ou enquanto não houver lista de referência com termos.
  const search = schedule ? (schedule.purpose === 'types' ? 'types' : 'terms') : query.get('busca') === 'tipos' || usable.length === 0 ? 'types' : 'terms';

  if (repos.length === 0) {
    paint(
      root,
      html`<div class="page-head"><div><h1>${title}</h1></div></div>
        <div class="alert info">${icon('info')}<div>
          Para ${fixed ? 'agendar' : 'iniciar'} uma análise é preciso ter ao menos um repositório${usable.length === 0 ? ' (e, para procurar termos, uma lista de referência com termos)' : ''}.
          <div class="inline page-actions">
            <a class="btn small" href="#/repositorios">Cadastrar repositório</a>
            ${usable.length === 0 ? html`<a class="btn small" href="#/listas/nova">Criar lista de referência</a>` : ''}
          </div>
        </div></div>`,
    );
    return null;
  }

  const chosen = (list, id) => (schedule ? (list || []).includes(id) : null);
  const chk = (value) => (value ? 'checked' : '');
  const categoryBox = (c) => html`<label class="check">
    <input type="checkbox" name="categories" value="${c.key}" data-work="${c.work ? '1' : ''}" ${chk((t.categories || []).includes(c.key))} />
    <span><b>${c.label}</b><br /><small class="muted">${sampleExtensions(c)}</small></span>
  </label>`;
  const media = catalog.categories.filter((c) => !c.work);
  const work = catalog.categories.filter((c) => c.work);
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
            <input type="text" name="name" maxlength="200" value="${schedule?.name || ''}" />
          </label>
          ${fixed ? section : ''}

          <fieldset class="full">
            <legend>O que procurar</legend>
            <div class="type-choice">
              <label>
                <input type="radio" name="search" value="terms" ${chk(search === 'terms')} />
                <span><b>Termos das listas de referência</b><small>Nomes, palavras e modelos (CPF, CNPJ...) no nome e no conteúdo dos arquivos.</small></span>
              </label>
              <label>
                <input type="radio" name="search" value="types" ${chk(search === 'types')} />
                <span><b>Tipos de arquivo</b><small>Vídeos, músicas, executáveis, compactados ou as extensões que você informar — por exemplo, para liberar espaço.</small></span>
              </label>
            </div>
          </fieldset>

          <fieldset data-repositories>
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

          <fieldset data-for-search="terms">
            <legend>Listas de referência</legend>
            ${usable.length
              ? html`<div class="choice-list">
                  ${usable.map(
                    (l) => html`<label class="check">
                      <input type="checkbox" name="listIds" value="${l.id}" ${(chosen(schedule?.listIds, l.id) ?? usable.length === 1) ? 'checked' : ''} />
                      <span><b>${l.name}</b><br /><span class="muted small">${plural(l.termCount, 'termo', 'termos')}</span></span>
                    </label>`,
                  )}
                </div>`
              : html`<p class="muted">Nenhuma lista de referência com termos ainda.</p><a class="btn small" href="#/listas/nova">Criar lista de referência</a>`}
          </fieldset>

          <fieldset class="full" data-for-search="types">
            <legend>Tipos de arquivo</legend>
            <div class="field-label">Mídia, programas e temporários</div>
            <div class="category-grid">${media.map(categoryBox)}</div>
            <div class="field-label spaced">Arquivos de trabalho <small class="muted">— cuidado ao excluir automaticamente: são os documentos das pessoas</small></div>
            <div class="category-grid">${work.map(categoryBox)}</div>
            <details class="extensions-list">
              <summary>Ver as extensões de cada tipo</summary>
              <dl class="kv">${catalog.categories.map((c) => html`<dt>${c.label}</dt><dd class="mono small">${c.extensions.join(' ')}</dd>`)}</dl>
            </details>
            <div class="form-grid spaced">
              <label class="field">
                <span>Outras extensões</span>
                <input type="text" name="extensions" value="${(t.extensions || []).join(', ')}" placeholder="Ex.: .dwg, .log, .tar.gz" autocomplete="off" spellcheck="false" />
                <small>Separe por vírgula ou espaço. Vale o fim do nome (.tar.gz, .bak...). Podem ser usadas sozinhas, sem marcar tipos.</small>
              </label>
              <label class="field">
                <span>Tamanho mínimo (MB)</span>
                <input type="number" name="minSizeMB" min="0" step="any" value="${t.minSizeMB || 0}" />
                <small>Só os arquivos a partir deste tamanho (ex.: 100, para os vídeos grandes). 0: todos os tamanhos.</small>
              </label>
              <label class="check full">
                <input type="checkbox" name="realType" ${chk(t.checkContent)} />
                <span><b>Conferir o tipo real pelo conteúdo</b><br /><small class="muted">Acha arquivos renomeados (ex.: um vídeo salvo como .pdf) lendo o início dos arquivos sem extensão ou com a extensão de outro tipo da lista. Reconhece vídeos, áudios, imagens, executáveis, compactados, imagens de disco, PDF, arquivos de dados do Outlook e bancos de dados SQLite e Access; planilhas, apresentações, textos e temporários são achados só pela extensão. Extensões fora da lista (.dll, .ai...) valem pelo que são. Vale nas pastas do Windows (no OneDrive e no SharePoint, só a extensão) e deixa a busca mais lenta.</small></span>
              </label>
            </div>
          </fieldset>

          <fieldset class="full" data-for-search="terms">
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
                <small>Em branco: todos os arquivos. Vale a data mais recente entre a modificação e a criação (um arquivo copiado para o repositório entra pela data da cópia).</small>
              </label>
              <label class="field" data-for-search="terms">
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
              <span><b data-text="analyze-label"></b><br /><small class="muted" data-text="analyze-hint"></small></span>
            </label>
            <label class="check">
              <input type="radio" name="action" value="delete" ${schedule?.action === 'delete' ? 'checked' : ''} />
              <span><b data-text="delete-label"></b><br /><small class="muted" data-text="delete-hint"></small></span>
            </label>
            <label class="field limit-field" data-limit hidden>
              <span>Limite de exclusões por execução</span>
              <input type="number" name="maxDeletions" min="0" max="${MAX_LIMIT}" value="${t.maxDeletions ?? catalog.defaultMaxDeletions}" />
              <small>Proteção contra uma escolha errada: acima do limite, os arquivos encontrados só são listados (o relatório avisa). 0: sem limite.</small>
            </label>
            <div class="alert error" data-delete-confirm hidden>
              ${icon('alert')}
              <div>
                <b>Exclusão sem volta nas pastas do Windows:</b> arquivos excluídos pela rede não vão para a Lixeira. No OneDrive e no SharePoint, vale a forma definida em cada repositório (para a lixeira ou definitiva). <span data-text="check"></span>
                <p data-path-warning hidden><b>Atenção:</b> com "Caminho completo", um termo no nome de uma pasta faz todos os arquivos dela (e das subpastas) serem encontrados — e excluídos.</p>
                <p data-work-warning hidden><b>Atenção:</b> você marcou arquivos de trabalho (<span data-work-names></span>): todos os arquivos desses tipos nos repositórios escolhidos serão excluídos.</p>
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
  const mode = () => (form.elements.search.value === 'types' ? 'types' : 'terms');
  const setText = (key, value) => {
    form.querySelector(`[data-text="${key}"]`).textContent = value;
  };

  /** Limite de exclusões (null se inválido); em branco, o padrão. */
  const limitValue = () => {
    const raw = String(form.elements.maxDeletions.value || '').trim();
    if (raw === '') return form.elements.maxDeletions.validity.badInput ? null : catalog.defaultMaxDeletions;
    const n = Number(raw);
    return Number.isInteger(n) && n >= 0 && n <= MAX_LIMIT ? n : null;
  };
  /** Tamanho mínimo em MB (null se inválido); em branco, 0 (todos os tamanhos). */
  const minSizeValue = () => {
    const input = form.elements.minSizeMB;
    if (input.validity.badInput) return null;
    const raw = String(input.value || '').trim();
    if (raw === '') return 0;
    const n = Number(raw);
    return Number.isFinite(n) && n >= 0 ? n : null;
  };

  const sync = () => {
    const types = mode() === 'types';
    const T = MODE_TEXT[mode()];
    const deleting = form.elements.action.value === 'delete';
    const later = scheduling(form);
    form.querySelectorAll('[data-for-search]').forEach((el) => {
      el.hidden = el.dataset.forSearch !== mode();
    });
    // Sem as listas ao lado, os repositórios ocupam a linha inteira.
    form.querySelector('[data-repositories]').classList.toggle('full', types);
    setText('analyze-label', T.analyze[0]);
    setText('analyze-hint', T.analyze[1]);
    setText('delete-label', T.delete[0]);
    setText('delete-hint', T.delete[1]);
    setText('check', T.check);
    form.querySelector('[data-delete-confirm]').hidden = !deleting;
    form.querySelector('[data-limit]').hidden = !(types && deleting);
    form.querySelector('[data-path-warning]').hidden = types || !(form.elements.checkName.checked && form.elements.nameTarget.value === 'path');
    const workNames = types ? [...form.querySelectorAll('[name="categories"]:checked')].filter((el) => el.dataset.work).map((el) => catalog.categories.find((c) => c.key === el.value)?.label) : [];
    form.querySelector('[data-work-warning]').hidden = !workNames.length;
    form.querySelector('[data-work-names]').textContent = workNames.join(', ');
    form.querySelector('[data-schedule-only]').hidden = !later;
    form.querySelectorAll('[data-now-only]').forEach((el) => {
      el.hidden = later;
    });
    form.querySelector('[data-name-label]').textContent = later ? 'Nome do agendamento' : types ? 'Nome da busca (opcional)' : 'Nome da análise (opcional)';
    form.elements.name.maxLength = later ? 120 : 200;
    form.elements.name.placeholder = T.namePlaceholder[later ? 1 : 0];
    submit.className = `btn ${deleting ? 'danger' : 'primary'}`;
    const label = later ? (deleting ? 'Salvar agendamento com exclusão' : 'Salvar agendamento') : T.start[deleting ? 1 : 0];
    paint(submit, html`${icon(later ? 'clock' : 'play')} ${label}`);
  };
  const onChange = (event) => {
    if (['search', 'action', 'nameTarget', 'checkName', 'categories'].includes(event.target.name)) sync();
  };
  form.addEventListener('change', onChange);
  const unbind = bindSchedule(form, { onModeChange: sync, scheduleId: schedule?.id || null });

  const onSubmit = async (event) => {
    event.preventDefault();
    const f = new FormData(form);
    const types = mode() === 'types';
    const deleting = f.get('action') === 'delete';
    const later = scheduling(form);
    const options = {
      modifiedAfter: !later && f.get('modifiedAfter') ? `${f.get('modifiedAfter')}T00:00:00` : null,
      concurrency: Number(f.get('concurrency')),
      resolveOwner: f.get('resolveOwner') === 'on',
      deleteMatches: deleting,
    };
    const body = {
      kind: 'files',
      purpose: types ? 'types' : 'terms',
      name: String(f.get('name') || '').trim(),
      repositoryIds: f.getAll('repositoryIds'),
      confirmDelete: deleting ? String(f.get('confirmDelete') || '') : '',
    };
    if (types) {
      const minSizeMB = minSizeValue();
      let maxDeletions = limitValue();
      // O limite só vale com a exclusão (o campo fica escondido em "Somente procurar"): o salvo é mantido.
      if (maxDeletions === null && !deleting) maxDeletions = schedule?.fileTypes?.maxDeletions ?? catalog.defaultMaxDeletions;
      body.fileTypes = { categories: f.getAll('categories'), extensions: String(f.get('extensions') || ''), checkContent: f.get('realType') === 'on', minSizeMB, maxDeletions };
      body.options = options;
      if (!body.fileTypes.categories.length && !body.fileTypes.extensions.trim()) return toast('Escolha ao menos um tipo de arquivo ou informe uma extensão.', 'error');
      if (minSizeMB === null) {
        form.elements.minSizeMB.focus();
        return toast('Informe o tamanho mínimo em MB (0 = todos os tamanhos).', 'error');
      }
      if (maxDeletions === null) {
        form.elements.maxDeletions.focus();
        return toast('Informe o limite de exclusões por execução: um número inteiro (0 = sem limite).', 'error');
      }
    } else {
      body.listIds = f.getAll('listIds');
      body.options = {
        ...options,
        checkName: f.get('checkName') === 'on',
        nameTarget: f.get('nameTarget'),
        checkContent: f.get('checkContent') === 'on',
        maxFileSizeMB: Number(f.get('maxFileSizeMB')),
      };
    }
    if (later && !body.name) {
      form.elements.name.focus();
      return toast('Dê um nome ao agendamento.', 'error');
    }
    if (body.repositoryIds.length === 0) return toast('Selecione ao menos um repositório.', 'error');
    if (!types && body.listIds.length === 0) return toast('Selecione ao menos uma lista de referência.', 'error');
    if (deleting) {
      const blocked = repos.filter((r) => body.repositoryIds.includes(r.id) && !r.allowDelete).map((r) => r.name);
      if (blocked.length) return toast(`A exclusão não está permitida em: ${blocked.join(', ')}. Ative em Repositórios ou escolha "${MODE_TEXT[mode()].analyze[0]}".`, 'error');
      if (body.confirmDelete.trim().toUpperCase() !== 'EXCLUIR') {
        form.elements.confirmDelete.focus();
        return toast('Digite EXCLUIR para confirmar a exclusão.', 'error');
      }
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
      toast(types ? 'Busca por tipo iniciada.' : `Análise iniciada (${fmtNum(scan.summary.termCount)} termos).`, 'success');
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
