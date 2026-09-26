// Formulário de nova análise de e-mail: agora ou agendada (também edita agendamentos).
import { get, post, put } from '../api.js';
import { html, render as paint, icon, toast, fmtNum, plural, fmtServerDateTime } from '../ui.js';
import { go } from '../nav.js';
import { scheduleSection, bindSchedule, readSchedule, scheduling } from '../schedule-form.js';

const TYPE_LABELS = { graph: 'Microsoft 365', gmail: 'Google Workspace', imap: 'IMAP' };

function sourceDetail(s) {
  if (s.type !== 'imap' && s.scope === 'all') return `${TYPE_LABELS[s.type]} · todas as caixas`;
  return `${TYPE_LABELS[s.type] || s.type} · ${plural(s.mailboxes.length, 'caixa', 'caixas')}`;
}

/**
 * props.scheduling: tela só de agendamento (vinda de Agendamentos); props.schedule: agendamento em
 * edição.
 */
export async function render(root, { ctx, props = {} }) {
  const [sources, lists] = await Promise.all([get('/api/mail-sources'), get('/api/lists')]);
  const schedule = props.schedule || null;
  const fixed = Boolean(props.scheduling || schedule);
  const d = schedule?.options || ctx.info.mailDefaults || {};
  const usable = lists.filter((l) => l.termCount > 0);
  const title = schedule ? 'Editar agendamento' : fixed ? 'Novo agendamento de e-mail' : 'Nova análise de e-mail';
  const back = fixed ? '#/agendamentos' : '#/email/analises';

  if (sources.length === 0 || usable.length === 0) {
    paint(
      root,
      html`<div class="page-head"><div><h1>${title}</h1></div></div>
        <div class="alert info">${icon('info')}<div>
          Para ${fixed ? 'agendar' : 'iniciar'} uma análise de e-mail é preciso ter ao menos uma conexão de e-mail e uma lista de referência com termos.
          <div class="inline page-actions">
            ${sources.length === 0 ? html`<a class="btn small" href="#/email/caixas">Cadastrar caixas de e-mail</a>` : ''}
            ${usable.length === 0 ? html`<a class="btn small" href="#/listas/nova">Criar lista de referência</a>` : ''}
          </div>
        </div></div>`,
    );
    return null;
  }

  const check = (name, label, hint = '') =>
    html`<label class="check"><input type="checkbox" name="${name}" ${d[name] ? 'checked' : ''} /><span><b>${label}</b>${hint ? html`<br /><small class="muted">${hint}</small>` : ''}</span></label>`;
  const chosen = (list, id) => (schedule ? (list || []).includes(id) : null);
  const section = scheduleSection({ kind: 'mail', schedule, fixed, info: ctx.info });

  paint(
    root,
    html`<div class="page-head">
        <div>
          <h1>${title}</h1>
          <div class="sub">Procura os termos das listas de referência nas mensagens de todas as pastas das caixas escolhidas.</div>
        </div>
      </div>
      <form class="card" data-form novalidate>
        <div class="form-grid">
          <label class="field full">
            <span data-name-label>Nome da análise (opcional)</span>
            <input type="text" name="name" maxlength="200" value="${schedule?.name || ''}" placeholder="Ex.: Varredura LGPD dos e-mails – setembro" />
          </label>
          ${fixed ? section : ''}

          <fieldset>
            <legend>Caixas de e-mail</legend>
            <div class="choice-list">
              ${sources.map(
                (s) => html`<label class="check">
                  <input type="checkbox" name="sourceIds" value="${s.id}" ${(chosen(schedule?.targetIds, s.id) ?? sources.length === 1) ? 'checked' : ''} />
                  <span><b>${s.name}</b>${s.allowDelete ? html` <span class="chip danger">exclusão ${s.deleteMode === 'trash' ? 'para a lixeira' : 'definitiva'}</span>` : ''}<br /><span class="muted small">${sourceDetail(s)}</span></span>
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
            <legend>Onde procurar em cada mensagem</legend>
            <div class="form-grid">
              ${check('checkSubject', 'Assunto')}
              ${check('checkBody', 'Corpo da mensagem', 'Texto e HTML, inclusive mensagens respondidas e encaminhadas no corpo.')}
              ${check('checkAttachmentNames', 'Nomes dos anexos')}
              ${check('checkAttachments', 'Conteúdo dos anexos', 'Word, Excel, PowerPoint (novos e 97-2003), PDF, OpenDocument, RTF, textos, CSV, HTML, e-mails anexados (.eml/.msg) e nomes dentro de .zip.')}
              ${check('checkAddresses', 'Remetente e destinatários', 'Nomes e endereços de De, Para, Cc e Cco.')}
            </div>
          </fieldset>

          <fieldset class="full">
            <legend>Filtros e desempenho</legend>
            <div class="form-grid">
              <label class="field" data-now-only>
                <span>Somente mensagens recebidas a partir de</span>
                <input type="date" name="receivedAfter" />
                <small>Em branco: todas as mensagens.</small>
              </label>
              <div class="field">
                ${check('includeTrash', 'Incluir a Lixeira (Itens Excluídos)')}
                ${check('includeJunk', 'Incluir o Lixo Eletrônico (spam)')}
              </div>
              <label class="field">
                <span>Tamanho máximo por mensagem (MB)</span>
                <input type="number" name="maxMessageSizeMB" min="1" max="500" value="${d.maxMessageSizeMB || 50}" />
                <small>Mensagens maiores: só o início é baixado e analisado (os anexos que ficarem de fora têm apenas o nome verificado).</small>
              </label>
              <label class="field">
                <span>Mensagens baixadas em paralelo</span>
                <input type="number" name="concurrency" min="1" max="8" value="${d.concurrency || 4}" />
                <small>O Microsoft 365 aceita até 4 por caixa; valores maiores podem causar esperas por limite de requisições.</small>
              </label>
            </div>
          </fieldset>
          <fieldset class="full">
            <legend>O que fazer com as mensagens encontradas</legend>
            <label class="check">
              <input type="radio" name="action" value="analyze" ${schedule?.action === 'delete' ? '' : 'checked'} />
              <span><b>Somente analisar</b><br /><small class="muted">Gera o relatório; nenhuma mensagem é alterada, movida ou marcada como lida. Depois, se quiser, exclua item a item pelo relatório.</small></span>
            </label>
            <label class="check">
              <input type="radio" name="action" value="delete" ${schedule?.action === 'delete' ? 'checked' : ''} />
              <span><b>Analisar e excluir automaticamente</b><br /><small class="muted">Toda mensagem em que algum termo for encontrado é excluída ao fim de cada caixa, sem confirmação item a item, da forma definida em cada conexão (definitiva ou para a lixeira). Só para conexões com "Permitir exclusão".</small></span>
            </label>
            <div class="alert error" data-delete-confirm hidden>
              ${icon('alert')}
              <div>
                <b>Exclusão sem volta.</b> Nas conexões com exclusão definitiva, as mensagens não ficam na lixeira do usuário. Confira as listas de referência antes de continuar: tudo o que for encontrado será excluído, inclusive a mensagem inteira quando o termo estiver só em um anexo.
                <p data-schedule-only hidden><b>Agendamento:</b> a exclusão vale para todas as execuções, sem nova confirmação. Ela é conferida em cada execução: se uma conexão deixar de permitir a exclusão, mudar de conta, servidor ou caixas, ou passar a excluir de forma definitiva, as execuções falham até o agendamento ser salvo e confirmado de novo.</p>
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
    form.querySelector('[data-schedule-only]').hidden = !later;
    form.querySelectorAll('[data-now-only]').forEach((el) => {
      el.hidden = later;
    });
    form.querySelector('[data-name-label]').textContent = later ? 'Nome do agendamento' : 'Nome da análise (opcional)';
    form.elements.name.maxLength = later ? 120 : 200;
    form.elements.name.placeholder = later ? 'Ex.: Varredura LGPD dos e-mails – diária' : 'Ex.: Varredura LGPD dos e-mails – setembro';
    submit.className = `btn ${deleting ? 'danger' : 'primary'}`;
    const label = later ? (deleting ? 'Salvar agendamento com exclusão' : 'Salvar agendamento') : deleting ? 'Iniciar análise e exclusão' : 'Iniciar análise';
    paint(submit, html`${icon(later ? 'clock' : 'play')} ${label}`);
  };
  const onChange = (event) => {
    if (event.target.name === 'action') sync();
  };
  form.addEventListener('change', onChange);
  const unbind = bindSchedule(form, { onModeChange: sync, scheduleId: schedule?.id || null });
  const onSubmit = async (event) => {
    event.preventDefault();
    const f = new FormData(form);
    const on = (name) => f.get(name) === 'on';
    const deleting = f.get('action') === 'delete';
    const later = scheduling(form);
    const body = {
      kind: 'mail',
      name: String(f.get('name') || '').trim(),
      sourceIds: f.getAll('sourceIds'),
      listIds: f.getAll('listIds'),
      options: {
        checkSubject: on('checkSubject'),
        checkBody: on('checkBody'),
        checkAttachmentNames: on('checkAttachmentNames'),
        checkAttachments: on('checkAttachments'),
        checkAddresses: on('checkAddresses'),
        includeTrash: on('includeTrash'),
        includeJunk: on('includeJunk'),
        receivedAfter: !later && f.get('receivedAfter') ? `${f.get('receivedAfter')}T00:00:00` : null,
        maxMessageSizeMB: Number(f.get('maxMessageSizeMB')),
        concurrency: Number(f.get('concurrency')),
        deleteMatches: deleting,
      },
      confirmDelete: deleting ? String(f.get('confirmDelete') || '') : '',
    };
    if (later && !body.name) {
      form.elements.name.focus();
      return toast('Dê um nome ao agendamento.', 'error');
    }
    if (body.sourceIds.length === 0) return toast('Selecione ao menos uma conexão de e-mail.', 'error');
    if (body.listIds.length === 0) return toast('Selecione ao menos uma lista de referência.', 'error');
    if (deleting) {
      const blocked = sources.filter((s) => body.sourceIds.includes(s.id) && !s.allowDelete).map((s) => s.name);
      if (blocked.length) return toast(`A exclusão não está permitida em: ${blocked.join(', ')}. Ative em Caixas de e-mail ou escolha "Somente analisar".`, 'error');
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
      go(`/email/analises/${scan.id}`);
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
