// Nova política de retenção (#/retencao/nova?tipo=email) ou edição (#/retencao/<id>): arquivos ou
// mensagens mais antigos que a idade máxima, pelo critério de data escolhido, são listados no
// relatório e — se a política excluir — eliminados. Execução manual ou agendada.
import { get, post, put } from '../api.js';
import { html, render as paint, icon, toast, fmtServerDateTime, plural } from '../ui.js';
import { go } from '../nav.js';
import { scheduleSection, bindSchedule, readSchedule, scheduling, keepField } from '../schedule-form.js';
import { FILE_CRITERIA, MAIL_CRITERIA, UNITS, DEFAULT_MAX_DELETIONS, describeRetention, cutoffPreview } from '../retention.js';

const CLOUD_LABELS = { onedrive: 'OneDrive', sharepoint: 'SharePoint' };
const TYPE_LABELS = { graph: 'Microsoft 365', gmail: 'Google Workspace', imap: 'IMAP' };

function sourceDetail(s) {
  if (s.type !== 'imap' && s.scope === 'all') return `${TYPE_LABELS[s.type]} · todas as caixas`;
  return `${TYPE_LABELS[s.type] || s.type} · ${plural(s.mailboxes.length, 'caixa', 'caixas')}`;
}

/** O que expira, na prévia: "arquivos modificados pela última vez antes de 26/09/2021". */
const EXPIRES = {
  used: (d) => `arquivos sem nenhuma data (modificação, último acesso ou criação) a partir de ${d}`,
  modified: (d) => `arquivos modificados pela última vez antes de ${d}`,
  accessed: (d) => `arquivos acessados pela última vez antes de ${d}`,
  created: (d) => `arquivos criados antes de ${d}`,
  received: (d) => `mensagens recebidas antes de ${d}`,
};

export async function render(root, { params, query, ctx }) {
  const id = params[0];
  const policy = id ? await get(`/api/schedules/${encodeURIComponent(id)}`) : null;
  if (policy && policy.purpose !== 'retention') {
    // Um agendamento de análise aberto pelo endereço das políticas.
    go(`/agendamentos/${policy.id}`);
    return null;
  }
  const kind = policy ? policy.kind : query.get('tipo') === 'email' ? 'mail' : 'files';
  const mail = kind === 'mail';
  const targets = await get(mail ? '/api/mail-sources' : '/api/repositories');
  const items = mail ? 'mensagens' : 'arquivos';
  const o = mail ? 'a' : 'o'; // gênero: "mensagens expiradas", "arquivos expirados"
  const title = policy ? 'Editar política de retenção' : mail ? 'Nova política de retenção de e-mails' : 'Nova política de retenção de arquivos';

  if (targets.length === 0) {
    paint(
      root,
      html`<div class="page-head"><div><h1>${title}</h1></div></div>
        <div class="alert info">${icon('info')}<div>
          Para criar a política é preciso ter ao menos ${mail ? 'uma conexão de e-mail' : 'um repositório'} cadastrad${o}.
          <div class="inline page-actions"><a class="btn small" href="${mail ? '#/email/caixas' : '#/repositorios'}">${mail ? 'Cadastrar conexão de e-mail' : 'Cadastrar repositório'}</a></div>
        </div></div>`,
    );
    return null;
  }

  const r = policy?.retention || {};
  const criteria = mail ? MAIL_CRITERIA : FILE_CRITERIA;
  const criterion = r.criterion || (mail ? 'received' : 'used');
  const deleting = policy?.action === 'delete';
  const chosen = (tid) => (policy ? policy.targetIds.includes(tid) : targets.length === 1);
  const chk = (value) => (value ? 'checked' : '');
  const sel = (a, b) => (String(a) === String(b) ? 'selected' : '');
  const isCloud = (t) => Boolean(CLOUD_LABELS[t.type]);

  paint(
    root,
    html`<div class="page-head">
        <div>
          <h1>${title}</h1>
          <div class="sub">${mail ? 'Mensagens' : 'Arquivos'} mais antig${o}s que a idade máxima são listad${o}s no relatório e, se a política excluir, eliminad${o}s.</div>
        </div>
      </div>
      <form class="card" data-form novalidate>
        <div class="form-grid">
          <label class="field full">
            <span>Nome da política</span>
            <input type="text" name="name" maxlength="120" required value="${policy?.name || ''}" placeholder="${mail ? 'Ex.: E-mails com mais de 5 anos' : 'Ex.: Arquivos sem uso há mais de 5 anos'}" />
          </label>

          <fieldset>
            <legend>${mail ? 'Conexões de e-mail' : 'Repositórios'}</legend>
            <div class="choice-list">
              ${targets.map(
                (t) => html`<label class="check">
                  <input type="checkbox" name="targetIds" value="${t.id}" data-cloud="${isCloud(t) ? '1' : ''}" ${chk(chosen(t.id))} />
                  <span><b>${t.name}</b>${!mail && isCloud(t) ? html` <span class="chip">${CLOUD_LABELS[t.type]}</span>` : ''}${t.allowDelete ? html` <span class="chip danger">exclusão permitida</span>` : ''}<br />${mail
                      ? html`<span class="muted small">${sourceDetail(t)}</span>`
                      : html`<span class="${isCloud(t) ? 'muted small' : 'mono muted'}">${t.path}</span>`}</span>
                </label>`,
              )}
            </div>
          </fieldset>

          <fieldset>
            <legend>Regra de idade</legend>
            <div class="stack">
            ${mail
              ? html`<div class="field"><span>Critério de data</span><b>${MAIL_CRITERIA.received.label}</b><small>${MAIL_CRITERIA.received.hint}</small></div>`
              : html`<label class="field">
                  <span>Critério de data</span>
                  <select name="criterion">${Object.entries(FILE_CRITERIA).map(([value, c]) => html`<option value="${value}" ${sel(criterion, value)}>${c.label}</option>`)}</select>
                  <small data-criterion-hint>${criteria[criterion]?.hint}</small>
                </label>`}
            <div class="field">
              <span id="age-label">Idade máxima</span>
              <div class="option-row" role="group" aria-labelledby="age-label">
                <input type="number" name="amount" min="1" max="${UNITS[r.unit || 'years'].max}" value="${r.amount || 5}" aria-label="Idade máxima (quantidade)" />
                <select name="unit" aria-label="Unidade da idade máxima">
                  ${Object.entries(UNITS).map(([value, u]) => html`<option value="${value}" ${sel(r.unit || 'years', value)}>${u.many}</option>`)}
                </select>
              </div>
              <small>${mail ? 'Mensagens' : 'Arquivos'} com a data do critério mais antiga que isso expiram.</small>
            </div>
            ${mail
              ? html`<label class="check"><input type="checkbox" name="includeTrash" ${chk(r.includeTrash !== false)} /><span>Incluir a Lixeira (Itens Excluídos)</span></label>
                  <label class="check"><input type="checkbox" name="includeJunk" ${chk(r.includeJunk !== false)} /><span>Incluir o Lixo Eletrônico (spam)</span></label>`
              : html`<label class="field">
                  <span>Somente os arquivos com estes nomes (opcional)</span>
                  <textarea name="patterns" rows="3" spellcheck="false" placeholder="*.tmp&#10;*.bak&#10;~$*">${(r.patterns || []).join('\n')}</textarea>
                  <small>Um padrão por linha: * vale qualquer sequência e ? um caractere (ex.: *.tmp). Em branco: todos os arquivos.</small>
                </label>`}
            </div>
          </fieldset>

          <div class="schedule-preview full" data-rule-preview aria-live="polite"></div>

          ${mail
            ? ''
            : html`<div class="alert full" data-access-warning hidden>
                  ${icon('alert')}
                  <div>
                    <b>O último acesso só é confiável se o Windows do servidor de arquivos o registrar</b> — em muitos servidores esse registro fica desligado.
                    Para conferir, rode no servidor <span class="mono">fsutil behavior query disablelastaccess</span>: 0 ou 2, registra; 1 ou 3, não registra.
                    Sem o registro, a data fica parada (na criação ou na cópia do arquivo) e um arquivo aberto todos os dias parece antigo — e seria excluído.
                    Programas que leem todos os arquivos (antivírus, backup, indexação e as análises de conteúdo do CLEAN) também podem atualizar o último acesso.
                    Na dúvida, prefira <b>Sem uso</b>: o arquivo só expira se nenhuma data for recente.
                  </div>
                </div>
                <div class="alert error full" data-cloud-warning hidden>${icon('alert')}<div>O OneDrive e o SharePoint não informam o último acesso de cada arquivo: escolha outro critério ou desmarque os repositórios da nuvem.</div></div>`}

          <fieldset class="full">
            <legend>O que fazer com ${o}s ${items} expirad${o}s</legend>
            <label class="check">
              <input type="radio" name="action" value="analyze" ${chk(!deleting)} />
              <span><b>Somente listar (simulação)</b><br /><small class="muted">Cada execução gera o relatório com ${o}s ${items} que seriam excluíd${o}s. Nada é excluído.</small></span>
            </label>
            <label class="check">
              <input type="radio" name="action" value="delete" ${chk(deleting)} />
              <span><b>Excluir ${o}s ${items} expirad${o}s</b><br /><small class="muted">Sem confirmação item a item, só em ${mail ? 'conexões' : 'repositórios'} com "Permitir exclusão". Antes, use "Simular agora" na lista de políticas para ver o que seria excluído.</small></span>
            </label>
            <div class="form-grid" data-delete-options hidden>
              <fieldset class="plain">
                <legend>Forma de exclusão</legend>
                <label class="check">
                  <input type="radio" name="deleteMode" value="permanent" ${chk(r.deleteMode !== 'trash')} />
                  <span><b>Definitiva</b><br /><small class="muted">${mail ? 'A mensagem não fica na Lixeira da caixa.' : 'O arquivo não fica em nenhuma lixeira.'}</small></span>
                </label>
                <label class="check">
                  <input type="radio" name="deleteMode" value="trash" ${chk(r.deleteMode === 'trash')} />
                  <span><b>Para a lixeira</b><br /><small class="muted">${mail
                    ? 'Vai para Itens Excluídos (Microsoft 365) ou para a Lixeira (Google e IMAP), de onde ainda pode ser recuperada.'
                    : 'Vale no OneDrive e no SharePoint (Lixeira do site). Nas pastas do Windows a exclusão é sempre definitiva: arquivos excluídos pela rede não vão para a Lixeira.'}</small></span>
                </label>
              </fieldset>
              <label class="field">
                <span>Limite de exclusões por execução</span>
                <input type="number" name="maxDeletions" min="0" max="10000000" value="${r.maxDeletions ?? DEFAULT_MAX_DELETIONS}" />
                <small>Proteção contra uma regra errada: acima do limite, ${o}s ${items} expirad${o}s só são listad${o}s (o relatório avisa). 0: sem limite.</small>
              </label>
              <div class="alert error full">
                ${icon('alert')}
                <div>
                  <b>${mail ? 'Exclusão das mensagens expiradas' : 'Exclusão sem volta nas pastas do Windows'}:</b> tudo o que a regra considerar expirado será excluído, em todas as execuções, sem nova confirmação.
                  Cada execução confere o cadastro: se ${mail ? 'uma conexão' : 'um repositório'} deixar de permitir a exclusão, mudar de ${mail ? 'caixas' : 'caminho'} ou de forma de exclusão, as execuções não excluem até a política ser salva e confirmada de novo.
                  <label class="field"><span>Digite EXCLUIR para confirmar${deleting ? ' (de novo, a cada vez que a política é salva)' : ''}</span><input type="text" name="confirmDelete" autocomplete="off" spellcheck="false" /></label>
                </div>
              </div>
            </div>
          </fieldset>

          ${mail
            ? ''
            : html`<fieldset class="full">
                <legend>Relatório</legend>
                <label class="check">
                  <input type="checkbox" name="resolveOwner" ${chk(policy ? policy.options?.resolveOwner !== false : true)} />
                  <span><b>Identificar o proprietário de cada arquivo expirado (NTFS)</b><br /><small class="muted">Pastas do Windows: mostra de quem são os arquivos. No OneDrive e no SharePoint, o último usuário vem do Microsoft 365.</small></span>
                </label>
              </fieldset>`}

          ${scheduleSection({ kind, schedule: policy, info: ctx.info, manual: true, period: false, keep: false })}

          <div class="form-grid full">${keepField(policy, 'desta política')}</div>
        </div>
        <div class="inline page-actions">
          <button type="submit" class="btn primary" data-submit>${icon('check')} Salvar política</button>
          <a class="btn" href="#/retencao">Cancelar</a>
        </div>
      </form>`,
  );

  const form = root.querySelector('[data-form]');
  const submit = form.querySelector('[data-submit]');
  const preview = form.querySelector('[data-rule-preview]');
  const currentCriterion = () => (mail ? 'received' : form.elements.criterion.value);
  const cloudChosen = () => [...form.querySelectorAll('[name="targetIds"]:checked')].some((el) => el.dataset.cloud);

  /** A regra como a API espera. */
  const readRetention = () => {
    const f = new FormData(form);
    return {
      criterion: currentCriterion(),
      amount: Number(f.get('amount')),
      unit: f.get('unit'),
      ...(mail
        ? { includeTrash: f.get('includeTrash') === 'on', includeJunk: f.get('includeJunk') === 'on' }
        : { patterns: String(f.get('patterns') || '').split(/\r?\n/).map((p) => p.trim()).filter(Boolean) }),
      maxDeletions: f.get('maxDeletions') === '' ? DEFAULT_MAX_DELETIONS : Number(f.get('maxDeletions')),
      deleteMode: f.get('deleteMode') === 'trash' ? 'trash' : 'permanent',
    };
  };

  const sync = () => {
    const deletingNow = form.elements.action.value === 'delete';
    const retention = readRetention();
    const c = currentCriterion();
    form.querySelector('[data-delete-options]').hidden = !deletingNow;
    if (!mail) {
      form.querySelector('[data-criterion-hint]').textContent = FILE_CRITERIA[c]?.hint || '';
      form.querySelector('[data-access-warning]').hidden = c !== 'accessed';
      form.querySelector('[data-cloud-warning]').hidden = !(c === 'accessed' && cloudChosen());
    }
    form.elements.amount.max = String(UNITS[retention.unit]?.max || 100);
    const cutoff = cutoffPreview(retention);
    paint(
      preview,
      cutoff
        ? html`<div class="preview-title">${icon('clock')} <b>${describeRetention(retention, kind)}</b></div>
            <div class="muted small">Se executada hoje, a política ${deletingNow ? 'excluiria' : 'listaria'} os ${EXPIRES[c](cutoff.toLocaleDateString('pt-BR'))}.${deletingNow && retention.maxDeletions ? ` No máximo ${retention.maxDeletions.toLocaleString('pt-BR')} exclusões por execução.` : ''}</div>`
        : html`<div class="danger-text small">${icon('alert')} Informe a idade máxima: de 1 a ${UNITS[retention.unit]?.max || 100} ${UNITS[retention.unit]?.many || 'anos'}.</div>`,
    );
    submit.className = `btn ${deletingNow ? 'danger' : 'primary'}`;
    paint(submit, html`${icon(deletingNow ? 'trash' : 'check')} ${deletingNow ? 'Salvar política com exclusão' : 'Salvar política'}`);
  };

  const onChange = (event) => {
    if (event.target.name !== 'confirmDelete' && event.target.name !== 'name') sync();
  };
  form.addEventListener('change', onChange);
  form.addEventListener('input', onChange);
  const unbind = bindSchedule(form, { onModeChange: sync, scheduleId: policy?.id || null });

  const onSubmit = async (event) => {
    event.preventDefault();
    const f = new FormData(form);
    const deletingNow = f.get('action') === 'delete';
    const later = scheduling(form);
    const targetIds = f.getAll('targetIds');
    const retention = readRetention();
    const { rule, catchUp, keepLast } = readSchedule(form);
    const body = {
      purpose: 'retention',
      kind,
      name: String(f.get('name') || '').trim(),
      [mail ? 'sourceIds' : 'repositoryIds']: targetIds,
      retention,
      options: mail ? {} : { resolveOwner: f.get('resolveOwner') === 'on' },
      action: deletingNow ? 'delete' : 'analyze',
      confirmDelete: deletingNow ? String(f.get('confirmDelete') || '') : '',
      rule: later ? rule : null,
      catchUp,
      keepLast,
    };
    if (!body.name) {
      form.elements.name.focus();
      return toast('Dê um nome à política.', 'error');
    }
    if (targetIds.length === 0) return toast(mail ? 'Selecione ao menos uma conexão de e-mail.' : 'Selecione ao menos um repositório.', 'error');
    if (!cutoffPreview(retention)) {
      form.elements.amount.focus();
      return toast(`Informe a idade máxima: de 1 a ${UNITS[retention.unit].max} ${UNITS[retention.unit].many}.`, 'error');
    }
    if (!mail && retention.criterion === 'accessed' && cloudChosen()) return toast('O OneDrive e o SharePoint não informam o último acesso: escolha outro critério.', 'error');
    if (deletingNow) {
      const blocked = targets.filter((t) => targetIds.includes(t.id) && !t.allowDelete).map((t) => t.name);
      if (blocked.length) {
        return toast(`A exclusão não está permitida em: ${blocked.join(', ')}. Ative em ${mail ? 'Caixas de e-mail' : 'Repositórios'} ou escolha "Somente listar (simulação)".`, 'error');
      }
      if (body.confirmDelete.trim().toUpperCase() !== 'EXCLUIR') {
        form.elements.confirmDelete.focus();
        return toast('Digite EXCLUIR para confirmar a exclusão.', 'error');
      }
    }
    submit.disabled = true;
    try {
      const saved = await (policy ? put(`/api/schedules/${policy.id}`, body) : post('/api/schedules', body));
      const next = saved.nextRunAt ? ` Próxima execução: ${fmtServerDateTime(saved.nextRunAt)}.` : saved.state === 'paused' ? ' Ela está pausada.' : '';
      toast(`Política salva.${next}`, 'success');
      go('/retencao');
    } catch (err) {
      toast(err.message, 'error');
      submit.disabled = false;
    }
  };
  form.addEventListener('submit', onSubmit);
  sync();
  return () => {
    unbind();
    form.removeEventListener('submit', onSubmit);
    form.removeEventListener('change', onChange);
    form.removeEventListener('input', onChange);
  };
}
