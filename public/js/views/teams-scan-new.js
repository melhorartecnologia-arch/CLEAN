// Formulário de nova análise do Microsoft Teams. Usa as conexões Microsoft 365 (Graph) já cadastradas
// em Caixas de e-mail (o mesmo aplicativo, com as permissões do Teams consentidas) e as listas de
// referência.
import { get, post } from '../api.js';
import { html, render as paint, icon, toast, fmtNum, plural } from '../ui.js';
import { go } from '../nav.js';

export async function render(root, { ctx }) {
  const [sources, lists] = await Promise.all([get('/api/mail-sources'), get('/api/lists')]);
  const graph = sources.filter((s) => s.type === 'graph');
  const usable = lists.filter((l) => l.termCount > 0);
  const back = '#/teams/analises';

  if (graph.length === 0 || usable.length === 0) {
    paint(
      root,
      html`<div class="page-head"><div><h1>Nova análise do Teams</h1></div></div>
        <div class="alert info">${icon('info')}<div>
          Para analisar o Microsoft Teams é preciso ter ao menos uma conexão <b>Microsoft 365</b> (o mesmo aplicativo do e-mail, com as permissões do Teams consentidas) e uma lista de referência com termos.
          <div class="inline page-actions">
            ${graph.length === 0 ? html`<a class="btn small" href="#/email/caixas">Cadastrar conexão Microsoft 365</a>` : ''}
            ${usable.length === 0 ? html`<a class="btn small" href="#/listas/nova">Criar lista de referência</a>` : ''}
          </div>
        </div></div>`,
    );
    return null;
  }

  const d = ctx.info?.teamsDefaults || {};
  const check = (name, label, checked, hint = '') =>
    html`<label class="check"><input type="checkbox" name="${name}" ${checked ? 'checked' : ''} /><span><b>${label}</b>${hint ? html`<br /><small class="muted">${hint}</small>` : ''}</span></label>`;

  paint(
    root,
    html`<div class="page-head">
        <div>
          <h1>Nova análise do Microsoft Teams</h1>
          <div class="sub">Procura os termos das listas de referência nas mensagens dos canais das equipes e dos chats (1:1, em grupo e de reunião).</div>
        </div>
      </div>
      <form class="card" data-form novalidate>
        <div class="form-grid">
          <label class="field full">
            <span>Nome da análise (opcional)</span>
            <input type="text" name="name" maxlength="200" placeholder="Ex.: Varredura LGPD do Teams – outubro" />
          </label>

          <fieldset>
            <legend>Conexões Microsoft 365</legend>
            <div class="choice-list">
              ${graph.map(
                (s) => html`<label class="check">
                  <input type="checkbox" name="sourceIds" value="${s.id}" ${graph.length === 1 ? 'checked' : ''} />
                  <span><b>${s.name}</b>${s.allowDelete ? html` <span class="chip danger">exclusão permitida</span>` : ''}<br /><span class="muted small">Microsoft 365${s.graph?.account ? ' · conta conectada' : s.graph?.auth === 'certificate' ? ' · certificado' : ' · segredo do cliente'}</span></span>
                </label>`,
              )}
            </div>
            <p class="muted small">A conta conectada vê só as conversas de quem entrou; para varrer todo o locatário, use um aplicativo (segredo ou certificado) com as permissões ChannelMessage.Read.All e Chat.Read.All (consentidas pelo administrador).</p>
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
            <legend>O que varrer no Teams</legend>
            <div class="form-grid">
              ${check('scanChannels', 'Canais das equipes', true, 'Mensagens dos canais (padrão, privados e compartilhados) e as respostas.')}
              ${check('scanChats', 'Chats', true, 'Conversas 1:1, em grupo e de reunião.')}
            </div>
            <label class="check"><input type="radio" name="scope" value="all" checked /><span><b>Todo o locatário</b><br /><small class="muted">Todas as equipes e, nos chats, todos os usuários.</small></span></label>
            <label class="check"><input type="radio" name="scope" value="list" /><span><b>Apenas uma lista</b><br /><small class="muted">Informe os identificadores das equipes e/ou os e-mails dos usuários (para os chats).</small></span></label>
            <div class="form-grid" data-list-only hidden>
              <label class="field"><span>Equipes (um id por linha)</span><textarea name="teamIds" rows="2" placeholder="id da equipe (GUID)"></textarea></label>
              <label class="field"><span>Usuários para os chats (um e-mail por linha)</span><textarea name="userEmails" rows="2" placeholder="pessoa@empresa.com"></textarea></label>
            </div>
          </fieldset>

          <fieldset class="full">
            <legend>Onde procurar em cada mensagem</legend>
            <div class="form-grid">
              ${check('checkSubject', 'Assunto', true)}
              ${check('checkBody', 'Corpo da mensagem', true)}
              ${check('checkAttachmentNames', 'Nomes dos anexos', true)}
              ${check('checkAttachments', 'Conteúdo dos anexos', true, 'Baixa e lê os arquivos anexados (SharePoint/OneDrive): Word, Excel, PDF, textos etc.')}
            </div>
          </fieldset>

          <fieldset class="full">
            <legend>Filtros e desempenho</legend>
            <div class="form-grid">
              <label class="field"><span>Somente mensagens a partir de</span><input type="date" name="receivedAfter" /><small>Em branco: todas as mensagens.</small></label>
              <label class="field"><span>Tamanho máximo por anexo (MB)</span><input type="number" name="maxMessageSizeMB" min="1" max="500" value="${d.maxMessageSizeMB || 50}" /><small>Anexos maiores: só o início é lido.</small></label>
              <label class="field"><span>Downloads em paralelo</span><input type="number" name="concurrency" min="1" max="8" value="${d.concurrency || 4}" /></label>
            </div>
          </fieldset>

          <fieldset class="full">
            <legend>O que fazer com as mensagens encontradas</legend>
            <label class="check"><input type="radio" name="action" value="analyze" checked /><span><b>Somente analisar</b><br /><small class="muted">Gera o relatório; nenhuma mensagem é alterada. Depois, se quiser, exclua pelo relatório.</small></span></label>
            <label class="check"><input type="radio" name="action" value="delete" /><span><b>Analisar e excluir automaticamente</b><br /><small class="muted">Exclui (softDelete, recuperável por um administrador) as mensagens de canal encontradas. Só para conexões com "Permitir exclusão". A exclusão de chats não é oferecida pelo Microsoft Graph.</small></span></label>
            <div class="alert error" data-delete-confirm hidden>${icon('alert')}
              <div><b>Exclusão de mensagens do Teams.</b> As mensagens de canal encontradas são removidas (softDelete: ficam recuperáveis por um administrador por um período). Confira as listas antes de continuar.
                <label class="field"><span>Digite EXCLUIR para confirmar</span><input type="text" name="confirmDelete" autocomplete="off" spellcheck="false" /></label>
              </div>
            </div>
          </fieldset>
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
    form.querySelector('[data-delete-confirm]').hidden = !deleting;
    form.querySelector('[data-list-only]').hidden = form.elements.scope.value !== 'list';
    submit.className = `btn ${deleting ? 'danger' : 'primary'}`;
    paint(submit, html`${icon('play')} ${deleting ? 'Iniciar análise e exclusão' : 'Iniciar análise'}`);
  };
  const onChange = (event) => {
    if (['action', 'scope'].includes(event.target.name)) sync();
  };
  form.addEventListener('change', onChange);
  sync();

  const onSubmit = async (event) => {
    event.preventDefault();
    const f = new FormData(form);
    const on = (name) => f.get(name) === 'on';
    const deleting = f.get('action') === 'delete';
    const sourceIds = f.getAll('sourceIds');
    const listIds = f.getAll('listIds');
    if (sourceIds.length === 0) return toast('Selecione ao menos uma conexão Microsoft 365.', 'error');
    if (listIds.length === 0) return toast('Selecione ao menos uma lista de referência.', 'error');
    if (!on('scanChannels') && !on('scanChats')) return toast('Escolha o que varrer: canais, chats ou os dois.', 'error');
    if (f.get('scope') === 'list') {
      if (on('scanChannels') && !String(f.get('teamIds') || '').trim()) return toast('No escopo por lista, informe ao menos uma equipe (ou desmarque "Canais das equipes").', 'error');
      if (on('scanChats') && !String(f.get('userEmails') || '').trim()) return toast('No escopo por lista, informe ao menos um usuário para os chats (ou desmarque "Chats").', 'error');
    }
    if (deleting) {
      const blocked = graph.filter((s) => sourceIds.includes(s.id) && !s.allowDelete).map((s) => s.name);
      if (blocked.length) return toast(`A exclusão não está permitida em: ${blocked.join(', ')}. Ative em Caixas de e-mail ou escolha "Somente analisar".`, 'error');
      if (String(f.get('confirmDelete') || '').trim().toUpperCase() !== 'EXCLUIR') return toast('Digite EXCLUIR para confirmar a exclusão.', 'error');
    }
    const body = {
      kind: 'teams',
      name: String(f.get('name') || '').trim(),
      sourceIds,
      listIds,
      teams: { scope: f.get('scope') === 'list' ? 'list' : 'all', teamIds: String(f.get('teamIds') || ''), userEmails: String(f.get('userEmails') || '') },
      options: {
        scanChannels: on('scanChannels'),
        scanChats: on('scanChats'),
        checkSubject: on('checkSubject'),
        checkBody: on('checkBody'),
        checkAttachmentNames: on('checkAttachmentNames'),
        checkAttachments: on('checkAttachments'),
        receivedAfter: f.get('receivedAfter') ? `${f.get('receivedAfter')}T00:00:00` : null,
        maxMessageSizeMB: Number(f.get('maxMessageSizeMB')),
        concurrency: Number(f.get('concurrency')),
        deleteMatches: deleting,
      },
      confirmDelete: deleting ? String(f.get('confirmDelete') || '') : '',
    };
    submit.disabled = true;
    try {
      const scan = await post('/api/scans', body);
      toast(`Análise iniciada (${fmtNum(scan.summary.termCount)} termos).`, 'success');
      go(`/teams/analises/${scan.id}`);
    } catch (err) {
      toast(err.message, 'error');
      submit.disabled = false;
    }
  };
  form.addEventListener('submit', onSubmit);
  return () => {
    form.removeEventListener('submit', onSubmit);
    form.removeEventListener('change', onChange);
  };
}
