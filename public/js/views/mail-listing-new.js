// Formulário de nova listagem de e-mail (somente leitura): catálogo de contas do domínio ou listagem
// de mensagens por caixa. Sem termos e sem exclusão.
import { get, post } from '../api.js';
import { html, render as paint, icon, toast, fmtNum, plural } from '../ui.js';
import { go } from '../nav.js';

const TYPE_LABELS = { graph: 'Microsoft 365', gmail: 'Google Workspace', imap: 'IMAP' };

function sourceDetail(s) {
  if (s.type !== 'imap' && s.scope === 'all') return `${TYPE_LABELS[s.type]} · todas as caixas`;
  return `${TYPE_LABELS[s.type] || s.type} · ${plural(s.mailboxes.length, 'caixa', 'caixas')}`;
}

const KIND_INFO = {
  directory:
    'Lista todas as contas registradas no domínio das conexões, com os dados do cadastro (nome, endereço principal, apelidos, situação, tipo, licença, data de criação, departamento). No Microsoft 365 exige a permissão User.Read.All; no Google Workspace, um administrador (Admin SDK); no IMAP, traz apenas as caixas cadastradas na conexão.',
  messages:
    'Lista os dados de cada mensagem (remetente, destinatários, data, assunto, pasta, tamanho) de todas as pastas das caixas escolhidas. Lê só os cabeçalhos: o corpo e os anexos não são baixados.',
};

export async function render(root, { query }) {
  const sources = await get('/api/mail-sources');
  const back = '#/email/listagens';
  let kind = query.get('tipo') === 'directory' ? 'directory' : 'messages';

  if (sources.length === 0) {
    paint(
      root,
      html`<div class="page-head"><div><h1>Nova listagem de e-mail</h1></div></div>
        <div class="alert info">${icon('info')}<div>
          Para fazer uma listagem é preciso ter ao menos uma conexão de e-mail cadastrada.
          <div class="inline page-actions"><a class="btn small" href="#/email/caixas">Cadastrar caixas de e-mail</a></div>
        </div></div>`,
    );
    return null;
  }

  const check = (name, label, checked, hint = '') =>
    html`<label class="check"><input type="checkbox" name="${name}" ${checked ? 'checked' : ''} /><span><b>${label}</b>${hint ? html`<br /><small class="muted">${hint}</small>` : ''}</span></label>`;

  paint(
    root,
    html`<div class="page-head">
        <div>
          <h1>Nova listagem de e-mail</h1>
          <div class="sub">Inventário somente leitura: nada é procurado, alterado ou excluído.</div>
        </div>
      </div>
      <form class="card" data-form novalidate>
        <div class="form-grid">
          <fieldset class="full">
            <legend>O que listar</legend>
            <label class="check">
              <input type="radio" name="kind" value="directory" ${kind === 'directory' ? 'checked' : ''} />
              <span><b>Contas do domínio</b><br /><small class="muted">Todas as contas de e-mail registradas no domínio, com os dados de cada uma.</small></span>
            </label>
            <label class="check">
              <input type="radio" name="kind" value="messages" ${kind === 'messages' ? 'checked' : ''} />
              <span><b>Mensagens por caixa</b><br /><small class="muted">Todos os e-mails de cada caixa, com remetente, destinatários, data, assunto, pasta e tamanho.</small></span>
            </label>
            <div class="alert info" data-kind-info>${icon('info')}<div></div></div>
          </fieldset>

          <label class="field full">
            <span>Nome da listagem (opcional)</span>
            <input type="text" name="name" maxlength="200" placeholder="Ex.: Inventário de contas – setembro" />
          </label>

          <fieldset>
            <legend>Conexões de e-mail</legend>
            <div class="choice-list">
              ${sources.map(
                (s) => html`<label class="check">
                  <input type="checkbox" name="sourceIds" value="${s.id}" ${sources.length === 1 ? 'checked' : ''} />
                  <span><b>${s.name}</b><br /><span class="muted small">${sourceDetail(s)}</span></span>
                </label>`,
              )}
            </div>
          </fieldset>

          <fieldset class="full" data-messages-only>
            <legend>Filtros e desempenho (listagem de mensagens)</legend>
            <div class="form-grid">
              <label class="field">
                <span>Somente mensagens recebidas a partir de</span>
                <input type="date" name="receivedAfter" />
                <small>Em branco: todas as mensagens.</small>
              </label>
              <div class="field">
                ${check('includeTrash', 'Incluir a Lixeira (Itens Excluídos)', true)}
                ${check('includeJunk', 'Incluir o Lixo Eletrônico (spam)', false)}
              </div>
              <label class="field">
                <span>Mensagens lidas em paralelo</span>
                <input type="number" name="concurrency" min="1" max="8" value="4" />
                <small>O Microsoft 365 aceita até 4 por caixa; valores maiores podem causar esperas por limite de requisições.</small>
              </label>
            </div>
          </fieldset>
        </div>
        <div class="inline page-actions">
          <button type="submit" class="btn primary" data-submit>${icon('play')} Iniciar listagem</button>
          <a class="btn" href="${back}">Cancelar</a>
        </div>
      </form>`,
  );

  const form = root.querySelector('[data-form]');
  const submit = form.querySelector('[data-submit]');
  const sync = () => {
    kind = form.elements.kind.value;
    paint(form.querySelector('[data-kind-info] div'), html`${KIND_INFO[kind]}`);
    form.querySelector('[data-messages-only]').hidden = kind !== 'messages';
  };
  const onChange = (event) => {
    if (event.target.name === 'kind') sync();
  };
  form.addEventListener('change', onChange);
  sync();

  const onSubmit = async (event) => {
    event.preventDefault();
    const f = new FormData(form);
    const on = (name) => f.get(name) === 'on';
    const chosen = f.get('kind') === 'directory' ? 'directory' : 'messages';
    const sourceIds = f.getAll('sourceIds');
    if (sourceIds.length === 0) return toast('Selecione ao menos uma conexão de e-mail.', 'error');
    const body = {
      listing: { kind: chosen },
      name: String(f.get('name') || '').trim(),
      sourceIds,
      options:
        chosen === 'messages'
          ? {
              includeTrash: on('includeTrash'),
              includeJunk: on('includeJunk'),
              receivedAfter: f.get('receivedAfter') ? `${f.get('receivedAfter')}T00:00:00` : null,
              concurrency: Number(f.get('concurrency')),
            }
          : {},
    };
    submit.disabled = true;
    try {
      const scan = await post('/api/scans', body);
      toast('Listagem iniciada.', 'success');
      go(`/email/listagens/${scan.id}`);
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
