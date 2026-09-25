// Formulário de nova análise de e-mail.
import { get, post } from '../api.js';
import { html, render as paint, icon, toast, fmtNum, plural } from '../ui.js';
import { go } from '../nav.js';

const TYPE_LABELS = { graph: 'Microsoft 365', gmail: 'Google Workspace', imap: 'IMAP' };

function sourceDetail(s) {
  if (s.type !== 'imap' && s.scope === 'all') return `${TYPE_LABELS[s.type]} · todas as caixas`;
  return `${TYPE_LABELS[s.type] || s.type} · ${plural(s.mailboxes.length, 'caixa', 'caixas')}`;
}

export async function render(root, { ctx }) {
  const [sources, lists] = await Promise.all([get('/api/mail-sources'), get('/api/lists')]);
  const d = ctx.info.mailDefaults || {};
  const usable = lists.filter((l) => l.termCount > 0);

  if (sources.length === 0 || usable.length === 0) {
    paint(
      root,
      html`<div class="page-head"><div><h1>Nova análise de e-mail</h1></div></div>
        <div class="alert info">${icon('info')}<div>
          Para iniciar uma análise de e-mail é preciso ter ao menos uma conexão de e-mail e uma lista de referência com termos.
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

  paint(
    root,
    html`<div class="page-head">
        <div>
          <h1>Nova análise de e-mail</h1>
          <div class="sub">Procura os termos das listas de referência nas mensagens de todas as pastas das caixas escolhidas.</div>
        </div>
      </div>
      <form class="card" data-form novalidate>
        <div class="form-grid">
          <label class="field full">
            <span>Nome da análise (opcional)</span>
            <input type="text" name="name" maxlength="200" placeholder="Ex.: Varredura LGPD dos e-mails – setembro" />
          </label>

          <fieldset>
            <legend>Caixas de e-mail</legend>
            <div class="choice-list">
              ${sources.map(
                (s) => html`<label class="check">
                  <input type="checkbox" name="sourceIds" value="${s.id}" ${sources.length === 1 ? 'checked' : ''} />
                  <span><b>${s.name}</b><br /><span class="muted small">${sourceDetail(s)}</span></span>
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
              <label class="field">
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
        </div>
        <p class="hint">O acesso às caixas é somente leitura: nenhuma mensagem é alterada, movida ou marcada como lida.</p>
        <div class="inline page-actions">
          <button type="submit" class="btn primary">${icon('play')} Iniciar análise</button>
          <a class="btn" href="#/email/analises">Cancelar</a>
        </div>
      </form>`,
  );

  const form = root.querySelector('[data-form]');
  const onSubmit = async (event) => {
    event.preventDefault();
    const f = new FormData(form);
    const on = (name) => f.get(name) === 'on';
    const body = {
      kind: 'mail',
      name: f.get('name'),
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
        receivedAfter: f.get('receivedAfter') ? `${f.get('receivedAfter')}T00:00:00` : null,
        maxMessageSizeMB: Number(f.get('maxMessageSizeMB')),
        concurrency: Number(f.get('concurrency')),
      },
    };
    if (body.sourceIds.length === 0) return toast('Selecione ao menos uma conexão de e-mail.', 'error');
    if (body.listIds.length === 0) return toast('Selecione ao menos uma lista de referência.', 'error');
    const button = form.querySelector('[type="submit"]');
    button.disabled = true;
    try {
      const scan = await post('/api/scans', body);
      toast(`Análise iniciada (${fmtNum(scan.summary.termCount)} termos).`, 'success');
      go(`/email/analises/${scan.id}`);
    } catch (err) {
      toast(err.message, 'error');
      button.disabled = false;
    }
  };
  form.addEventListener('submit', onSubmit);
  return () => form.removeEventListener('submit', onSubmit);
}
