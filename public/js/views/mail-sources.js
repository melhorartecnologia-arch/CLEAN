// Caixas de e-mail: conexões com o Microsoft 365, o Google Workspace ou servidores IMAP.
import { get, post, put, del } from '../api.js';
import { html, render as paint, icon, openDialog, confirmDialog, toast, fmtNum, plural } from '../ui.js';

const TYPES = {
  graph: { label: 'Microsoft 365', detail: 'Exchange Online (Outlook)', domain: 'do locatário' },
  gmail: { label: 'Google Workspace', detail: 'Gmail da empresa', domain: 'do domínio' },
  imap: { label: 'IMAP', detail: 'Exchange local, Zimbra, hospedagem…' },
};
const SECURITY = {
  tls: { label: 'SSL/TLS', port: 993 },
  starttls: { label: 'STARTTLS', port: 143 },
  none: { label: 'Sem criptografia (não recomendado)', port: 143 },
};
const GMAIL_SCOPES = 'https://www.googleapis.com/auth/gmail.readonly,https://www.googleapis.com/auth/admin.directory.user.readonly';
const SAVED = '•••••• salvo — deixe em branco para manter';

function mailboxRow(m = {}) {
  return html`<tr data-mailbox-row>
    <td><input type="email" name="mbAddress" value="${m.address || ''}" placeholder="nome@empresa.com.br" aria-label="E-mail da caixa" /></td>
    <td><input type="text" name="mbLogin" value="${m.login || ''}" placeholder="igual ao e-mail" aria-label="Login da caixa" /></td>
    <td><input type="password" name="mbPassword" autocomplete="new-password" placeholder="${m.hasPassword ? 'salva' : 'usa a senha padrão'}" aria-label="Senha da caixa" /></td>
    <td><button type="button" class="icon-btn danger" data-action="remove-row" aria-label="Remover caixa" title="Remover">${icon('x')}</button></td>
  </tr>`;
}

function sourceForm(src) {
  const type = src?.type || 'graph';
  const g = src?.graph || {};
  const gm = src?.gmail || {};
  const im = src?.imap || {};
  const scope = src?.scope || 'all';
  const list = type !== 'imap' ? (src?.mailboxes || []).map((m) => m.address).join('\n') : '';
  const imapRows = type === 'imap' && src?.mailboxes?.length ? src.mailboxes : [{}];
  return html`<div class="form-grid">
    <label class="field full">
      <span>Nome da conexão</span>
      <input type="text" name="name" required maxlength="200" value="${src?.name || ''}" placeholder="Ex.: E-mail corporativo" />
    </label>
    <fieldset class="full">
      <legend>Tipo</legend>
      <div class="type-choice">
        ${Object.entries(TYPES).map(
          ([value, t]) => html`<label>
            <input type="radio" name="type" value="${value}" ${type === value ? 'checked' : ''} />
            <span><b>${t.label}</b><small>${t.detail}</small></span>
          </label>`,
        )}
      </div>
    </fieldset>

    <fieldset class="full" data-type="graph">
      <legend>Microsoft 365 (Microsoft Graph)</legend>
      <div class="form-grid">
        <label class="field">
          <span>ID do locatário (diretório)</span>
          <input type="text" name="tenantId" value="${g.tenantId || ''}" placeholder="GUID ou empresa.onmicrosoft.com" />
        </label>
        <label class="field">
          <span>ID do aplicativo (cliente)</span>
          <input type="text" name="clientId" value="${g.clientId || ''}" placeholder="00000000-0000-0000-0000-000000000000" />
        </label>
        <label class="field full">
          <span>Segredo do cliente (valor)</span>
          <input type="password" name="clientSecret" autocomplete="new-password" placeholder="${g.hasClientSecret ? SAVED : ''}" />
        </label>
      </div>
      <details class="help">
        <summary>Como criar o registro do aplicativo</summary>
        <ol>
          <li>No centro de administração do Microsoft Entra (entra.microsoft.com), abra <b>Registros de aplicativo › Novo registro</b> (ex.: "CLEAN", somente esta organização).</li>
          <li>Em <b>Permissões de API › Adicionar › Microsoft Graph › Permissões de aplicativo</b>, inclua <code>Mail.Read</code> e <code>User.Read.All</code> e clique em <b>Conceder consentimento do administrador</b>.</li>
          <li>Em <b>Certificados e segredos › Novo segredo do cliente</b>, copie o <b>Valor</b> (não o ID do segredo).</li>
          <li>Na página <b>Visão geral</b>, copie o ID do aplicativo (cliente) e o ID do diretório (locatário).</li>
          <li>Recomendado: limite o aplicativo às caixas que devem ser analisadas com o RBAC para aplicativos do Exchange Online (veja o LEIA-ME).</li>
        </ol>
      </details>
    </fieldset>

    <fieldset class="full" data-type="gmail">
      <legend>Google Workspace (API do Gmail)</legend>
      <div class="form-grid">
        <div class="field full">
          <label class="field-label" for="sa-file">Chave da conta de serviço (arquivo JSON)</label>
          <input id="sa-file" type="file" accept=".json,application/json" data-sa-file />
          <textarea name="serviceAccountJson" hidden></textarea>
          <small data-sa-info>${gm.clientEmail ? `Conta de serviço atual: ${gm.clientEmail} (chave salva). Envie outro arquivo só para trocar.` : 'O arquivo é baixado no Google Cloud (IAM › Contas de serviço › Chaves).'}</small>
        </div>
        <label class="field full">
          <span>E-mail de um administrador do Google Workspace</span>
          <input type="email" name="adminEmail" value="${gm.adminEmail || ''}" placeholder="admin@empresa.com.br" />
          <small>Usado apenas para listar os usuários quando todas as caixas do domínio são analisadas.</small>
        </label>
      </div>
      <details class="help">
        <summary>Como configurar a conta de serviço</summary>
        <ol>
          <li>No Google Cloud Console, crie (ou escolha) um projeto e ative a <b>Gmail API</b> e a <b>Admin SDK API</b>.</li>
          <li>Em <b>IAM e administrador › Contas de serviço</b>, crie uma conta e, em <b>Chaves › Adicionar chave › JSON</b>, baixe o arquivo.</li>
          <li>No Admin Console (admin.google.com), abra <b>Segurança › Acesso e controle de dados › Controles de API › Delegação em todo o domínio › Adicionar novo</b>.</li>
          <li>Informe o ID do cliente da conta de serviço${gm.clientId ? html` (<code>${gm.clientId}</code>)` : ''} e os escopos:<br /><code>${GMAIL_SCOPES}</code></li>
        </ol>
      </details>
    </fieldset>

    <fieldset class="full" data-type="imap">
      <legend>Servidor IMAP</legend>
      <div class="form-grid">
        <label class="field">
          <span>Servidor</span>
          <input type="text" name="host" value="${im.host || ''}" placeholder="imap.empresa.com.br" />
        </label>
        <div class="field">
          <span class="field-label">Porta e segurança</span>
          <div class="inline">
            <input type="number" name="port" min="1" max="65535" value="${im.port || 993}" aria-label="Porta" />
            <select name="security" aria-label="Segurança">
              ${Object.entries(SECURITY).map(([value, s]) => html`<option value="${value}" ${(im.security || 'tls') === value ? 'selected' : ''}>${s.label}</option>`)}
            </select>
          </div>
        </div>
        <label class="check full">
          <input type="checkbox" name="allowSelfSigned" ${im.allowSelfSigned ? 'checked' : ''} />
          <span>Aceitar certificado não confiável <small class="muted">(servidor interno com certificado próprio)</small></span>
        </label>
        <label class="field full">
          <span>Senha padrão (opcional)</span>
          <input type="password" name="defaultPassword" autocomplete="new-password" placeholder="${im.hasDefaultPassword ? SAVED : 'usada nas caixas sem senha própria'}" />
          <small>Útil com uma conta de serviço que acessa todas as caixas (ex.: Exchange local com login DOMINIO\\servico\\caixa; Dovecot com caixa*mestre).</small>
        </label>
      </div>
    </fieldset>

    <fieldset class="full">
      <legend>Caixas a analisar</legend>
      <div data-scope-box>
        <label class="check"><input type="radio" name="scope" value="all" ${scope === 'all' ? 'checked' : ''} /><span>Todas as caixas <span data-domain></span></span></label>
        <label class="check"><input type="radio" name="scope" value="list" ${scope === 'list' ? 'checked' : ''} /><span>Somente as caixas informadas</span></label>
        <label class="field" data-scope="list">
          <span>Caixas (uma por linha)</span>
          <textarea name="mailboxList" rows="4" placeholder="financeiro@empresa.com.br&#10;rh@empresa.com.br">${list}</textarea>
        </label>
        <label class="field" data-scope="all">
          <span>Ignorar estas caixas (uma por linha; aceita *)</span>
          <textarea name="excludeMailboxes" rows="2" placeholder="noreply@*&#10;teste@empresa.com.br">${(src?.excludeMailboxes || []).join('\n')}</textarea>
        </label>
      </div>
      <div data-imap-box>
        <div class="table-wrap">
          <table class="mailbox-rows">
            <thead><tr><th>E-mail</th><th>Login (se diferente)</th><th>Senha</th><th><span class="sr-only">Remover</span></th></tr></thead>
            <tbody data-rows>${imapRows.map(mailboxRow)}</tbody>
          </table>
        </div>
        <div class="inline page-actions">
          <button type="button" class="btn small" data-action="add-row">${icon('plus')} Adicionar caixa</button>
          <button type="button" class="btn small" data-action="bulk">Adicionar várias…</button>
        </div>
        <div class="field" data-bulk hidden>
          <label class="field-label" for="bulk-list">E-mails (um por linha) — usam a senha padrão</label>
          <textarea id="bulk-list" rows="4"></textarea>
          <div class="inline"><button type="button" class="btn small" data-action="bulk-add">Incluir na lista</button></div>
        </div>
      </div>
    </fieldset>

    <label class="field full">
      <span>Pastas ignoradas (uma por linha)</span>
      <textarea name="excludeFolders" rows="2" placeholder="Pessoal&#10;Caixa de Entrada/Newsletters">${(src?.excludeFolders || []).join('\n')}</textarea>
      <small>Aceita * e ?; ignora também as subpastas. A Lixeira e o Lixo Eletrônico são escolhidos em cada análise.</small>
    </label>
    <label class="field full">
      <span>Descrição (opcional)</span>
      <input type="text" name="description" maxlength="1000" value="${src?.description || ''}" />
    </label>

    <div class="field full">
      <div class="inline"><button type="button" class="btn" data-action="test">${icon('check')} Testar conexão</button></div>
      <div class="test-result" data-test-result aria-live="polite"></div>
    </div>
  </div>`;
}

function readForm(form, existing) {
  const f = new FormData(form);
  const type = f.get('type');
  const body = { name: f.get('name'), type, description: f.get('description'), excludeFolders: f.get('excludeFolders') };
  if (existing) body.id = existing.id;
  if (type === 'graph') body.graph = { tenantId: f.get('tenantId'), clientId: f.get('clientId'), clientSecret: f.get('clientSecret') };
  if (type === 'gmail') body.gmail = { serviceAccountJson: f.get('serviceAccountJson'), adminEmail: f.get('adminEmail') };
  if (type === 'imap') {
    body.imap = {
      host: f.get('host'),
      port: Number(f.get('port')) || undefined,
      security: f.get('security'),
      allowSelfSigned: f.get('allowSelfSigned') === 'on',
      defaultPassword: f.get('defaultPassword'),
    };
    body.scope = 'list';
    body.mailboxes = [...form.querySelectorAll('[data-mailbox-row]')]
      .map((row) => ({
        address: row.querySelector('[name="mbAddress"]').value.trim(),
        login: row.querySelector('[name="mbLogin"]').value.trim(),
        password: row.querySelector('[name="mbPassword"]').value,
      }))
      .filter((m) => m.address);
  } else {
    body.scope = f.get('scope');
    body.mailboxes = f.get('mailboxList');
    body.excludeMailboxes = f.get('excludeMailboxes');
  }
  return body;
}

function showTest(box, result) {
  box.className = `test-result ${result.ok ? 'ok' : 'fail'}`;
  paint(box, html`${result.message}${result.details?.length ? html`<ul class="test-details">${result.details.map((d) => html`<li>${d}</li>`)}</ul>` : ''}`);
}

function wireForm(form, existing) {
  const rows = form.querySelector('[data-rows]');
  const sync = () => {
    const type = form.elements.type.value;
    form.querySelectorAll('[data-type]').forEach((el) => {
      el.hidden = el.dataset.type !== type;
    });
    const imap = type === 'imap';
    form.querySelector('[data-scope-box]').hidden = imap;
    form.querySelector('[data-imap-box]').hidden = !imap;
    const scope = form.elements.scope.value || 'all';
    form.querySelectorAll('[data-scope]').forEach((el) => {
      el.hidden = el.dataset.scope !== scope;
    });
    form.querySelector('[data-domain]').textContent = TYPES[type]?.domain || '';
  };
  form.addEventListener('change', (event) => {
    if (event.target.name === 'security') {
      const port = form.elements.port;
      const defaults = Object.values(SECURITY).map((s) => String(s.port));
      if (!port.value || defaults.includes(port.value)) port.value = SECURITY[event.target.value].port;
    }
    sync();
  });
  form.querySelector('[data-sa-file]').addEventListener('change', async (event) => {
    const file = event.target.files[0];
    const info = form.querySelector('[data-sa-info]');
    if (!file) return;
    const text = await file.text();
    form.elements.serviceAccountJson.value = text;
    try {
      const data = JSON.parse(text);
      info.textContent = data.client_email ? `Conta de serviço: ${data.client_email} (ID do cliente ${data.client_id || '—'}).` : 'O arquivo não parece ser a chave de uma conta de serviço.';
    } catch {
      info.textContent = 'O arquivo escolhido não é um JSON válido.';
    }
  });
  form.addEventListener('click', async (event) => {
    const button = event.target.closest('[data-action]');
    if (!button) return;
    const action = button.dataset.action;
    if (action === 'add-row') {
      rows.insertAdjacentHTML('beforeend', mailboxRow().toString());
      rows.lastElementChild.querySelector('input').focus();
    } else if (action === 'remove-row') {
      button.closest('tr').remove();
      if (!rows.children.length) rows.insertAdjacentHTML('beforeend', mailboxRow().toString());
    } else if (action === 'bulk') {
      const box = form.querySelector('[data-bulk]');
      box.hidden = !box.hidden;
      if (!box.hidden) box.querySelector('textarea').focus();
    } else if (action === 'bulk-add') {
      const area = form.querySelector('#bulk-list');
      const existingAddresses = new Set([...rows.querySelectorAll('[name="mbAddress"]')].map((i) => i.value.trim().toLowerCase()).filter(Boolean));
      const added = area.value
        .split(/[\s,;]+/)
        .map((v) => v.trim())
        .filter((v) => v && !existingAddresses.has(v.toLowerCase()));
      [...rows.querySelectorAll('[data-mailbox-row]')].forEach((row) => {
        if (!row.querySelector('[name="mbAddress"]').value.trim()) row.remove();
      });
      for (const address of new Set(added)) rows.insertAdjacentHTML('beforeend', mailboxRow({ address }).toString());
      area.value = '';
      form.querySelector('[data-bulk]').hidden = true;
      toast(`${plural(added.length, 'caixa incluída', 'caixas incluídas')}.`, 'success');
    } else if (action === 'test') {
      const box = form.querySelector('[data-test-result]');
      box.className = 'test-result';
      box.textContent = 'Testando a conexão… (pode levar alguns segundos)';
      button.disabled = true;
      try {
        showTest(box, await post('/api/mail-sources/test', readForm(form, existing)));
      } catch (err) {
        showTest(box, { ok: false, message: err.message });
      } finally {
        button.disabled = false;
      }
    }
  });
  sync();
}

function scopeText(s) {
  if (s.type !== 'imap' && s.scope === 'all') {
    const except = s.excludeMailboxes?.length ? ` (exceto ${fmtNum(s.excludeMailboxes.length)})` : '';
    return `Todas as caixas ${TYPES[s.type].domain}${except}`;
  }
  return plural(s.mailboxes.length, 'caixa', 'caixas');
}

function credentialText(s) {
  if (s.type === 'graph') return `Aplicativo ${s.graph?.clientId || '—'} · ${s.graph?.hasClientSecret ? 'segredo salvo' : 'sem segredo'}`;
  if (s.type === 'gmail') return `Conta de serviço ${s.gmail?.clientEmail || '—'}`;
  const im = s.imap || {};
  return `${im.host}:${im.port} · ${SECURITY[im.security]?.label || im.security}`;
}

export async function render(root) {
  let sources = [];

  const draw = () =>
    paint(
      root,
      html`<div class="page-head">
          <div>
            <h1>Caixas de e-mail</h1>
            <div class="sub">Conexões com o Microsoft 365, o Google Workspace ou servidores IMAP. Senhas, segredos e chaves ficam gravados cifrados no servidor do CLEAN e não são exibidos novamente.</div>
          </div>
          <div class="actions">
            <button class="btn primary" data-action="new">${icon('plus')} Nova conexão</button>
          </div>
        </div>
        <section class="card">
          ${sources.length === 0
            ? html`<div class="empty">
                <p>Nenhuma conexão de e-mail cadastrada.</p>
                <button class="btn primary" data-action="new">${icon('plus')} Cadastrar a primeira</button>
              </div>`
            : html`<div class="table-wrap">
                <table class="data">
                  <thead><tr><th>Nome</th><th>Tipo</th><th>Caixas</th><th>Credencial</th><th><span class="sr-only">Ações</span></th></tr></thead>
                  <tbody>
                    ${sources.map(
                      (s) => html`<tr>
                        <td><b>${s.name}</b>${s.description ? html`<div class="muted small">${s.description}</div>` : ''}</td>
                        <td class="nowrap">${TYPES[s.type]?.label || s.type}</td>
                        <td class="small">${scopeText(s)}${s.type === 'imap' || s.scope === 'list' ? html`<div class="muted">${s.mailboxes.slice(0, 3).map((m) => m.address).join(', ')}${s.mailboxes.length > 3 ? '…' : ''}</div>` : ''}</td>
                        <td class="small">${credentialText(s)}</td>
                        <td class="actions">
                          <button class="icon-btn" data-action="test" data-id="${s.id}" aria-label="Testar ${s.name}" title="Testar conexão">${icon('check')}</button>
                          <button class="icon-btn" data-action="edit" data-id="${s.id}" aria-label="Editar ${s.name}" title="Editar">${icon('edit')}</button>
                          <button class="icon-btn danger" data-action="delete" data-id="${s.id}" aria-label="Excluir ${s.name}" title="Excluir">${icon('trash')}</button>
                        </td>
                      </tr>`,
                    )}
                  </tbody>
                </table>
              </div>`}
        </section>
        <section class="card">
          <h2>Como funciona</h2>
          <p class="muted small">
            A análise de e-mail percorre todas as pastas de cada caixa (inclusive subpastas e, se escolhido, a Lixeira e o Lixo Eletrônico),
            baixa cada mensagem e procura os termos das listas de referência no assunto, no corpo e nos anexos — Word, Excel, PowerPoint, PDF,
            textos, mensagens encaminhadas e os nomes dos arquivos dentro de .zip. O acesso é somente leitura: nenhuma mensagem é alterada ou marcada como lida.
          </p>
        </section>`,
    );

  const refresh = async () => {
    sources = await get('/api/mail-sources');
    draw();
  };

  const edit = async (source) => {
    const saved = await openDialog({
      title: source ? 'Editar conexão de e-mail' : 'Nova conexão de e-mail',
      body: sourceForm(source),
      wide: true,
      onOpen: (form) => wireForm(form, source),
      onSubmit: (form) => (source ? put(`/api/mail-sources/${source.id}`, readForm(form, source)) : post('/api/mail-sources', readForm(form, null))),
    });
    if (saved) {
      toast(source ? 'Conexão atualizada.' : 'Conexão cadastrada. Use "Testar conexão" para conferir o acesso.', 'success');
      await refresh();
    }
  };

  const test = async (source, button) => {
    button.disabled = true;
    toast(`Testando "${source.name}"…`);
    try {
      const result = await post('/api/mail-sources/test', { ...source, id: source.id, mailboxes: source.type === 'imap' ? source.mailboxes : source.mailboxes.map((m) => m.address) });
      await openDialog({
        title: `Teste: ${source.name}`,
        body: html`<div class="test-result ${result.ok ? 'ok' : 'fail'}">${result.message}${result.details?.length ? html`<ul class="test-details">${result.details.map((d) => html`<li>${d}</li>`)}</ul>` : ''}</div>`,
        submitLabel: 'Fechar',
        cancelLabel: null,
      });
    } catch (err) {
      toast(err.message, 'error');
    } finally {
      button.disabled = false;
    }
  };

  const onClick = async (event) => {
    const button = event.target.closest('[data-action]');
    if (!button || !root.contains(button)) return;
    const source = sources.find((s) => s.id === button.dataset.id);
    const action = button.dataset.action;
    if (action === 'new') return edit(null);
    if (action === 'edit') return edit(source);
    if (action === 'test') return test(source, button);
    if (action === 'delete') {
      const ok = await confirmDialog(`Excluir a conexão "${source.name}"? As senhas e chaves salvas serão apagadas. As análises já feitas continuam disponíveis.`, { confirmLabel: 'Excluir' });
      if (!ok) return;
      try {
        await del(`/api/mail-sources/${source.id}`);
        toast('Conexão excluída.', 'success');
        await refresh();
      } catch (err) {
        toast(err.message, 'error');
      }
    }
  };

  await refresh();
  root.addEventListener('click', onClick);
  return () => root.removeEventListener('click', onClick);
}
