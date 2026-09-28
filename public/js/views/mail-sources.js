// Caixas de e-mail: conexões com o Microsoft 365, o Google Workspace ou servidores IMAP. Na Microsoft,
// a autenticação é OAuth 2.0: aplicativo com segredo do cliente, aplicativo com certificado ou conta
// Microsoft conectada (entrada pelo código de dispositivo) — também no IMAP do Exchange Online.
import { get, post, put, del } from '../api.js';
import { html, render as paint, icon, openDialog, confirmDialog, toast, fmtNum, fmtDate, fmtDateTime, plural, copyText } from '../ui.js';

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
// Formas de autenticação na Microsoft (OAuth 2.0 no Microsoft Entra ID); detail: no Microsoft 365
// (Microsoft Graph) e no IMAP.
const MS_AUTH = {
  secret: {
    label: 'Aplicativo com segredo',
    detail: { graph: 'Permissões de aplicativo, concedidas pelo administrador: todas as caixas do locatário.', imap: 'Permissão IMAP.AccessAsApp e Acesso Total às caixas no Exchange Online.' },
  },
  certificate: {
    label: 'Aplicativo com certificado',
    detail: { graph: 'O mesmo, com um certificado no lugar do segredo (recomendado pela Microsoft).', imap: 'O mesmo do segredo, com um certificado (recomendado pela Microsoft).' },
  },
  delegated: {
    label: 'Conta Microsoft conectada',
    detail: { graph: 'Entrar com a conta (trabalho, escola ou pessoal): a caixa dela e as compartilhadas com ela.', imap: 'Entrar com a conta: a caixa dela e as caixas a que ela tem Acesso Total.' },
  },
};
// Servidores IMAP da Microsoft: não aceitam mais senha (só OAuth).
const MICROSOFT_IMAP = new Set(['outlook.office365.com', 'outlook.office.com', 'imap-mail.outlook.com']);
const isMicrosoftHost = (host) => MICROSOFT_IMAP.has(String(host || '').trim().toLowerCase());
const GMAIL_SCOPES = 'https://www.googleapis.com/auth/gmail.readonly,https://www.googleapis.com/auth/admin.directory.user.readonly';
const SAVED = '•••••• salvo — deixe em branco para manter';

function mailboxRow(m = {}) {
  return html`<tr data-mailbox-row>
    <td><input type="email" name="mbAddress" value="${m.address || ''}" placeholder="nome@empresa.com.br" aria-label="E-mail da caixa" /></td>
    <td><input type="text" name="mbLogin" value="${m.login || ''}" placeholder="igual ao e-mail" aria-label="Login da caixa" /></td>
    <td data-password-cell><input type="password" name="mbPassword" autocomplete="new-password" placeholder="${m.hasPassword ? 'salva' : 'usa a senha padrão'}" data-saved="${m.hasPassword ? '1' : ''}" aria-label="Senha da caixa" /></td>
    <td><button type="button" class="icon-btn danger" data-action="remove-row" aria-label="Remover caixa" title="Remover">${icon('x')}</button></td>
  </tr>`;
}

/** Dias até o certificado vencer (negativo: vencido). */
const daysLeft = (iso) => Math.floor((Date.parse(iso) - Date.now()) / 86400000);

function expiryChip(certificate) {
  if (!certificate?.notAfter) return '';
  const days = daysLeft(certificate.notAfter);
  if (days < 0) return html` <span class="chip danger">vencido</span>`;
  if (days < 30) return html` <span class="chip danger">vence em ${plural(days, 'dia', 'dias')}</span>`;
  return '';
}

function certificateInfo(c, { pending = false } = {}) {
  if (!c) return html`<span class="muted">Nenhum certificado. Gere um aqui (ou use um existente) e envie o arquivo .cer ao registro do aplicativo.</span>`;
  return html`<b>${c.subject || 'Certificado'}</b> · válido até ${fmtDate(c.notAfter)}${expiryChip(c)}
    <br /><span class="small muted">Impressão digital (SHA-1): <code>${c.thumbprint}</code>${pending ? ' · ainda não salvo: baixe o arquivo, envie-o ao registro do aplicativo e salve a conexão.' : ''}</span>`;
}

function accountInfo(a, { pending = false, stale = '' } = {}) {
  if (!a) return html`<span class="muted">Nenhuma conta conectada.</span>`;
  const note = stale
    ? html`<span class="small warn-text">${stale}</span>`
    : html`<span class="small muted">${pending ? 'Conectada agora: salve a conexão para guardar a autorização.' : `Conectada em ${fmtDate(a.connectedAt)}. O acesso é renovado sozinho a cada uso (a autorização vence após 90 dias sem uso).`}</span>`;
  return html`${icon('user')} <b>${a.name || a.username || a.address}</b>${a.address || a.username ? html` — ${a.address || a.username}` : ''}${a.canDelete === false ? html` <span class="chip">somente leitura</span>` : ''}<br />${note}`;
}

function sourceForm(src) {
  const type = src?.type || 'graph';
  const g = src?.graph || {};
  const gm = src?.gmail || {};
  const im = src?.imap || {};
  const imapAuth = im.auth === 'oauth' ? 'oauth' : 'password';
  const msAuth = g.auth || (type === 'imap' ? 'delegated' : 'secret');
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
        <label class="field full">
          <span>Autenticação</span>
          <select name="imapAuth">
            <option value="password" ${imapAuth === 'password' ? 'selected' : ''}>Usuário e senha</option>
            <option value="oauth" ${imapAuth === 'oauth' ? 'selected' : ''}>OAuth 2.0 da Microsoft (Exchange Online e Outlook.com)</option>
          </select>
          <small data-imap-oauth-hint>Servidor da Microsoft: <code>outlook.office365.com</code>, porta 993, SSL/TLS. O login usa o token da Microsoft (XOAUTH2), sem senha — escolha abaixo como o CLEAN se autentica.</small>
        </label>
        <p class="alert full" data-ms-password-warning hidden>
          ${icon('alert')}<span>A Microsoft não aceita mais senha no IMAP do Exchange Online e do Outlook.com. Em <b>Autenticação</b>, escolha <b>OAuth 2.0 da Microsoft</b> — ou use o tipo de conexão <b>Microsoft 365</b>.</span>
        </p>
        <label class="check full" data-imap-password>
          <input type="checkbox" name="allowSelfSigned" ${im.allowSelfSigned ? 'checked' : ''} />
          <span>Aceitar certificado não confiável <small class="muted">(servidor interno com certificado próprio)</small></span>
        </label>
        <label class="field full" data-imap-password>
          <span>Senha padrão (opcional)</span>
          <input type="password" name="defaultPassword" autocomplete="new-password" placeholder="${im.hasDefaultPassword ? SAVED : 'usada nas caixas sem senha própria'}" />
          <small>Útil com uma conta de serviço que acessa todas as caixas (ex.: Exchange local com login DOMINIO\\servico\\caixa; Dovecot com caixa*mestre).</small>
        </label>
        <p class="alert full" data-reenter hidden>${icon('alert')}<span>Servidor, porta ou segurança alterados: por proteção, as senhas salvas não serão usadas. Informe-as novamente.</span></p>
      </div>
    </fieldset>

    <fieldset class="full" data-ms>
      <legend>Autenticação na Microsoft (OAuth 2.0)</legend>
      <div class="type-choice">
        ${Object.entries(MS_AUTH).map(
          ([value, a]) => html`<label>
            <input type="radio" name="msAuth" value="${value}" ${msAuth === value ? 'checked' : ''} />
            <span><b>${a.label}</b><small data-ms-for="graph">${a.detail.graph}</small><small data-ms-for="imap">${a.detail.imap}</small></span>
          </label>`,
        )}
      </div>
      <div class="form-grid ms-fields">
        <label class="field">
          <span>ID do locatário (diretório)</span>
          <input type="text" name="tenantId" value="${g.tenantId || ''}" placeholder="GUID ou empresa.onmicrosoft.com" />
          <small data-ms-show="delegated">Com a conta conectada, também <code>organizations</code> (contas de trabalho ou escola) ou <code>consumers</code> (contas pessoais: Outlook.com, Hotmail).</small>
        </label>
        <label class="field">
          <span>ID do aplicativo (cliente)</span>
          <input type="text" name="clientId" value="${g.clientId || ''}" placeholder="00000000-0000-0000-0000-000000000000" />
        </label>
        <label class="field full" data-ms-show="secret">
          <span>Segredo do cliente (valor)</span>
          <input type="password" name="clientSecret" autocomplete="new-password" placeholder="${g.hasClientSecret ? SAVED : ''}" />
        </label>
        <div class="field full" data-ms-show="certificate">
          <span class="field-label">Certificado do aplicativo</span>
          <div class="ms-status" data-cert-info>${certificateInfo(g.certificate)}</div>
          <div class="inline">
            <button type="button" class="btn small" data-action="cert-generate">${icon('plus')} <span data-cert-generate-label>${g.certificate ? 'Gerar novo certificado' : 'Gerar certificado'}</span></button>
            <button type="button" class="btn small" data-action="cert-download" ${g.certificate ? '' : 'hidden'}>${icon('download')} Baixar certificado (.cer)</button>
          </div>
          <div class="alert" data-cert-confirm hidden>
            ${icon('alert')}
            <div>
              Gerar um novo certificado? Depois de salvar a conexão, o CLEAN passa a usar o novo: envie o novo arquivo .cer ao registro do aplicativo
              antes (o certificado atual pode continuar lá até você removê-lo).
              <div class="inline"><button type="button" class="btn small primary" data-action="cert-generate-confirm">Gerar novo certificado</button><button type="button" class="btn small" data-action="cert-generate-cancel">Cancelar</button></div>
            </div>
          </div>
          <details class="help">
            <summary>Usar um certificado existente (arquivo PEM)</summary>
            <p>Escolha um arquivo PEM com o certificado e a chave privada sem senha. Para um .pfx, converta antes: <code>openssl pkcs12 -in certificado.pfx -out certificado.pem -nodes</code>.</p>
            <input type="file" accept=".pem,.crt,.cer,.key,.txt" data-cert-file aria-label="Arquivo PEM com o certificado e a chave privada" />
          </details>
          <input type="hidden" name="certificateId" value="" />
          <small>Envie o arquivo .cer ao registro do aplicativo, em <b>Certificados e segredos › Certificados › Carregar certificado</b>. A chave privada não sai do CLEAN: fica gravada cifrada no servidor.</small>
        </div>
        <div class="field full" data-ms-show="delegated">
          <span class="field-label">Conta conectada</span>
          <div class="ms-status" data-account-info>${accountInfo(g.account)}</div>
          <div class="inline">
            <button type="button" class="btn small primary" data-action="connect">${icon('user')} <span data-connect-label>${g.account ? 'Conectar outra conta' : 'Conectar conta'}</span></button>
          </div>
          <p class="alert full" data-read-only-warning hidden>${icon('alert')}<span>A exclusão está permitida, mas a conta conectada autorizou só a leitura das mensagens: clique em <b>Conectar conta</b> de novo (com "Permitir excluir" marcado) para autorizar a exclusão.</span></p>
          <div class="device-box" data-device hidden></div>
          <input type="hidden" name="signIn" value="" />
        </div>
      </div>
      <details class="help">
        <summary>Como configurar o registro do aplicativo</summary>
        <ol data-ms-show="secret certificate">
          <li>No centro de administração do Microsoft Entra (entra.microsoft.com), abra <b>Registros de aplicativo › Novo registro</b> (ex.: "CLEAN", somente esta organização).</li>
          <li data-ms-for="graph">Em <b>Permissões de API › Adicionar › Microsoft Graph › Permissões de aplicativo</b>, inclua <code>Mail.Read</code> e <code>User.Read.All</code> e clique em <b>Conceder consentimento do administrador</b>.</li>
          <li data-ms-for="imap">Em <b>Permissões de API › Adicionar › APIs que minha organização usa › Office 365 Exchange Online › Permissões de aplicativo</b>, inclua <code>IMAP.AccessAsApp</code> e conceda o consentimento do administrador. No Exchange Online (PowerShell), registre o aplicativo (<code>New-ServicePrincipal</code>) e dê a ele Acesso Total a cada caixa (<code>Add-MailboxPermission</code>) — veja o LEIA-ME.</li>
          <li data-ms-show="secret">Em <b>Certificados e segredos › Novo segredo do cliente</b>, copie o <b>Valor</b> (não o ID do segredo).</li>
          <li data-ms-show="certificate">Clique em <b>Gerar certificado</b>, baixe o arquivo .cer e envie-o em <b>Certificados e segredos › Certificados › Carregar certificado</b>.</li>
          <li>Na página <b>Visão geral</b>, copie o ID do aplicativo (cliente) e o ID do diretório (locatário).</li>
          <li data-ms-for="graph">Recomendado: limite o aplicativo às caixas que devem ser analisadas com o RBAC para aplicativos do Exchange Online (veja o LEIA-ME).</li>
        </ol>
        <ol data-ms-show="delegated">
          <li>No centro de administração do Microsoft Entra (entra.microsoft.com), abra <b>Registros de aplicativo › Novo registro</b> (ex.: "CLEAN"). Para contas pessoais (Outlook.com), escolha "Contas em qualquer diretório organizacional e contas Microsoft pessoais".</li>
          <li>Em <b>Autenticação</b>, ative <b>Permitir fluxos de cliente público</b> (entrada pelo código de dispositivo) e salve.</li>
          <li data-ms-for="graph">Em <b>Permissões de API › Adicionar › Microsoft Graph › Permissões delegadas</b>, inclua <code>User.Read</code>, <code>Mail.Read</code> e <code>Mail.Read.Shared</code> (para excluir: <code>Mail.ReadWrite</code> e <code>Mail.ReadWrite.Shared</code>).</li>
          <li data-ms-for="imap">Em <b>Permissões de API › Adicionar › Microsoft Graph › Permissões delegadas</b>, inclua <code>IMAP.AccessAsUser.All</code> e <code>offline_access</code>. O IMAP precisa estar habilitado nas caixas.</li>
          <li>Copie o ID do aplicativo (cliente) e o ID do diretório (locatário), clique em <b>Conectar conta</b> e entre com a conta. A própria pessoa autoriza as permissões ao entrar (se a organização não permitir, o administrador concede o consentimento no registro do aplicativo).</li>
          <li>Caixas compartilhadas ou de outras pessoas: a conta conectada precisa ter <b>Acesso Total</b> a elas no Exchange Online.</li>
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

    <fieldset class="full">
      <legend>Caixas a analisar</legend>
      <div data-scope-box>
        <label class="check" data-scope-all-choice><input type="radio" name="scope" value="all" ${scope === 'all' ? 'checked' : ''} /><span>Todas as caixas <span data-domain></span></span></label>
        <label class="check"><input type="radio" name="scope" value="list" ${scope === 'list' ? 'checked' : ''} /><span>Somente as caixas informadas</span></label>
        <label class="field" data-scope="list">
          <span>Caixas (uma por linha)</span>
          <textarea name="mailboxList" rows="4" placeholder="financeiro@empresa.com.br&#10;rh@empresa.com.br">${list}</textarea>
          <small data-delegated-list-hint>Com a conta conectada: a caixa da própria conta e as caixas compartilhadas (ou de outras pessoas) às quais ela tem acesso. Em branco, só a caixa da conta.</small>
        </label>
        <label class="field" data-scope="all">
          <span>Ignorar estas caixas (uma por linha; aceita *)</span>
          <textarea name="excludeMailboxes" rows="2" placeholder="noreply@*&#10;teste@empresa.com.br">${(src?.excludeMailboxes || []).join('\n')}</textarea>
        </label>
      </div>
      <div data-imap-box>
        <div class="table-wrap">
          <table class="mailbox-rows">
            <thead><tr><th>E-mail</th><th>Login (se diferente)</th><th data-password-cell>Senha</th><th><span class="sr-only">Remover</span></th></tr></thead>
            <tbody data-rows>${imapRows.map(mailboxRow)}</tbody>
          </table>
        </div>
        <small class="hint" data-imap-oauth-rows>Com OAuth, o login de cada caixa é o próprio e-mail: a caixa da conta conectada e as caixas às quais ela (ou o aplicativo) tem Acesso Total. Em branco, só a caixa da conta conectada.</small>
        <div class="inline page-actions">
          <button type="button" class="btn small" data-action="add-row">${icon('plus')} Adicionar caixa</button>
          <button type="button" class="btn small" data-action="bulk">Adicionar várias…</button>
        </div>
        <div class="field" data-bulk hidden>
          <label class="field-label" for="bulk-list" data-bulk-label>E-mails (um por linha) — usam a senha padrão</label>
          <textarea id="bulk-list" rows="4"></textarea>
          <div class="inline"><button type="button" class="btn small" data-action="bulk-add">Incluir na lista</button></div>
        </div>
      </div>
    </fieldset>

    <fieldset class="full">
      <legend>Exclusão das mensagens encontradas</legend>
      <label class="check">
        <input type="checkbox" name="allowDelete" ${src?.allowDelete ? 'checked' : ''} />
        <span><b>Permitir excluir as mensagens em que os termos forem encontrados</b><br /><small class="muted">Na análise ("analisar e excluir") ou item a item pelo relatório.</small></span>
      </label>
      <label class="field" data-delete-mode ${src?.allowDelete ? '' : 'hidden'}>
        <span>Como excluir</span>
        <select name="deleteMode">
          <option value="permanent" ${src?.deleteMode !== 'trash' ? 'selected' : ''}>Excluir definitivamente</option>
          <option value="trash" ${src?.deleteMode === 'trash' ? 'selected' : ''}>Mover para a Lixeira (no Outlook, Itens Excluídos)</option>
        </select>
      </label>
      <p class="hint" data-delete-help ${src?.allowDelete ? '' : 'hidden'}>
        Permissões necessárias — <b>Microsoft 365:</b> com um aplicativo, Mail.ReadWrite (tipo Aplicativo) no lugar de Mail.Read; com a
        conta conectada, conecte a conta com esta opção marcada (Mail.ReadWrite).
        <b>Google Workspace:</b> autorize também o escopo <code>https://mail.google.com/</code> (definitiva) ou
        <code>https://www.googleapis.com/auth/gmail.modify</code> (lixeira) na delegação em todo o domínio.
        <b>IMAP:</b> a conta precisa poder alterar a caixa, e o servidor precisa oferecer UIDPLUS para excluir
        definitivamente (ou MOVE para mover para a Lixeira) — sem isso, a exclusão é recusada para não apagar outras
        mensagens. Retenções e bloqueios legais do provedor continuam valendo.
      </p>
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

/** A conexão usa as credenciais da Microsoft: Microsoft 365, ou IMAP com OAuth da Microsoft. */
const usesMicrosoft = (form) => form.elements.type.value === 'graph' || (form.elements.type.value === 'imap' && form.elements.imapAuth.value === 'oauth');

function readForm(form, existing) {
  const f = new FormData(form);
  const type = f.get('type');
  const body = {
    name: f.get('name'),
    type,
    description: f.get('description'),
    excludeFolders: f.get('excludeFolders'),
    allowDelete: f.get('allowDelete') === 'on',
    deleteMode: f.get('deleteMode') || 'permanent',
  };
  if (existing) body.id = existing.id;
  if (usesMicrosoft(form)) {
    body.graph = {
      tenantId: f.get('tenantId'),
      clientId: f.get('clientId'),
      auth: f.get('msAuth') || 'secret',
      clientSecret: f.get('clientSecret'),
      certificateId: f.get('certificateId') || '',
      signIn: f.get('signIn') || '',
    };
  }
  if (type === 'gmail') body.gmail = { serviceAccountJson: f.get('serviceAccountJson'), adminEmail: f.get('adminEmail') };
  if (type === 'imap') {
    const oauth = f.get('imapAuth') === 'oauth';
    body.imap = {
      host: f.get('host'),
      port: Number(f.get('port')) || undefined,
      security: f.get('security'),
      auth: oauth ? 'oauth' : 'password',
      allowSelfSigned: !oauth && f.get('allowSelfSigned') === 'on',
      defaultPassword: oauth ? '' : f.get('defaultPassword'),
    };
    body.scope = 'list';
    body.mailboxes = [...form.querySelectorAll('[data-mailbox-row]')]
      .map((row) => ({
        address: row.querySelector('[name="mbAddress"]').value.trim(),
        login: row.querySelector('[name="mbLogin"]').value.trim(),
        password: oauth ? '' : row.querySelector('[name="mbPassword"]').value,
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

/** Mesma regra do servidor: as senhas salvas valem só para o mesmo servidor, porta e segurança. */
function sameEndpoint(saved, form) {
  const security = form.elements.security.value;
  const port = Number(form.elements.port.value) || SECURITY[security]?.port;
  return (
    saved.host === form.elements.host.value.trim().toLowerCase() &&
    Number(saved.port) === port &&
    saved.security === security &&
    (Boolean(saved.allowSelfSigned) || !form.elements.allowSelfSigned.checked)
  );
}

/** Endereço da página de entrada da Microsoft (só https; o padrão, se vier outra coisa). */
function safeUri(value) {
  try {
    const u = new URL(value);
    if (u.protocol === 'https:') return u.href;
  } catch {
    // endereço inválido
  }
  return 'https://microsoft.com/devicelogin';
}

/** Baixa o certificado (só a parte pública) para enviar ao registro do aplicativo. */
function downloadCertificate(certificate, name) {
  const slug = String(name || 'conexao').normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^\w-]+/g, '-').replace(/^-+|-+$/g, '').slice(0, 60) || 'conexao';
  const url = URL.createObjectURL(new Blob([certificate.pem], { type: 'application/x-x509-ca-cert' }));
  const a = document.createElement('a');
  a.href = url;
  a.download = `CLEAN-${slug}.cer`;
  document.body.append(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 10000);
}

function wireForm(form, existing) {
  const rows = form.querySelector('[data-rows]');
  const saved = existing?.type === 'imap' && existing.imap?.auth !== 'oauth' ? existing.imap : null;
  const savedGraph = existing?.graph || null;
  // Conta conectada e certificado mostrados no formulário: os salvos ou os novos (ainda não salvos).
  const state = {
    account: savedGraph?.account || null,
    accountPending: false,
    // Para que a conta vale: tipo de conexão, locatário e aplicativo em que ela entrou.
    accountFor: savedGraph?.account ? { purpose: existing.type === 'imap' ? 'imap' : 'graph', tenantId: savedGraph.tenantId, clientId: savedGraph.clientId } : null,
    certificate: savedGraph?.certificate || null,
    flow: null, // entrada em andamento: { id, timer, stopped }
  };
  const purpose = () => (form.elements.type.value === 'imap' ? 'imap' : 'graph');
  const lower = (v) => String(v || '').trim().toLowerCase();
  // Motivo para a conta mostrada não valer mais (tipo, locatário ou aplicativo trocados), ou ''.
  const accountStale = () => {
    const a = state.accountFor;
    if (!state.account || !a) return '';
    if (a.purpose !== purpose()) return 'A conta foi conectada para outro tipo de conexão: clique em "Conectar conta" de novo.';
    if (lower(a.tenantId) !== lower(form.elements.tenantId.value) || lower(a.clientId) !== lower(form.elements.clientId.value)) {
      return 'O locatário ou o aplicativo foi alterado: clique em "Conectar conta" de novo.';
    }
    return '';
  };
  const drawAccount = () => {
    paint(form.querySelector('[data-account-info]'), accountInfo(state.account, { pending: state.accountPending, stale: accountStale() }));
    form.querySelector('[data-connect-label]').textContent = state.account ? 'Conectar outra conta' : 'Conectar conta';
  };
  const drawCertificate = (pending = false) => {
    paint(form.querySelector('[data-cert-info]'), certificateInfo(state.certificate, { pending }));
    form.querySelector('[data-action="cert-download"]').hidden = !state.certificate;
    form.querySelector('[data-cert-generate-label]').textContent = state.certificate ? 'Gerar novo certificado' : 'Gerar certificado';
  };

  // Com outro servidor, porta ou segurança, as senhas salvas deixam de valer: os avisos acompanham.
  const syncSaved = () => {
    if (!saved) return;
    const same = form.elements.type.value === 'imap' && form.elements.imapAuth.value !== 'oauth' && sameEndpoint(saved, form);
    const anySaved = saved.hasDefaultPassword || rows.querySelector('[data-saved="1"]');
    form.querySelector('[data-reenter]').hidden = same || !anySaved || form.elements.imapAuth.value === 'oauth';
    if (saved.hasDefaultPassword) form.elements.defaultPassword.placeholder = same ? SAVED : 'informe novamente';
    rows.querySelectorAll('[name="mbPassword"][data-saved="1"]').forEach((input) => {
      input.placeholder = same ? 'salva' : 'informe novamente';
    });
  };
  const sync = () => {
    const type = form.elements.type.value;
    form.querySelectorAll('[data-type]').forEach((el) => {
      el.hidden = el.dataset.type !== type;
    });
    const imap = type === 'imap';
    const oauth = imap && form.elements.imapAuth.value === 'oauth';
    const microsoft = usesMicrosoft(form);
    const auth = form.elements.msAuth.value || 'secret';
    const delegated = microsoft && auth === 'delegated';
    form.querySelector('[data-ms]').hidden = !microsoft;
    form.querySelectorAll('[data-ms-show]').forEach((el) => {
      el.hidden = !el.dataset.msShow.split(' ').includes(auth);
    });
    form.querySelectorAll('[data-ms-for]').forEach((el) => {
      el.hidden = el.dataset.msFor !== purpose();
    });
    // IMAP: com OAuth, sem senhas; com senha num servidor da Microsoft, o aviso.
    form.querySelector('[data-imap-oauth-hint]').hidden = !oauth;
    form.querySelectorAll('[data-imap-password]').forEach((el) => {
      el.hidden = oauth;
    });
    form.querySelectorAll('[data-password-cell]').forEach((el) => {
      el.hidden = oauth;
    });
    form.querySelector('[data-imap-oauth-rows]').hidden = !oauth;
    form.querySelector('[data-bulk-label]').textContent = oauth ? 'E-mails (um por linha)' : 'E-mails (um por linha) — usam a senha padrão';
    form.querySelector('[data-ms-password-warning]').hidden = !(imap && !oauth && isMicrosoftHost(form.elements.host.value));
    form.querySelector('[data-scope-box]').hidden = imap;
    form.querySelector('[data-imap-box]').hidden = !imap;
    // Conta conectada no Microsoft 365: as caixas são as da lista (a da conta e as compartilhadas).
    const allChoice = form.querySelector('[data-scope-all-choice]');
    allChoice.hidden = delegated;
    if (delegated && form.elements.scope.value !== 'list') form.querySelector('[name="scope"][value="list"]').checked = true;
    form.querySelector('[data-delegated-list-hint]').hidden = !delegated;
    const scope = form.elements.scope.value || 'all';
    form.querySelectorAll('[data-scope]').forEach((el) => {
      el.hidden = el.dataset.scope !== scope;
    });
    form.querySelector('[data-domain]').textContent = TYPES[type]?.domain || '';
    const deletion = form.elements.allowDelete.checked;
    form.querySelector('[data-delete-mode]').hidden = !deletion;
    form.querySelector('[data-delete-help]').hidden = !deletion;
    // Exclusão com a conta conectada (Microsoft Graph): a entrada precisa ter autorizado a escrita.
    form.querySelector('[data-read-only-warning]').hidden = !(delegated && type === 'graph' && deletion && state.account && state.account.canDelete === false && !accountStale());
    drawAccount();
  };

  // ---------- Conta conectada: entrada pelo código de dispositivo ----------
  const deviceBox = form.querySelector('[data-device]');
  const stopFlow = ({ cancel = false } = {}) => {
    const flow = state.flow;
    if (!flow) return;
    flow.stopped = true;
    clearTimeout(flow.timer);
    state.flow = null;
    if (cancel) del(`/api/mail-sources/oauth/device/${flow.id}`).catch(() => {});
  };
  const connected = (flowId, account) => {
    state.account = account;
    state.accountPending = true;
    state.accountFor = { purpose: purpose(), tenantId: form.elements.tenantId.value, clientId: form.elements.clientId.value };
    form.elements.signIn.value = flowId;
    deviceBox.hidden = true;
    // Sem caixas informadas, a lista começa com a caixa da conta.
    const address = account.address || account.username;
    if (address && form.elements.type.value === 'imap') {
      const filled = [...rows.querySelectorAll('[name="mbAddress"]')].some((i) => i.value.trim());
      if (!filled) rows.querySelector('[name="mbAddress"]').value = address;
    } else if (address && !form.elements.mailboxList.value.trim()) {
      form.elements.mailboxList.value = address;
    }
    sync();
    toast(`Conta ${address || ''} conectada. Salve a conexão para guardar a autorização.`, 'success');
  };
  const poll = async (flow) => {
    if (flow.stopped) return;
    if (!form.isConnected) return stopFlow({ cancel: true });
    let result;
    try {
      result = await get(`/api/mail-sources/oauth/device/${flow.id}`);
    } catch (err) {
      result = { status: 'retry', error: err.message };
    }
    if (flow.stopped) return;
    if (result.status === 'connected') {
      stopFlow();
      connected(flow.id, result.account);
      return;
    }
    if (result.status === 'failed') {
      stopFlow();
      paint(deviceBox, html`<div class="alert error">${icon('alert')}<div>${result.error || 'A entrada não foi concluída.'}</div></div>`);
      return;
    }
    // Pendente (ou falha momentânea ao consultar): consulta de novo em alguns segundos.
    flow.timer = setTimeout(() => poll(flow), 3000);
  };
  const startSignIn = async (button) => {
    stopFlow({ cancel: true });
    deviceBox.hidden = false;
    paint(deviceBox, html`<p class="muted">Pedindo um código de entrada à Microsoft…</p>`);
    button.disabled = true;
    let started;
    try {
      started = await post('/api/mail-sources/oauth/device', {
        type: form.elements.type.value,
        graph: { tenantId: form.elements.tenantId.value, clientId: form.elements.clientId.value },
        allowDelete: form.elements.allowDelete.checked,
      });
    } catch (err) {
      paint(deviceBox, html`<div class="alert error">${icon('alert')}<div>${err.message}</div></div>`);
      return;
    } finally {
      button.disabled = false;
    }
    const uri = safeUri(started.verificationUri);
    paint(
      deviceBox,
      html`<ol>
          <li>Abra <a href="${uri}" target="_blank" rel="noopener noreferrer">${uri.replace(/^https:\/\//, '').replace(/\/$/, '')}</a> (neste computador ou no celular).</li>
          <li>Digite o código <code class="device-code">${started.userCode}</code> <button type="button" class="btn small" data-action="copy-code" data-code="${started.userCode}">${icon('copy')} Copiar</button></li>
          <li>Entre com a conta cujas caixas serão analisadas e aceite as permissões pedidas.</li>
        </ol>
        <p class="small muted" aria-live="polite">Aguardando a entrada na página da Microsoft… O código vale até ${fmtDateTime(started.expiresAt)}.</p>
        <button type="button" class="btn small" data-action="connect-cancel">Cancelar</button>`,
    );
    const flow = { id: started.flowId, timer: null, stopped: false };
    state.flow = flow;
    flow.timer = setTimeout(() => poll(flow), Math.max(2, Number(started.interval) || 5) * 1000);
  };
  // Ao fechar o formulário, a entrada em andamento (ou não salva) é descartada no servidor.
  form.closest('dialog')?.addEventListener(
    'close',
    () => {
      stopFlow({ cancel: true });
      if (form.elements.signIn.value) del(`/api/mail-sources/oauth/device/${form.elements.signIn.value}`).catch(() => {});
    },
    { once: true },
  );

  // ---------- Certificado: gerado no servidor ou importado (PEM) ----------
  const useCertificate = (result) => {
    state.certificate = result.certificate;
    form.elements.certificateId.value = result.certificateId;
    drawCertificate(true);
  };

  form.addEventListener('change', (event) => {
    const name = event.target.name;
    if (name === 'security') {
      const port = form.elements.port;
      const defaults = Object.values(SECURITY).map((s) => String(s.port));
      if (!port.value || defaults.includes(port.value)) port.value = SECURITY[event.target.value].port;
    }
    if (name === 'imapAuth' && event.target.value === 'oauth') {
      // Servidor da Microsoft, porta 993 e SSL/TLS: o que o login OAuth exige.
      if (!form.elements.host.value.trim() || !isMicrosoftHost(form.elements.host.value)) form.elements.host.value = 'outlook.office365.com';
      form.elements.security.value = 'tls';
      form.elements.port.value = 993;
      form.elements.allowSelfSigned.checked = false;
      // No IMAP, o comum é entrar com a conta (o aplicativo exige configurar o Exchange Online).
      if (!savedGraph) form.querySelector('[name="msAuth"][value="delegated"]').checked = true;
    }
    if (name === 'type' || name === 'imapAuth' || name === 'msAuth') stopFlow({ cancel: true });
    sync();
    syncSaved();
  });
  form.addEventListener('input', (event) => {
    const name = event.target.name;
    if (['host', 'port'].includes(name)) {
      syncSaved();
      sync();
    }
    if (name === 'tenantId' || name === 'clientId') {
      // A entrada vale para o locatário e o aplicativo em que foi feita.
      stopFlow({ cancel: true });
      deviceBox.hidden = true;
      sync();
    }
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
  form.querySelector('[data-cert-file]').addEventListener('change', async (event) => {
    const input = event.target;
    const file = input.files[0];
    if (!file) return;
    try {
      if (file.size > 200000) throw new Error('O arquivo é grande demais para um certificado.');
      useCertificate(await post('/api/mail-sources/certificate', { pem: await file.text() }));
      toast('Certificado lido. Envie o mesmo certificado ao registro do aplicativo (se ainda não estiver lá) e salve a conexão.', 'success');
    } catch (err) {
      toast(err.message, 'error');
    } finally {
      input.value = '';
    }
  });
  form.addEventListener('click', async (event) => {
    const button = event.target.closest('[data-action]');
    if (!button) return;
    const action = button.dataset.action;
    if (action === 'add-row') {
      rows.insertAdjacentHTML('beforeend', mailboxRow().toString());
      sync();
      rows.lastElementChild.querySelector('input').focus();
    } else if (action === 'remove-row') {
      button.closest('tr').remove();
      if (!rows.children.length) rows.insertAdjacentHTML('beforeend', mailboxRow().toString());
      sync();
    } else if (action === 'bulk') {
      const box = form.querySelector('[data-bulk]');
      box.hidden = !box.hidden;
      if (!box.hidden) box.querySelector('textarea').focus();
    } else if (action === 'bulk-add') {
      const area = form.querySelector('#bulk-list');
      const known = new Set([...rows.querySelectorAll('[name="mbAddress"]')].map((i) => i.value.trim().toLowerCase()).filter(Boolean));
      const added = [];
      for (const value of area.value.split(/[\s,;]+/)) {
        const address = value.trim();
        if (!address || known.has(address.toLowerCase())) continue;
        known.add(address.toLowerCase());
        added.push(address);
      }
      if (added.length === 0) {
        toast('Nenhum e-mail novo para incluir.', 'error');
        return;
      }
      // As linhas em branco dão lugar às caixas incluídas.
      [...rows.querySelectorAll('[data-mailbox-row]')].forEach((row) => {
        if (!row.querySelector('[name="mbAddress"]').value.trim() && !row.querySelector('[name="mbLogin"]').value.trim()) row.remove();
      });
      for (const address of added) rows.insertAdjacentHTML('beforeend', mailboxRow({ address }).toString());
      area.value = '';
      form.querySelector('[data-bulk]').hidden = true;
      sync();
      toast(`${plural(added.length, 'caixa incluída', 'caixas incluídas')}.`, 'success');
    } else if (action === 'connect') {
      await startSignIn(button);
    } else if (action === 'connect-cancel') {
      stopFlow({ cancel: true });
      deviceBox.hidden = true;
      form.querySelector('[data-action="connect"]').focus();
    } else if (action === 'copy-code') {
      try {
        await copyText(button.dataset.code);
        toast('Código copiado.', 'success');
      } catch {
        toast('Não foi possível copiar: digite o código mostrado.', 'error');
      }
    } else if (action === 'cert-generate' || action === 'cert-generate-confirm') {
      const confirmBox = form.querySelector('[data-cert-confirm]');
      // Já há um certificado: confirma antes (o novo precisa ser enviado ao registro do aplicativo).
      if (action === 'cert-generate' && state.certificate) {
        confirmBox.hidden = false;
        confirmBox.querySelector('[data-action="cert-generate-confirm"]').focus();
        return;
      }
      confirmBox.hidden = true;
      const generate = form.querySelector('[data-action="cert-generate"]');
      generate.disabled = true;
      try {
        useCertificate(await post('/api/mail-sources/certificate', { name: form.elements.name.value }));
        toast('Certificado gerado. Baixe o arquivo .cer e envie-o ao registro do aplicativo.', 'success');
      } catch (err) {
        toast(err.message, 'error');
      } finally {
        generate.disabled = false;
        generate.focus();
      }
    } else if (action === 'cert-generate-cancel') {
      form.querySelector('[data-cert-confirm]').hidden = true;
      form.querySelector('[data-action="cert-generate"]').focus();
    } else if (action === 'cert-download') {
      if (state.certificate) downloadCertificate(state.certificate, form.elements.name.value);
    } else if (action === 'test') {
      const box = form.querySelector('[data-test-result]');
      box.className = 'test-result';
      box.textContent = 'Testando a conexão… (pode levar alguns segundos)';
      const hadFocus = document.activeElement === button;
      button.disabled = true;
      try {
        showTest(box, await post('/api/mail-sources/test', readForm(form, existing)));
      } catch (err) {
        showTest(box, { ok: false, message: err.message });
      } finally {
        button.disabled = false;
        if (hadFocus) button.focus(); // o botão desabilitado perde o foco do teclado
      }
    }
  });
  sync();
  syncSaved();
}

function scopeText(s) {
  if (s.type !== 'imap' && s.scope === 'all') {
    const except = s.excludeMailboxes?.length ? ` (exceto ${fmtNum(s.excludeMailboxes.length)})` : '';
    return `Todas as caixas ${TYPES[s.type].domain}${except}`;
  }
  return plural(s.mailboxes.length, 'caixa', 'caixas');
}

/** Credencial da Microsoft (aplicativo com segredo ou certificado, ou conta conectada). */
function microsoftCredential(g = {}) {
  if (g.auth === 'delegated') {
    return g.account ? html`Conta conectada: ${g.account.address || g.account.username}` : html`<span class="chip danger">sem conta conectada</span>`;
  }
  if (g.auth === 'certificate') {
    return html`Aplicativo ${g.clientId || '—'} · certificado${g.certificate ? html` até ${fmtDate(g.certificate.notAfter)}${expiryChip(g.certificate)}` : ''}`;
  }
  return `Aplicativo ${g.clientId || '—'} · ${g.hasClientSecret ? 'segredo salvo' : 'sem segredo'}`;
}

function credentialText(s) {
  if (s.type === 'graph') return microsoftCredential(s.graph);
  if (s.type === 'gmail') return `Conta de serviço ${s.gmail?.clientEmail || '—'}`;
  const im = s.imap || {};
  const server = `${im.host}:${im.port} · ${SECURITY[im.security]?.label || im.security}`;
  return im.auth === 'oauth' ? html`${server} · OAuth da Microsoft<br />${microsoftCredential(s.graph)}` : server;
}

function deletionText(s) {
  if (!s.allowDelete) return html`<span class="muted">Não</span>`;
  return html`<span class="chip danger">${s.deleteMode === 'trash' ? 'para a lixeira' : 'definitiva'}</span>`;
}

export async function render(root) {
  let sources = [];

  const draw = () =>
    paint(
      root,
      html`<div class="page-head">
          <div>
            <h1>Caixas de e-mail</h1>
            <div class="sub">Conexões com o Microsoft 365, o Google Workspace ou servidores IMAP. Senhas, segredos, certificados e as autorizações das contas conectadas ficam gravados cifrados no servidor do CLEAN e não são exibidos novamente.</div>
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
                  <thead><tr><th>Nome</th><th>Tipo</th><th>Caixas</th><th>Credencial</th><th>Exclusão</th><th><span class="sr-only">Ações</span></th></tr></thead>
                  <tbody>
                    ${sources.map(
                      (s) => html`<tr>
                        <td><b>${s.name}</b>${s.description ? html`<div class="muted small">${s.description}</div>` : ''}</td>
                        <td class="nowrap">${TYPES[s.type]?.label || s.type}</td>
                        <td class="small">${scopeText(s)}${s.type === 'imap' || s.scope === 'list' ? html`<div class="muted">${s.mailboxes.slice(0, 3).map((m) => m.address).join(', ')}${s.mailboxes.length > 3 ? '…' : ''}</div>` : ''}</td>
                        <td class="small">${credentialText(s)}</td>
                        <td class="small">${deletionText(s)}</td>
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
          <p class="muted small">
            <b>Microsoft (Exchange Online e Outlook.com):</b> o acesso é por OAuth 2.0 — um aplicativo registrado no Microsoft Entra ID, com segredo
            do cliente ou certificado (todas as caixas do locatário), ou uma conta Microsoft conectada pelo código de entrada (a caixa dela e as
            compartilhadas com ela). A Microsoft não aceita mais senha no IMAP: nas conexões IMAP com o Exchange Online, escolha o OAuth da Microsoft.
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
      if (saved.scheduleWarning) toast(saved.scheduleWarning, 'warn');
      await refresh();
    }
  };

  const test = async (source, button) => {
    button.disabled = true;
    toast(`Testando "${source.name}"…`);
    try {
      const result = await post('/api/mail-sources/test', { ...source, id: source.id, mailboxes: source.type === 'imap' ? source.mailboxes : source.mailboxes.map((m) => m.address) });
      // Se outro diálogo foi aberto enquanto o teste rodava, não o substitui: mostra só um aviso.
      if (document.getElementById('modal').open || !root.isConnected) {
        toast(`Teste de "${source.name}": ${result.message}`, result.ok ? 'success' : 'error');
        return;
      }
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
      const ok = await confirmDialog(`Excluir a conexão "${source.name}"? As senhas, chaves e autorizações salvas serão apagadas. As análises já feitas continuam disponíveis.`, { confirmLabel: 'Excluir' });
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
