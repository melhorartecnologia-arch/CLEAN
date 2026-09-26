// Repositórios: pastas do Windows (locais ou compartilhamentos de rede) e bibliotecas do OneDrive e
// do SharePoint que serão analisadas.
import { get, post, put, del } from '../api.js';
import { html, render as paint, icon, openDialog, confirmDialog, toast, fmtNum } from '../ui.js';

const TYPES = {
  local: { label: 'Pasta do Windows', detail: 'Servidor de arquivos, NAS ou pasta local' },
  onedrive: { label: 'OneDrive', detail: 'Arquivos dos usuários do Microsoft 365' },
  sharepoint: { label: 'SharePoint', detail: 'Sites e bibliotecas de documentos' },
};
const SAVED = '•••••• salvo — deixe em branco para manter';
const isCloud = (type) => type === 'onedrive' || type === 'sharepoint';

function repoForm(repo, ctx, mailSources) {
  const type = repo?.type || 'local';
  const audit = repo?.audit || {};
  const g = repo?.graph || {};
  const cloud = repo?.cloud || {};
  const scope = cloud.scope || 'all';
  const defaults = (ctx.info?.defaultExcludes || []).join(', ');
  const m365 = mailSources.filter((s) => s.type === 'graph');
  return html`<div class="form-grid">
    <label class="field full">
      <span>Nome</span>
      <input type="text" name="name" required maxlength="200" value="${repo?.name || ''}" placeholder="Ex.: Financeiro" />
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

    <div class="field full" data-kind="local">
      <label class="field-label" for="repo-path">Caminho da pasta</label>
      <div class="inline">
        <input id="repo-path" type="text" name="path" value="${type === 'local' ? repo?.path || '' : ''}" placeholder="\\\\servidor\\compartilhamento\\pasta ou D:\\Dados" />
        <button type="button" class="btn" data-test-path>${icon('check')} Testar acesso</button>
      </div>
      <div class="test-result" data-test-result aria-live="polite"></div>
      <small>O caminho é lido pela conta que executa o CLEAN no servidor (${ctx.info?.user || 'serviço'}), que precisa ter permissão de leitura.</small>
    </div>

    <fieldset class="full" data-kind="cloud">
      <legend>Microsoft 365 (Microsoft Graph)</legend>
      ${m365.length
        ? html`<label class="field full">
            <span>Credenciais</span>
            <select name="credentialsFrom">
              <option value="">Informar manualmente</option>
              ${m365.map((s) => html`<option value="${s.id}" ${repo?.credentialsFrom === s.id ? 'selected' : ''}>Usar as credenciais da conexão "${s.name}"</option>`)}
            </select>
            <small>O mesmo registro de aplicativo pode ler e-mails e arquivos (inclua nele as permissões de arquivos abaixo). As credenciais ficam ligadas à conexão: um novo segredo salvo nela também vale aqui.</small>
          </label>`
        : ''}
      <div class="form-grid" data-credentials>
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
          <li>No centro de administração do Microsoft Entra (entra.microsoft.com), abra <b>Registros de aplicativo › Novo registro</b> (ex.: "CLEAN", somente esta organização) — ou use o mesmo registro das caixas de e-mail.</li>
          <li>Em <b>Permissões de API › Adicionar › Microsoft Graph › Permissões de aplicativo</b>, inclua <code>Files.Read.All</code>, <code>Sites.Read.All</code> e <code>User.Read.All</code> e clique em <b>Conceder consentimento do administrador</b>. Para excluir arquivos, use <code>Files.ReadWrite.All</code> no lugar de <code>Files.Read.All</code>.</li>
          <li>Em <b>Certificados e segredos › Novo segredo do cliente</b>, copie o <b>Valor</b> (não o ID do segredo).</li>
          <li>Na página <b>Visão geral</b>, copie o ID do aplicativo (cliente) e o ID do diretório (locatário).</li>
        </ol>
      </details>
    </fieldset>

    <fieldset class="full" data-kind="cloud">
      <legend data-scope-title>Contas e sites</legend>
      <label class="check"><input type="radio" name="scope" value="all" ${scope === 'all' ? 'checked' : ''} /><span data-scope-all>Todas</span></label>
      <label class="check"><input type="radio" name="scope" value="list" ${scope === 'list' ? 'checked' : ''} /><span data-scope-list>Somente as informadas</span></label>
      <label class="field" data-target="onedrive-list">
        <span>Contas (e-mail do usuário, uma por linha)</span>
        <textarea name="accounts" rows="4" placeholder="ana@empresa.com.br&#10;financeiro@empresa.com.br">${(cloud.accounts || []).join('\n')}</textarea>
      </label>
      <label class="field" data-target="sharepoint-list">
        <span>Sites (endereço, um por linha)</span>
        <textarea name="sites" rows="4" placeholder="https://empresa.sharepoint.com/sites/Financeiro&#10;https://empresa.sharepoint.com/sites/RH">${(cloud.sites || []).join('\n')}</textarea>
        <small>Os subsites e todas as bibliotecas de documentos de cada site são analisados. Pode colar o endereço de uma página ou biblioteca do site.</small>
      </label>
      <label class="field" data-target="onedrive">
        <span>Ignorar estas contas (uma por linha; aceita *)</span>
        <textarea name="excludeAccounts" rows="2" placeholder="teste@*&#10;sala.*@empresa.com.br">${type === 'onedrive' ? (cloud.exclude || []).join('\n') : ''}</textarea>
      </label>
      <label class="field" data-target="sharepoint">
        <span>Ignorar estes sites (endereço ou nome, um por linha; aceita *)</span>
        <textarea name="excludeSites" rows="2" placeholder="https://empresa.sharepoint.com/sites/Arquivo*&#10;Projetos antigos">${type === 'sharepoint' ? (cloud.exclude || []).join('\n') : ''}</textarea>
      </label>
    </fieldset>

    <label class="field full">
      <span>Descrição (opcional)</span>
      <input type="text" name="description" maxlength="1000" value="${repo?.description || ''}" />
    </label>
    <label class="field full">
      <span>Ignorar (um padrão por linha)</span>
      <textarea name="exclude" rows="3" placeholder="*.bak&#10;Backup&#10;Financeiro\\Antigo">${(repo?.exclude || []).join('\n')}</textarea>
      <small>Aceita curingas * e ?. Um nome vale para arquivos e pastas (no OneDrive e no SharePoint, também para bibliotecas, ex.: "Site Assets"); use barra para caminhos (Pasta\\Subpasta). Sempre ignorados: ${defaults}.</small>
    </label>

    <fieldset class="full">
      <legend>Exclusão dos arquivos encontrados</legend>
      <label class="check">
        <input type="checkbox" name="allowDelete" ${repo?.allowDelete ? 'checked' : ''} />
        <span><b>Permitir excluir os arquivos em que os termos forem encontrados</b><br /><small class="muted">Na análise ("analisar e excluir") ou item a item pelo relatório.</small></span>
      </label>
      <label class="field" data-kind="cloud" data-delete-mode>
        <span>Como excluir</span>
        <select name="deleteMode">
          <option value="trash" ${repo?.deleteMode !== 'permanent' ? 'selected' : ''}>Mover para a Lixeira (restaurável)</option>
          <option value="permanent" ${repo?.deleteMode === 'permanent' ? 'selected' : ''}>Excluir definitivamente</option>
        </select>
      </label>
      <p class="hint" data-kind="local">
        A exclusão é <b>definitiva</b>: arquivos apagados pela rede não vão para a Lixeira. A conta do CLEAN
        (${ctx.info?.user || 'serviço'}) precisa de permissão de <b>modificação</b> (NTFS e compartilhamento) nesta pasta.
      </p>
      <p class="hint" data-kind="cloud">
        O aplicativo precisa da permissão <b>Files.ReadWrite.All</b> (ou Sites.ReadWrite.All). Na lixeira, os arquivos podem
        ser restaurados pelo usuário ou pelo administrador do site (por até 93 dias). Arquivos abertos para edição, em
        check-out ou com rótulo de retenção não são excluídos.
      </p>
    </fieldset>

    <fieldset class="full" data-kind="local">
      <legend>Log de auditoria do Windows (opcional)</legend>
      <label class="check">
        <input type="checkbox" name="auditEnabled" ${audit.enabled ? 'checked' : ''} />
        <span>Consultar o log de Segurança para descobrir quem acessou ou alterou cada arquivo por último</span>
      </label>
      <div class="form-grid" data-audit-fields ${audit.enabled ? '' : 'hidden'}>
        <label class="field">
          <span>Computador</span>
          <input type="text" name="auditComputer" value="${audit.computer || ''}" placeholder="padrão: servidor do caminho \\\\servidor" />
        </label>
        <label class="field">
          <span>Caminho local no servidor</span>
          <input type="text" name="auditLocalPath" value="${audit.localPath || ''}" placeholder="opcional, ex.: E:\\Compartilhamentos\\Financeiro" />
        </label>
        <label class="field">
          <span>Período (dias)</span>
          <input type="number" name="auditDays" min="1" max="365" value="${audit.days || 30}" />
        </label>
        <label class="field">
          <span>Máximo de eventos lidos</span>
          <input type="number" name="auditMaxEvents" min="100" max="5000000" step="1000" value="${audit.maxEvents || 200000}" />
        </label>
        <label class="field full">
          <span>Contas ignoradas (uma por linha)</span>
          <textarea name="auditIgnoreUsers" rows="2" placeholder="svc-backup&#10;EMPRESA\\svc-antivirus">${(audit.ignoreUsers || []).join('\n')}</textarea>
          <small>Contas de backup, antivírus e indexação leem todos os arquivos e não devem aparecer como último usuário. A conta do CLEAN (${ctx.info?.user || 'serviço'}) é sempre ignorada.</small>
        </label>
        <p class="hint full">
          Requer a auditoria habilitada no servidor de arquivos (eventos 4663 "Auditoria do Sistema de Arquivos" e/ou 5145
          "Auditoria Detalhada de Compartilhamento") e que a conta do CLEAN possa ler o log de Segurança (Administradores ou
          "Leitores de Log de Eventos"). O caminho local é descoberto automaticamente pelos eventos 5145; informe-o se usar apenas 4663.
        </p>
      </div>
    </fieldset>

    <div class="field full" data-kind="cloud">
      <div class="inline"><button type="button" class="btn" data-test-cloud>${icon('check')} Testar conexão</button></div>
      <div class="test-result" data-cloud-result aria-live="polite"></div>
    </div>
  </div>`;
}

function readForm(form, existing) {
  const f = new FormData(form);
  const type = f.get('type') || 'local';
  const body = {
    type,
    name: f.get('name'),
    description: f.get('description'),
    exclude: f.get('exclude'),
    allowDelete: f.get('allowDelete') === 'on',
  };
  if (existing) body.id = existing.id;
  if (!isCloud(type)) {
    body.path = f.get('path');
    body.audit = {
      enabled: f.get('auditEnabled') === 'on',
      computer: f.get('auditComputer') || '',
      localPath: f.get('auditLocalPath') || '',
      days: Number(f.get('auditDays')) || 30,
      maxEvents: Number(f.get('auditMaxEvents')) || 200000,
      ignoreUsers: f.get('auditIgnoreUsers') || '',
    };
    return body;
  }
  const from = f.get('credentialsFrom');
  if (from) body.credentialsFrom = from;
  else body.graph = { tenantId: f.get('tenantId'), clientId: f.get('clientId'), clientSecret: f.get('clientSecret') };
  body.scope = f.get('scope') || 'all';
  body.accounts = type === 'onedrive' ? f.get('accounts') : '';
  body.sites = type === 'sharepoint' ? f.get('sites') : '';
  body.excludeTargets = type === 'onedrive' ? f.get('excludeAccounts') : f.get('excludeSites');
  body.deleteMode = f.get('deleteMode') || 'trash';
  return body;
}

function showResult(box, result) {
  box.className = `test-result ${result.ok ? 'ok' : 'fail'}`;
  paint(box, html`${result.message}${result.details?.length ? html`<ul class="test-details">${result.details.map((d) => html`<li>${d}</li>`)}</ul>` : ''}`);
}

function wireForm(form, existing) {
  const sync = () => {
    const type = form.elements.type.value || 'local';
    const cloud = isCloud(type);
    form.querySelectorAll('[data-kind]').forEach((el) => {
      el.hidden = el.dataset.kind !== (cloud ? 'cloud' : 'local');
    });
    if (cloud) {
      const scope = form.elements.scope.value || 'all';
      form.querySelector('[data-target="onedrive-list"]').hidden = !(type === 'onedrive' && scope === 'list');
      form.querySelector('[data-target="sharepoint-list"]').hidden = !(type === 'sharepoint' && scope === 'list');
      form.querySelector('[data-target="onedrive"]').hidden = type !== 'onedrive';
      form.querySelector('[data-target="sharepoint"]').hidden = type !== 'sharepoint';
      form.querySelector('[data-scope-title]').textContent = type === 'onedrive' ? 'Contas do OneDrive' : 'Sites do SharePoint';
      form.querySelector('[data-scope-all]').textContent = type === 'onedrive' ? 'Todas as contas do locatário (usuários com OneDrive)' : 'Todos os sites (sem os OneDrive pessoais)';
      form.querySelector('[data-scope-list]').textContent = type === 'onedrive' ? 'Somente as contas informadas' : 'Somente os sites informados';
      const copied = Boolean(form.elements.credentialsFrom?.value);
      form.querySelector('[data-credentials]').hidden = copied;
      form.querySelector('[data-delete-mode]').hidden = !form.elements.allowDelete.checked;
    }
    form.querySelector('[data-audit-fields]').hidden = !form.elements.auditEnabled.checked;
  };
  // Um resultado de "Testar conexão" deixa de valer quando o tipo, as credenciais ou o alcance mudam.
  const TEST_FIELDS = new Set(['type', 'credentialsFrom', 'tenantId', 'clientId', 'clientSecret', 'scope', 'accounts', 'sites', 'excludeAccounts', 'excludeSites']);
  const clearTest = (event) => {
    if (!TEST_FIELDS.has(event.target.name)) return;
    const box = form.querySelector('[data-cloud-result]');
    box.className = 'test-result';
    box.textContent = '';
  };
  form.addEventListener('change', (event) => {
    clearTest(event);
    sync();
  });
  form.addEventListener('input', clearTest);
  sync();

  const pathResult = form.querySelector('[data-test-result]');
  form.querySelector('[data-test-path]').addEventListener('click', async () => {
    pathResult.className = 'test-result';
    pathResult.textContent = 'Testando…';
    try {
      const r = await post('/api/repositories/test', { path: form.elements.path.value });
      pathResult.className = `test-result ${r.ok ? 'ok' : 'fail'}`;
      pathResult.textContent = r.ok ? `${r.message}${r.sample?.length ? ` Ex.: ${r.sample.slice(0, 4).join(', ')}` : ''}` : r.message;
    } catch (err) {
      pathResult.className = 'test-result fail';
      pathResult.textContent = err.message;
    }
  });
  const cloudResult = form.querySelector('[data-cloud-result]');
  const testButton = form.querySelector('[data-test-cloud]');
  testButton.addEventListener('click', async () => {
    cloudResult.className = 'test-result';
    cloudResult.textContent = 'Testando…';
    const hadFocus = document.activeElement === testButton;
    testButton.disabled = true;
    try {
      showResult(cloudResult, await post('/api/repositories/test', readForm(form, existing)));
    } catch (err) {
      showResult(cloudResult, { ok: false, message: err.message });
    } finally {
      testButton.disabled = false;
      if (hadFocus) testButton.focus(); // o botão desabilitado perde o foco do teclado
    }
  });
}

function whereText(r) {
  return r.type === 'onedrive' || r.type === 'sharepoint' ? r.path : html`<span class="path">${r.path}</span>`;
}

function ignoredText(r) {
  const parts = [];
  if (r.exclude?.length) parts.push(r.exclude.join(', '));
  if (r.cloud?.exclude?.length) parts.push(`${r.type === 'onedrive' ? 'contas' : 'sites'}: ${r.cloud.exclude.join(', ')}`);
  return parts.length ? parts.join(' · ') : html`<span class="muted">—</span>`;
}

function deletionText(r) {
  if (!r.allowDelete) return html`<span class="muted">Não</span>`;
  const mode = isCloud(r.type) ? (r.deleteMode === 'permanent' ? ' (definitiva)' : ' (lixeira)') : '';
  return html`<span class="chip danger">permitida${mode}</span>`;
}

export async function render(root, { ctx }) {
  let repos = [];
  let mailSources = [];

  const draw = () =>
    paint(
      root,
      html`<div class="page-head">
          <div>
            <h1>Repositórios</h1>
            <div class="sub">Pastas do Windows (locais ou compartilhamentos de rede) e bibliotecas do OneDrive e do SharePoint que serão analisadas.</div>
          </div>
          <div class="actions"><button class="btn primary" data-action="new">${icon('plus')} Novo repositório</button></div>
        </div>
        <section class="card">
          ${repos.length === 0
            ? html`<div class="empty">
                <p>Nenhum repositório cadastrado.</p>
                <button class="btn primary" data-action="new">${icon('plus')} Cadastrar o primeiro</button>
              </div>`
            : html`<div class="table-wrap">
                <table class="data">
                  <thead><tr><th>Nome</th><th>Tipo</th><th>Local</th><th>Ignorar</th><th>Auditoria</th><th>Exclusão</th><th><span class="sr-only">Ações</span></th></tr></thead>
                  <tbody>
                    ${repos.map(
                      (r) => html`<tr>
                        <td><b>${r.name}</b>${r.description ? html`<div class="muted small">${r.description}</div>` : ''}</td>
                        <td class="small nowrap">${TYPES[r.type || 'local'].label}</td>
                        <td class="small">${whereText(r)}</td>
                        <td class="small">${ignoredText(r)}</td>
                        <td class="small">${r.audit?.enabled ? `Sim (${fmtNum(r.audit.days)} dias${r.audit.computer ? `, ${r.audit.computer}` : ''})` : html`<span class="muted">Não</span>`}</td>
                        <td class="small">${deletionText(r)}</td>
                        <td class="actions">
                          <button class="icon-btn" data-action="edit" data-id="${r.id}" aria-label="Editar ${r.name}" title="Editar">${icon('edit')}</button>
                          <button class="icon-btn danger" data-action="delete" data-id="${r.id}" aria-label="Excluir ${r.name}" title="Excluir">${icon('trash')}</button>
                        </td>
                      </tr>`,
                    )}
                  </tbody>
                </table>
              </div>`}
        </section>`,
    );

  const refresh = async () => {
    [repos, mailSources] = await Promise.all([get('/api/repositories'), get('/api/mail-sources').catch(() => [])]);
    draw();
  };

  const edit = async (repo) => {
    const saved = await openDialog({
      title: repo ? 'Editar repositório' : 'Novo repositório',
      body: repoForm(repo, ctx, mailSources),
      wide: true,
      onOpen: (form) => wireForm(form, repo),
      onSubmit: (form) => (repo ? put(`/api/repositories/${repo.id}`, readForm(form, repo)) : post('/api/repositories', readForm(form, null))),
    });
    if (saved) {
      toast(repo ? 'Repositório atualizado.' : 'Repositório cadastrado.', 'success');
      if (saved.scheduleWarning) toast(saved.scheduleWarning, 'warn');
      await refresh();
    }
  };

  const onClick = async (event) => {
    const button = event.target.closest('[data-action]');
    if (!button) return;
    const repo = repos.find((r) => r.id === button.dataset.id);
    if (button.dataset.action === 'new') return edit(null);
    if (button.dataset.action === 'edit') return edit(repo);
    if (button.dataset.action === 'delete') {
      const ok = await confirmDialog(`Excluir o repositório "${repo.name}"? As análises já feitas continuam disponíveis.`, { confirmLabel: 'Excluir' });
      if (!ok) return;
      try {
        const result = await del(`/api/repositories/${repo.id}`);
        toast('Repositório excluído.', 'success');
        if (result?.scheduleWarning) toast(result.scheduleWarning, 'warn');
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
