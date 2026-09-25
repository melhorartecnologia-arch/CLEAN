// Repositórios: pastas locais ou compartilhamentos de rede que serão analisados.
import { get, post, put, del } from '../api.js';
import { html, render as paint, icon, openDialog, confirmDialog, toast, fmtNum } from '../ui.js';

function repoForm(repo, ctx) {
  const audit = repo?.audit || {};
  const defaults = (ctx.info?.defaultExcludes || []).join(', ');
  return html`<div class="form-grid">
    <label class="field full">
      <span>Nome</span>
      <input type="text" name="name" required maxlength="200" value="${repo?.name || ''}" placeholder="Ex.: Financeiro" />
    </label>
    <div class="field full">
      <label class="field-label" for="repo-path">Caminho da pasta</label>
      <div class="inline">
        <input id="repo-path" type="text" name="path" required value="${repo?.path || ''}" placeholder="\\\\servidor\\compartilhamento\\pasta ou D:\\Dados" />
        <button type="button" class="btn" data-test-path>${icon('check')} Testar acesso</button>
      </div>
      <div class="test-result" data-test-result aria-live="polite"></div>
      <small>O caminho é lido pela conta que executa o CLEAN no servidor (${ctx.info?.user || 'serviço'}), que precisa ter permissão de leitura.</small>
    </div>
    <label class="field full">
      <span>Descrição (opcional)</span>
      <input type="text" name="description" maxlength="1000" value="${repo?.description || ''}" />
    </label>
    <label class="field full">
      <span>Ignorar (um padrão por linha)</span>
      <textarea name="exclude" rows="3" placeholder="*.bak&#10;Backup&#10;Financeiro\\Antigo">${(repo?.exclude || []).join('\n')}</textarea>
      <small>Aceita curingas * e ?. Um nome vale para arquivos e pastas; use barra para caminhos (Pasta\\Subpasta). Sempre ignorados: ${defaults}.</small>
    </label>
    <fieldset class="full">
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
  </div>`;
}

function readForm(form) {
  const f = new FormData(form);
  return {
    name: f.get('name'),
    path: f.get('path'),
    description: f.get('description'),
    exclude: f.get('exclude'),
    audit: {
      enabled: f.get('auditEnabled') === 'on',
      computer: f.get('auditComputer') || '',
      localPath: f.get('auditLocalPath') || '',
      days: Number(f.get('auditDays')) || 30,
      maxEvents: Number(f.get('auditMaxEvents')) || 200000,
      ignoreUsers: f.get('auditIgnoreUsers') || '',
    },
  };
}

function wireForm(form) {
  const result = form.querySelector('[data-test-result]');
  form.querySelector('[data-test-path]').addEventListener('click', async () => {
    const path = form.elements.path.value;
    result.className = 'test-result';
    result.textContent = 'Testando…';
    try {
      const r = await post('/api/repositories/test', { path });
      result.className = `test-result ${r.ok ? 'ok' : 'fail'}`;
      result.textContent = r.ok ? `${r.message}${r.sample?.length ? ` Ex.: ${r.sample.slice(0, 4).join(', ')}` : ''}` : r.message;
    } catch (err) {
      result.className = 'test-result fail';
      result.textContent = err.message;
    }
  });
  const toggle = form.elements.auditEnabled;
  toggle.addEventListener('change', () => {
    form.querySelector('[data-audit-fields]').hidden = !toggle.checked;
  });
}

export async function render(root, { ctx }) {
  let repos = [];

  const draw = () =>
    paint(
      root,
      html`<div class="page-head">
          <div>
            <h1>Repositórios</h1>
            <div class="sub">Pastas locais ou compartilhamentos de rede do Windows que serão analisados.</div>
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
                  <thead><tr><th>Nome</th><th>Caminho</th><th>Ignorar</th><th>Auditoria</th><th><span class="sr-only">Ações</span></th></tr></thead>
                  <tbody>
                    ${repos.map(
                      (r) => html`<tr>
                        <td><b>${r.name}</b>${r.description ? html`<div class="muted small">${r.description}</div>` : ''}</td>
                        <td class="path">${r.path}</td>
                        <td class="small">${r.exclude?.length ? r.exclude.join(', ') : html`<span class="muted">—</span>`}</td>
                        <td class="small">${r.audit?.enabled ? `Sim (${fmtNum(r.audit.days)} dias${r.audit.computer ? `, ${r.audit.computer}` : ''})` : html`<span class="muted">Não</span>`}</td>
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
    repos = await get('/api/repositories');
    draw();
  };

  const edit = async (repo) => {
    const saved = await openDialog({
      title: repo ? 'Editar repositório' : 'Novo repositório',
      body: repoForm(repo, ctx),
      wide: true,
      onOpen: wireForm,
      onSubmit: (form) => (repo ? put(`/api/repositories/${repo.id}`, readForm(form)) : post('/api/repositories', readForm(form))),
    });
    if (saved) {
      toast(repo ? 'Repositório atualizado.' : 'Repositório cadastrado.', 'success');
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
        await del(`/api/repositories/${repo.id}`);
        toast('Repositório excluído.', 'success');
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
