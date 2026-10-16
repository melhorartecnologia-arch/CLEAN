// Novo agendamento de reinício (#/intune/novo) ou edição (#/intune/<id>): escolhe a conexão
// Microsoft 365 (credencial do Intune), o escopo (todos os iPhone/iPad ou uma lista) e a regra de
// recorrência, com a prévia das próximas execuções. O reinício em si é enviado pela API Microsoft
// Graph (rebootNow) a cada aparelho nos horários do agendamento.
import { get, post, put } from '../api.js';
import { html, render as paint, icon, toast, fmtServerDateTime } from '../ui.js';
import { go } from '../nav.js';
import { scheduleSection, bindSchedule, readSchedule } from '../schedule-form.js';

export async function render(root, { params, ctx }) {
  const id = params[0];
  const job = id ? await get(`/api/intune/${encodeURIComponent(id)}`) : null;
  const sources = await get('/api/intune/sources');
  const title = job ? 'Editar reinício agendado' : 'Novo reinício agendado';

  if (sources.length === 0) {
    paint(
      root,
      html`<div class="page-head"><div><h1>${title}</h1></div></div>
        <div class="alert info">${icon('info')}<div>
          Para agendar o reinício é preciso ter ao menos uma conexão do <b>Microsoft 365 (Graph)</b> cadastrada: ela é usada como credencial do Intune.
          O aplicativo dessa conexão precisa das permissões <span class="mono">DeviceManagementManagedDevices.Read.All</span> e
          <span class="mono">DeviceManagementManagedDevices.PrivilegedOperations.All</span> (tipo Aplicativo, com consentimento do administrador).
          <div class="inline page-actions"><a class="btn small" href="#/email/caixas">Cadastrar conexão Microsoft 365</a></div>
        </div></div>`,
    );
    return null;
  }

  const scope = job?.scope || { mode: 'all', names: [] };
  const sel = (a, b) => (String(a) === String(b) ? 'selected' : '');
  const chk = (a) => (a ? 'checked' : '');
  const sourceId = job?.sourceId || sources[0].id;

  paint(
    root,
    html`<div class="page-head">
        <div>
          <h1>${title}</h1>
          <div class="sub">Reinicia os aparelhos iPhone/iPad gerenciados no Intune, nos dias e horários escolhidos, pela API Microsoft Graph.</div>
        </div>
      </div>
      <form class="card" data-form novalidate>
        <div class="form-grid">
          <label class="field full">
            <span>Nome do agendamento</span>
            <input type="text" name="name" maxlength="120" required value="${job?.name || ''}" placeholder="Ex.: Reinício noturno dos iPhones" />
          </label>

          <label class="field full">
            <span>Conexão Microsoft 365 (credencial do Intune)</span>
            <select name="sourceId">
              ${sources.map((s) => html`<option value="${s.id}" ${sel(sourceId, s.id)}>${s.name}${s.auth === 'delegated' ? ' (conta conectada)' : ''}</option>`)}
            </select>
            <small>O reinício usa as permissões do aplicativo dessa conexão. É preciso conceder a ele <span class="mono">DeviceManagementManagedDevices.Read.All</span> e <span class="mono">DeviceManagementManagedDevices.PrivilegedOperations.All</span>.</small>
          </label>
          <div class="inline full">
            <button type="button" class="btn small" data-test>${icon('flask')} Testar acesso ao Intune</button>
            <button type="button" class="btn small" data-list-devices>${icon('phone')} Listar aparelhos</button>
            <span class="muted small" data-test-result aria-live="polite"></span>
          </div>

          <fieldset class="full">
            <legend>Quais aparelhos reiniciar</legend>
            <label class="check">
              <input type="radio" name="scopeMode" value="all" ${chk(scope.mode !== 'list')} />
              <span><b>Todos os iPhone e iPad gerenciados</b><br /><small class="muted">Todos os dispositivos Apple (iOS/iPadOS) do Intune nessa conexão.</small></span>
            </label>
            <label class="check">
              <input type="radio" name="scopeMode" value="list" ${chk(scope.mode === 'list')} />
              <span><b>Somente uma lista de aparelhos</b><br /><small class="muted">Informe um por linha: nome do dispositivo, e-mail do usuário ou número de série.</small></span>
            </label>
            <label class="field full" data-scope-list ${scope.mode === 'list' ? '' : 'hidden'}>
              <span class="sr-only">Lista de aparelhos</span>
              <textarea name="names" rows="4" spellcheck="false" placeholder="iPhone de João&#10;joao@empresa.com.br&#10;F2LW80ABCDEF">${(scope.names || []).join('\n')}</textarea>
            </label>
            <div class="devices-preview" data-devices hidden></div>
          </fieldset>

          ${scheduleSection({ kind: 'intune', schedule: job, fixed: true, info: ctx.info, period: false, keep: false })}
        </div>
        <div class="inline page-actions">
          <button type="submit" class="btn primary" data-submit>${icon('check')} Salvar agendamento</button>
          <a class="btn" href="#/intune">Cancelar</a>
        </div>
      </form>`,
  );

  const form = root.querySelector('[data-form]');
  const submit = form.querySelector('[data-submit]');
  const testResult = form.querySelector('[data-test-result]');
  const devicesBox = form.querySelector('[data-devices]');

  const syncScope = () => {
    form.querySelector('[data-scope-list]').hidden = form.elements.scopeMode.value !== 'list';
  };

  const onChange = (event) => {
    if (event.target.name === 'scopeMode') syncScope();
  };
  form.addEventListener('change', onChange);
  const unbind = bindSchedule(form, { scheduleId: job?.id || null, previewUrl: '/api/intune/preview', previewKey: 'jobId' });

  const testBtn = form.querySelector('[data-test]');
  const listBtn = form.querySelector('[data-list-devices]');

  testBtn.addEventListener('click', async () => {
    testBtn.disabled = true;
    paint(testResult, html`Testando…`);
    try {
      const result = await post('/api/intune/test', { sourceId: form.elements.sourceId.value });
      paint(testResult, result.ok ? html`<span class="ok-text">${icon('check')} ${result.message}</span>` : html`<span class="danger-text">${icon('alert')} ${result.message}</span>`);
    } catch (err) {
      paint(testResult, html`<span class="danger-text">${icon('alert')} ${err.message}</span>`);
    } finally {
      testBtn.disabled = false;
    }
  });

  listBtn.addEventListener('click', async () => {
    listBtn.disabled = true;
    devicesBox.hidden = false;
    paint(devicesBox, html`<p class="muted small">Carregando os aparelhos…</p>`);
    try {
      const { devices, total } = await get(`/api/intune/devices?sourceId=${encodeURIComponent(form.elements.sourceId.value)}`);
      paint(
        devicesBox,
        total === 0
          ? html`<p class="muted small">Nenhum iPhone/iPad gerenciado foi encontrado nessa conexão.</p>`
          : html`<p class="muted small">${total} aparelho(s) iPhone/iPad gerenciado(s). Um aparelho que não sincroniza com o Intune há muito tempo provavelmente está offline: o reinício fica pendente até ele reconectar.</p>
              <div class="table-wrap">
                <table class="data compact">
                  <thead><tr><th>Aparelho</th><th>Usuário</th><th>Sistema</th><th>Última sincronização</th><th>Nº de série</th></tr></thead>
                  <tbody>
                    ${devices.map(
                      (d) => html`<tr>
                        <td><b>${d.name}</b>${d.model ? html`<div class="muted small">${d.model}</div>` : ''}</td>
                        <td>${d.userName || d.user || html`<span class="muted">—</span>`}${d.userName && d.user ? html`<div class="muted small">${d.user}</div>` : ''}</td>
                        <td class="nowrap">${d.os}${d.osVersion ? ` ${d.osVersion}` : ''}</td>
                        <td class="nowrap">${d.lastSync ? fmtServerDateTime(d.lastSync, { weekday: false }) : html`<span class="muted">nunca</span>`}${d.stale ? html`<div class="small warn-text">${icon('alert')} provavelmente offline</div>` : ''}</td>
                        <td class="mono small">${d.serialNumber || html`<span class="muted">—</span>`}</td>
                      </tr>`,
                    )}
                  </tbody>
                </table>
              </div>`,
      );
    } catch (err) {
      paint(devicesBox, html`<p class="danger-text small">${icon('alert')} ${err.message}</p>`);
    } finally {
      listBtn.disabled = false;
    }
  });

  const onSubmit = async (event) => {
    event.preventDefault();
    const f = new FormData(form);
    const name = String(f.get('name') || '').trim();
    if (!name) {
      form.elements.name.focus();
      return toast('Dê um nome ao agendamento.', 'error');
    }
    const mode = f.get('scopeMode') === 'list' ? 'list' : 'all';
    const names = String(f.get('names') || '').split(/\r?\n/).map((n) => n.trim()).filter(Boolean);
    if (mode === 'list' && names.length === 0) {
      form.elements.names.focus();
      return toast('Informe ao menos um aparelho na lista (ou escolha "Todos os iPhone e iPad").', 'error');
    }
    const { rule, catchUp } = readSchedule(form);
    const body = { name, sourceId: f.get('sourceId'), scope: { mode, names }, rule, catchUp };
    submit.disabled = true;
    try {
      const saved = await (job ? put(`/api/intune/${job.id}`, body) : post('/api/intune', body));
      const next = saved.nextRunAt ? ` Próxima execução: ${fmtServerDateTime(saved.nextRunAt)}.` : saved.state === 'paused' ? ' Ele está pausado.' : '';
      toast(`Agendamento salvo.${next}`, 'success');
      go('/intune');
    } catch (err) {
      toast(err.message, 'error');
      submit.disabled = false;
    }
  };
  form.addEventListener('submit', onSubmit);
  syncScope();
  return () => {
    unbind();
    form.removeEventListener('submit', onSubmit);
    form.removeEventListener('change', onChange);
  };
}
