// Nova política de retenção (#/retencao/nova?tipo=email) ou edição (#/retencao/<id>): arquivos ou
// mensagens mais antigos que a idade máxima (ou sem uso desde uma data fixa), pelo critério de data
// escolhido, são listados no relatório e — se a política excluir — eliminados. Arquivos: todos os
// repositórios marcados ou só algumas pastas deles. Execução manual ou agendada.
import { get, post, put } from '../api.js';
import { html, render as paint, icon, toast, fmtServerDateTime, plural } from '../ui.js';
import { go } from '../nav.js';
import { scheduleSection, bindSchedule, readSchedule, scheduling, keepField } from '../schedule-form.js';
import { FILE_CRITERIA, MAIL_CRITERIA, UNITS, DEFAULT_MAX_DELETIONS, ACCESS_CRITERIA, describeRetention, cutoffPreview, periodPreview, sinceInstant, deletionsText } from '../retention.js';

const CLOUD_LABELS = { onedrive: 'OneDrive', sharepoint: 'SharePoint' };
const TYPE_LABELS = { graph: 'Microsoft 365', gmail: 'Google Workspace', imap: 'IMAP' };

function sourceDetail(s) {
  if (s.type !== 'imap' && s.scope === 'all') return `${TYPE_LABELS[s.type]} · todas as caixas`;
  return `${TYPE_LABELS[s.type] || s.type} · ${plural(s.mailboxes.length, 'caixa', 'caixas')}`;
}

/** O que a data fixa quer dizer, pelo critério: "sem uso desde", "criados antes de", "recebidas antes de". */
function sinceWords(criterion) {
  if (criterion === 'received') return 'recebidas antes de';
  if (criterion === 'created') return 'criados antes de';
  return `${FILE_CRITERIA[criterion]?.short || 'sem uso'} desde`;
}

const capitalize = (text) => `${text[0].toUpperCase()}${text.slice(1)}`;

const PERIOD_ERROR = 'Informe o período: a data "a partir de", a data "até" ou as duas, de 02/01/1980 até hoje (datas futuras não valem), com a inicial antes ou igual à final.';

/** Período na prévia, por extenso: "arquivos modificados pela última vez entre 01/01/2015 e 31/12/2018". */
function periodPhrase(criterion, r) {
  const p = periodPreview(r);
  const [a, b] = [p.from, p.to].map((d) => d?.toLocaleDateString('pt-BR'));
  const range = a && b ? `entre ${a} e ${b}` : a ? `a partir de ${a} (até o momento em que cada execução é criada)` : `até ${b}`;
  return `${EXPIRES_PERIOD[criterion] || EXPIRES_PERIOD.used} ${range}`;
}

/**
 * Rótulos dos campos do período, pelo critério: "Somente arquivos alterados a partir de" e
 * "Somente arquivos alterados até" (última modificação)...
 */
/** Quem entra sem a data "até", pelo critério: "os arquivos alterados recentemente"... */
const RECENT_VERB = { used: 'usados', 'opened-or-modified': 'abertos ou alterados', modified: 'alterados', accessed: 'abertos', created: 'criados' };

const PERIOD_LABELS = {
  used: 'Somente arquivos usados pela última vez (modificação, acesso ou criação)',
  'opened-or-modified': 'Somente arquivos abertos ou alterados pela última vez',
  modified: 'Somente arquivos alterados',
  accessed: 'Somente arquivos abertos pela última vez',
  created: 'Somente arquivos criados',
};

/** Hoje ("AAAA-MM-DD", hora local): o limite do campo de data (o servidor recusa datas futuras). */
function todayIso() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

/** O que expira, na prévia: "arquivos modificados pela última vez antes de 26/09/2021". */
const EXPIRES = {
  used: (d) => `arquivos sem nenhuma data (modificação, último acesso ou criação) a partir de ${d}`,
  'opened-or-modified': (d) => `arquivos sem abertura (último acesso) nem modificação a partir de ${d}`,
  modified: (d) => `arquivos modificados pela última vez antes de ${d}`,
  accessed: (d) => `arquivos acessados pela última vez antes de ${d}`,
  created: (d) => `arquivos criados antes de ${d}`,
  received: (d) => `mensagens recebidas antes de ${d}`,
};

/** O que expira num período, na prévia (seguido de "entre ...", "a partir de ..." ou "até ..."). */
const EXPIRES_PERIOD = {
  used: 'arquivos cuja data mais recente (modificação, último acesso ou criação) foi',
  'opened-or-modified': 'arquivos cuja última abertura (último acesso) ou modificação — a mais recente das duas — foi',
  modified: 'arquivos alterados (modificados) pela última vez',
  accessed: 'arquivos abertos (acessados) pela última vez',
  created: 'arquivos criados',
};

export async function render(root, { params, query, ctx }) {
  const id = params[0];
  const policy = id ? await get(`/api/schedules/${encodeURIComponent(id)}`) : null;
  if (policy && policy.purpose !== 'retention') {
    // Um agendamento de análise aberto pelo endereço das políticas (sem nova entrada no histórico:
    // "Voltar" não cai de novo aqui).
    location.replace(`#/agendamentos/${policy.id}`);
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
  // Forma do corte salva: período (arquivos), data fixa ou idade máxima.
  const savedMode = !mail && (r.from || r.to) ? 'period' : r.since ? 'date' : 'age';
  // Pastas salvas, por repositório (uma por linha no campo de cada um).
  const savedFolders = (tid) => (r.folders || []).filter((f) => f.repositoryId === tid).map((f) => f.path.split('/').join(isCloud(targets.find((t) => t.id === tid) || {}) ? '/' : '\\'));

  paint(
    root,
    html`<div class="page-head">
        <div>
          <h1>${title}</h1>
          <div class="sub">${mail
            ? 'Mensagens mais antigas que a idade máxima, ou recebidas antes de uma data, são listadas no relatório e, se a política excluir, eliminadas'
            : 'Arquivos (de qualquer tipo) mais antigos que a idade máxima, sem uso desde uma data ou com a data do critério dentro de um período são listados no relatório e, se a política excluir, eliminados'} — cada exclusão fica no Registro de exclusões.</div>
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

          ${mail
            ? ''
            : html`<fieldset class="full" data-folders-box>
                <legend>Pastas (opcional)</legend>
                <p class="muted small">Para limpar só algumas pastas de um repositório, informe-as abaixo, uma por linha; as subpastas entram. Em branco: o repositório inteiro. Os arquivos fora das pastas não são listados nem excluídos — nem com a exclusão manual pelo relatório.</p>
                <div class="stack">
                  ${targets.map(
                    (t) => html`<label class="field" data-folders-for="${t.id}" hidden>
                      <span>Somente estas pastas de <b>${t.name}</b></span>
                      <textarea name="folders:${t.id}" rows="2" spellcheck="false" placeholder="${isCloud(t) ? 'Projetos/2019\nDocumentos Compartilhados/Financeiro' : `Financeiro\\2019\n${t.path.replace(/[\\/]+$/, '')}\\Compartilhado\\Antigos`}">${savedFolders(t.id).join('\n')}</textarea>
                      <small>${isCloud(t)
                        ? 'Caminho dentro da biblioteca (vale em todas as bibliotecas) ou com o nome dela na frente (Documentos Compartilhados/Financeiro). Só o nome da biblioteca: a biblioteca inteira.'
                        : html`Relativo à raiz do repositório (Financeiro\\2019) ou o caminho completo dentro dele (<span class="mono">${t.path}</span>…). A pasta precisa existir e não pode ser atalho, link simbólico ou junção.`}</small>
                    </label>`,
                  )}
                </div>
              </fieldset>`}

          <fieldset>
            <legend>Regra de data</legend>
            <div class="stack">
            ${mail
              ? html`<div class="field"><span>Critério de data</span><b>${MAIL_CRITERIA.received.label}</b><small>${MAIL_CRITERIA.received.hint}</small></div>`
              : html`<label class="field">
                  <span>Critério de data</span>
                  <select name="criterion">${Object.entries(FILE_CRITERIA).map(([value, c]) => html`<option value="${value}" ${sel(criterion, value)}>${c.label}</option>`)}</select>
                  <small data-criterion-hint>${criteria[criterion]?.hint}</small>
                </label>`}
            <div class="field">
              <span id="cut-label">Corte</span>
              <div class="option-row" role="radiogroup" aria-labelledby="cut-label">
                <label class="check"><input type="radio" name="cutMode" value="age" ${chk(savedMode === 'age')} /><span>Idade máxima</span></label>
                <label class="check"><input type="radio" name="cutMode" value="date" ${chk(savedMode === 'date')} /><span>Data fixa (<span data-since-word>${sinceWords(criterion)}</span>)</span></label>
                ${mail ? '' : html`<label class="check"><input type="radio" name="cutMode" value="period" ${chk(savedMode === 'period')} /><span>Período (a partir de / até)</span></label>`}
              </div>
            </div>
            <div class="field" data-cut="age" ${savedMode === 'age' ? '' : 'hidden'}>
              <span id="age-label">Idade máxima</span>
              <div class="option-row" role="group" aria-labelledby="age-label">
                <input type="number" name="amount" min="1" max="${UNITS[r.unit || 'years'].max}" value="${r.amount || 5}" aria-label="Idade máxima (quantidade)" />
                <select name="unit" aria-label="Unidade da idade máxima">
                  ${Object.entries(UNITS).map(([value, u]) => html`<option value="${value}" ${sel(r.unit || 'years', value)}>${u.many}</option>`)}
                </select>
              </div>
              <small>${mail ? 'Mensagens' : 'Arquivos'} com a data do critério mais antiga que isso expiram. O corte anda a cada execução.</small>
            </div>
            <label class="field" data-cut="date" ${savedMode === 'date' ? '' : 'hidden'}>
              <span data-since-label>${capitalize(sinceWords(criterion))}</span>
              <input type="date" name="since" min="1980-01-02" max="${todayIso()}" value="${r.since || ''}" />
              <small>${mail ? 'Mensagens recebidas antes desse dia expiram' : 'Arquivos com a data do critério anterior a esse dia (00:00, hora do servidor) expiram: com "Sem abertura nem modificação", os que não foram abertos nem alterados desde então'}. A data não anda: todas as execuções usam o mesmo corte.</small>
            </label>
            ${mail
              ? ''
              : html`<div class="field" data-cut="period" ${savedMode === 'period' ? '' : 'hidden'}>
                  <label class="field"><span data-period-label="from">${PERIOD_LABELS[criterion] || PERIOD_LABELS.used} a partir de</span><input type="date" name="periodFrom" min="1980-01-02" max="${todayIso()}" value="${r.from || ''}" /></label>
                  <label class="field"><span data-period-label="to">${PERIOD_LABELS[criterion] || PERIOD_LABELS.used} até</span><input type="date" name="periodTo" min="1980-01-02" max="${todayIso()}" value="${r.to || ''}" /></label>
                  <small>Preencha uma das datas ou as duas; cada dia conta inteiro (hora do servidor). Só "a partir de": entram também os arquivos mais recentes, até o momento em que cada execução é criada. Só "até": todos os mais antigos que o fim do dia escolhido. Com "Sem abertura nem modificação", vale a mais recente das duas datas: um arquivo aberto ou modificado depois do fim do período fica. As datas não andam: todas as execuções usam o mesmo período.</small>
                </div>`}
            ${mail
              ? html`<label class="check"><input type="checkbox" name="includeTrash" ${chk(r.includeTrash !== false)} /><span>Incluir a Lixeira (Itens Excluídos)</span></label>
                  <label class="check"><input type="checkbox" name="includeJunk" ${chk(r.includeJunk !== false)} /><span>Incluir o Lixo Eletrônico (spam)</span></label>`
              : html`<label class="field">
                  <span>Somente os arquivos com estes nomes (opcional)</span>
                  <textarea name="patterns" rows="3" spellcheck="false" placeholder="*.tmp&#10;*.bak&#10;*.log">${(r.patterns || []).join('\n')}</textarea>
                  <small>Um padrão por linha: * vale qualquer sequência e ? um caractere (ex.: *.tmp). Em branco: todos os arquivos. Os arquivos ignorados no repositório e por padrão (temporários do Office ~$*, Thumbs.db...) continuam de fora.</small>
                </label>`}
            </div>
          </fieldset>

          <div class="schedule-preview full" data-rule-preview aria-live="polite"></div>

          ${mail
            ? ''
            : html`<div class="alert full" data-access-warning hidden>
                  ${icon('alert')}
                  <div>
                    <b>O último acesso (abertura) só é confiável se o Windows do servidor de arquivos o registrar</b> — em muitos servidores esse registro fica desligado.
                    Para conferir, rode no servidor <span class="mono">fsutil behavior query disablelastaccess</span>: 0 ou 2, registra; 1 ou 3, não registra.
                    Sem o registro, a data fica parada (na criação ou na cópia do arquivo) e um arquivo aberto todos os dias, mas não alterado, parece antigo — e seria excluído.
                    Com <b>Sem abertura nem modificação</b>, nesse caso só a modificação protege o arquivo. Isso vale também para <b>Sem uso</b>, que então só protege os arquivos modificados ou criados recentemente.
                    No sentido oposto, programas que leem todos os arquivos (antivírus, backup, indexação e as análises de conteúdo do CLEAN) podem atualizar o último acesso de tudo — e aí nada expira por esse critério.
                  </div>
                </div>
                <div class="alert error full" data-cloud-warning hidden>${icon('alert')}<div>O OneDrive e o SharePoint não informam o último acesso (abertura) de cada arquivo: escolha outro critério ou desmarque os repositórios da nuvem.</div></div>`}

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
                  <span><b>Para a lixeira</b><br /><small class="muted" data-trash-hint>${mail
                    ? 'Vai para Itens Excluídos (Microsoft 365) ou para a Lixeira (Google e IMAP), de onde ainda pode ser recuperada; as que já estão na Lixeira ficam lá até o provedor apagá-las.'
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
                  Cada execução confere o cadastro: se ${mail ? 'uma conexão' : 'um repositório'} deixar de permitir a exclusão ou mudar de ${mail ? 'conta, servidor ou caixas' : 'caminho, contas ou sites'}, as execuções deixam de ser iniciadas até a política ser salva e confirmada de novo. A forma de exclusão é sempre a desta política.
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

          <div class="form-grid full">${keepField(policy, 'desta política', 'Os relatórios das simulações contam à parte.')}</div>
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
  const checkedTargets = () => [...form.querySelectorAll('[name="targetIds"]:checked')];
  const cloudChosen = () => checkedTargets().some((el) => el.dataset.cloud);
  const cutMode = () => {
    const value = form.querySelector('[name="cutMode"]:checked')?.value;
    return value === 'date' || value === 'period' ? value : 'age';
  };

  /** Pastas escolhidas, só dos repositórios marcados: [{ repositoryId, path }] (o servidor valida). */
  const readFolders = () =>
    checkedTargets().flatMap((el) =>
      String(form.elements.namedItem(`folders:${el.value}`)?.value || '')
        .split(/\r?\n/)
        .map((p) => p.trim())
        .filter(Boolean)
        .map((p) => ({ repositoryId: el.value, path: p })),
    );

  /** Pasta para a prévia: o caminho completo dentro do repositório vira relativo (como o servidor salva). */
  const shortFolder = (f) => {
    const t = targets.find((x) => x.id === f.repositoryId);
    if (!t || isCloud(t)) return f.path;
    const norm = (v) => String(v).replace(/[\\/]+/g, '/').replace(/\/+$/, '');
    const [rootPath, value] = [norm(t.path), norm(f.path)];
    if (value.toLowerCase() === rootPath.toLowerCase()) return '(repositório inteiro)';
    return value.toLowerCase().startsWith(`${rootPath.toLowerCase()}/`) ? value.slice(rootPath.length + 1) : value;
  };

  /** A regra como a API espera. */
  const readRetention = () => {
    const f = new FormData(form);
    const folders = mail ? [] : readFolders();
    // Data digitada pela metade (sem o ano, por exemplo): o navegador a entrega vazia. Ela vai como
    // inválida, e não em branco — senão o período viraria só "a partir de" ou só "até" sem aviso.
    const day = (name) => (form.elements[name]?.validity?.badInput ? 'incompleta' : String(f.get(name) || ''));
    return {
      criterion: currentCriterion(),
      // Idade máxima (o corte anda a cada execução) ou data fixa (o mesmo corte sempre).
      ...(cutMode() === 'period'
        ? { mode: 'period', from: day('periodFrom'), to: day('periodTo') }
        : cutMode() === 'date'
          ? { mode: 'date', since: day('since') }
          : { mode: 'age', amount: Number(f.get('amount')), unit: f.get('unit') }),
      ...(mail
        ? { includeTrash: f.get('includeTrash') === 'on', includeJunk: f.get('includeJunk') === 'on' }
        : { patterns: String(f.get('patterns') || '').split(/\r?\n/).map((p) => p.trim()).filter(Boolean), ...(folders.length ? { folders } : {}) }),
      maxDeletions: limitValue(),
      // Lida da opção marcada (mesmo desativada: sem OneDrive/SharePoint, a escolha fica guardada).
      deleteMode: form.querySelector('[name="deleteMode"]:checked')?.value === 'trash' ? 'trash' : 'permanent',
    };
  };

  /** Limite de exclusões (null se inválido); em branco, o padrão. */
  const limitValue = () => {
    const raw = String(form.elements.maxDeletions.value || '').trim();
    if (raw === '') return DEFAULT_MAX_DELETIONS;
    const n = Number(raw);
    return Number.isInteger(n) && n >= 0 && n <= 10_000_000 ? n : null;
  };

  const sync = () => {
    const deletingNow = form.elements.action.value === 'delete';
    const retention = readRetention();
    const c = currentCriterion();
    const byDateNow = retention.mode === 'date';
    const byPeriodNow = retention.mode === 'period';
    form.querySelector('[data-delete-options]').hidden = !deletingNow;
    form.querySelectorAll('[data-cut]').forEach((el) => (el.hidden = el.dataset.cut !== retention.mode));
    form.querySelector('[data-since-word]').textContent = sinceWords(c);
    form.querySelectorAll('[data-period-label]').forEach((el) => {
      el.textContent = `${PERIOD_LABELS[c] || PERIOD_LABELS.used} ${el.dataset.periodLabel === 'from' ? 'a partir de' : 'até'}`;
    });
    form.querySelector('[data-since-label]').textContent = capitalize(sinceWords(c));
    if (!mail) {
      form.querySelector('[data-criterion-hint]').textContent = FILE_CRITERIA[c]?.hint || '';
      form.querySelector('[data-access-warning]').hidden = !ACCESS_CRITERIA.has(c);
      form.querySelector('[data-cloud-warning]').hidden = !(ACCESS_CRITERIA.has(c) && cloudChosen());
      // O campo de pastas de cada repositório marcado (o texto fica guardado se ele for desmarcado).
      const checked = new Set(checkedTargets().map((el) => el.value));
      form.querySelectorAll('[data-folders-for]').forEach((el) => (el.hidden = !checked.has(el.dataset.foldersFor)));
      form.querySelector('[data-folders-box]').hidden = checked.size === 0;
    }
    if (!mail) {
      // "Para a lixeira" só existe no OneDrive e no SharePoint: sem eles, a opção fica desativada (a
      // exclusão nas pastas do Windows é sempre definitiva), mas a escolha não muda — ela volta a valer
      // se um repositório da nuvem for marcado de novo.
      const trash = form.querySelector('[name="deleteMode"][value="trash"]');
      trash.disabled = !cloudChosen();
      form.querySelector('[data-trash-hint]').textContent = trash.disabled
        ? 'Só para o OneDrive e o SharePoint. Nos repositórios escolhidos (pastas do Windows) a exclusão é sempre definitiva: arquivos excluídos pela rede não vão para a Lixeira.'
        : 'Vale no OneDrive e no SharePoint (Lixeira do site). Nas pastas do Windows a exclusão é sempre definitiva: arquivos excluídos pela rede não vão para a Lixeira.';
    }
    if (retention.mode === 'age') form.elements.amount.max = String(UNITS[retention.unit]?.max || 100);
    const cutoff = cutoffPreview(retention);
    const limit = retention.maxDeletions;
    const limitText = !deletingNow ? '' : limit === null ? ' Informe um limite de exclusões válido (0 = sem limite).' : limit ? ` No máximo ${deletionsText(limit)} por execução.` : ' Sem limite de exclusões por execução.';
    const folderCount = retention.folders?.length || 0;
    const scoped = new Set((retention.folders || []).map((f) => f.repositoryId)).size;
    const others = checkedTargets().length > scoped ? '; nos demais repositórios marcados, em todas as pastas' : '';
    const where = mail || !folderCount ? '' : ` Somente ${folderCount === 1 ? 'na pasta informada' : `nas ${folderCount} pastas informadas`} (e subpastas)${others}.`;
    const content = cutoff
      ? html`<div class="preview-title">${icon('clock')} <b>${describeRetention({ ...retention, folders: retention.folders?.map((f) => ({ ...f, path: shortFolder(f) })) }, kind)}</b></div>
          <div class="muted small">${byDateNow || byPeriodNow ? 'Em cada execução' : 'Se executada hoje'}, a política ${deletingNow ? 'excluiria' : 'listaria'} os ${byPeriodNow ? periodPhrase(c, retention) : EXPIRES[c](cutoff.toLocaleDateString('pt-BR'))}.${where}${limitText}</div>
          ${byPeriodNow && !periodPreview(retention).to
            ? html`<div class="danger-text small">${icon('alert')} Sem a data "até", entram também os arquivos ${RECENT_VERB[c] || 'usados'} recentemente${deletingNow ? ' — e eles seriam excluídos' : ''}. Confira com "Simular agora".</div>`
            : ''}`
      : byPeriodNow
        ? html`<div class="danger-text small">${icon('alert')} ${PERIOD_ERROR}</div>`
        : byDateNow
          ? html`<div class="danger-text small">${icon('alert')} Informe a data de corte: de 02/01/1980 até hoje (datas futuras não valem).</div>`
          : html`<div class="danger-text small">${icon('alert')} Informe a idade máxima: de 1 a ${UNITS[retention.unit]?.max || 100} ${UNITS[retention.unit]?.many || 'anos'}.</div>`;
    // A prévia é uma região "aria-live": só é redesenhada quando o texto muda.
    if (content.value !== lastPreview) {
      lastPreview = content.value;
      paint(preview, content);
    }
    submit.className = `btn ${deletingNow ? 'danger' : 'primary'}`;
    paint(submit, html`${icon(deletingNow ? 'trash' : 'check')} ${deletingNow ? 'Salvar política com exclusão' : 'Salvar política'}`);
  };
  let lastPreview = '';

  const onChange = (event) => {
    if (event.target.name !== 'confirmDelete' && event.target.name !== 'name') sync();
  };
  form.addEventListener('change', onChange);
  form.addEventListener('input', onChange);
  // Data digitada pela metade não gera "input" nem "change": a prévia é refeita ao sair do campo.
  const onLeave = (event) => {
    if (event.target.type === 'date') sync();
  };
  form.addEventListener('focusout', onLeave);
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
      if (retention.mode === 'period') {
        // O campo com o problema: a data inicial inválida (ou vazia); senão a final (inválida ou antes da inicial).
        const fromBad = String(retention.from).trim() !== '' && !sinceInstant(retention.from);
        const toBad = String(retention.to).trim() !== '' && !sinceInstant(retention.to);
        form.elements[fromBad || (!toBad && !retention.from) ? 'periodFrom' : 'periodTo'].focus();
        return toast(PERIOD_ERROR, 'error');
      }
      if (retention.mode === 'date') {
        form.elements.since.focus();
        return toast('Informe a data de corte: de 02/01/1980 até hoje (datas futuras não valem).', 'error');
      }
      form.elements.amount.focus();
      return toast(`Informe a idade máxima: de 1 a ${UNITS[retention.unit].max} ${UNITS[retention.unit].many}.`, 'error');
    }
    if (!mail && ACCESS_CRITERIA.has(retention.criterion) && cloudChosen()) {
      return toast('O OneDrive e o SharePoint não informam o último acesso (abertura) de cada arquivo: escolha outro critério.', 'error');
    }
    if (retention.maxDeletions === null) {
      // Só vale com a exclusão (o campo fica escondido em "Somente listar"): o padrão é mantido.
      if (!deletingNow) retention.maxDeletions = policy?.retention?.maxDeletions ?? DEFAULT_MAX_DELETIONS;
      else {
        form.elements.maxDeletions.focus();
        return toast('Informe o limite de exclusões por execução: um número inteiro (0 = sem limite).', 'error');
      }
    }
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
    form.removeEventListener('focusout', onLeave);
  };
}
