// Agendamento nas telas de nova análise: quando executar (agora ou agendado), regra de
// recorrência com a prévia das próximas execuções, período de cada execução, horários perdidos e
// quantos relatórios guardar. Os horários valem no fuso do servidor.
import { post } from './api.js';
import { html, render as paint, icon, debounce, fmtServerDateTime } from './ui.js';

// Segunda a domingo: [valor (0 = domingo), abreviação, nome].
const WEEKDAYS = [
  [1, 'Seg', 'segunda-feira'],
  [2, 'Ter', 'terça-feira'],
  [3, 'Qua', 'quarta-feira'],
  [4, 'Qui', 'quinta-feira'],
  [5, 'Sex', 'sexta-feira'],
  [6, 'Sáb', 'sábado'],
  [0, 'Dom', 'domingo'],
];
const INTERVAL = {
  hourly: ['A cada quantas horas', 23],
  daily: ['A cada quantos dias', 365],
  weekly: ['A cada quantas semanas', 52],
  monthly: ['A cada quantos meses', 24],
};
// Semana do mês, concordando com o dia (segunda a sexta-feira: feminino; sábado e domingo: masculino).
const ORDINALS = [
  [1, 'primeir'],
  [2, 'segund'],
  [3, 'terceir'],
  [4, 'quart'],
  [-1, 'últim'],
];
const feminine = (weekday) => Number(weekday) >= 1 && Number(weekday) <= 5;
const ordinal = (stem, weekday) => `${stem}${feminine(weekday) ? 'a' : 'o'}`;

const KEEP = [
  [0, 'Todos'],
  [5, 'Os 5 mais recentes'],
  [10, 'Os 10 mais recentes'],
  [20, 'Os 20 mais recentes'],
  [50, 'Os 50 mais recentes'],
  [100, 'Os 100 mais recentes'],
];

/** Opções de "Relatórios guardados", incluindo um valor salvo fora da lista. */
function keepOptions(current) {
  const value = Number(current) || 0;
  if (KEEP.some(([v]) => v === value)) return KEEP;
  return [...KEEP, [value, `Os ${value} mais recentes`]].sort((a, b) => (a[0] || Infinity) - (b[0] || Infinity));
}

/** Hoje (AAAA-MM-DD) no fuso do servidor (ou no do navegador, se o fuso for desconhecido). */
function today(timeZone) {
  try {
    return new Intl.DateTimeFormat('en-CA', { timeZone: timeZone || undefined, year: 'numeric', month: '2-digit', day: '2-digit' }).format(new Date());
  } catch {
    return new Intl.DateTimeFormat('en-CA', { year: 'numeric', month: '2-digit', day: '2-digit' }).format(new Date());
  }
}

/** Nota sobre o fuso: os horários são os do servidor. */
export function zoneNote(info) {
  const server = info?.timeZone;
  let local = '';
  try {
    local = Intl.DateTimeFormat().resolvedOptions().timeZone;
  } catch {
    // navegador sem Intl completo
  }
  if (!server) return 'Horários do servidor do CLEAN.';
  return server === local ? `Horários do servidor do CLEAN (${server}).` : `Horários do servidor do CLEAN (${server}), diferente do fuso deste computador (${local}).`;
}

/** Opção "Relatórios guardados" (owner: "deste agendamento", "desta política"). */
export function keepField(schedule, owner = 'deste agendamento') {
  const sel = (a, b) => (String(a) === String(b) ? 'selected' : '');
  return html`<label class="field">
    <span>Relatórios guardados</span>
    <select name="keepLast">
      ${keepOptions(schedule?.keepLast ?? 0).map(([v, label]) => html`<option value="${v}" ${sel(schedule?.keepLast ?? 0, v)}>${label}</option>`)}
    </select>
    <small>Os mais antigos ${owner} são excluídos a cada nova execução (o registro geral de exclusões é mantido).</small>
  </label>`;
}

/**
 * Seção "Quando executar". kind: 'files' | 'mail'; schedule: agendamento em edição; fixed: a tela
 * é só de agendamento (sem a opção "Agora"). Políticas de retenção: manual (a primeira opção é
 * "Manualmente", sem regra), sem o período de cada execução (period: false) e com os relatórios
 * guardados fora da seção (keep: false).
 */
export function scheduleSection({ kind, schedule = null, fixed = false, info, manual = false, period: withPeriod = true, keep = true }) {
  const r = schedule?.rule || {};
  const p = schedule?.period || {};
  const frequency = r.frequency || 'daily';
  const weekdays = r.weekdays || (frequency === 'hourly' ? [1, 2, 3, 4, 5] : [1]);
  const mail = kind === 'mail';
  const items = mail ? 'mensagens' : 'arquivos';
  const sel = (a, b) => (String(a) === String(b) ? 'selected' : '');
  const chk = (a) => (a ? 'checked' : '');
  const later = fixed || (manual && Boolean(schedule?.rule));
  return html`<fieldset class="full" data-when-fieldset>
    <legend>Quando executar</legend>
    <div class="type-choice" ${fixed ? 'hidden' : ''}>
      <label>
        <input type="radio" name="when" value="${manual ? 'manual' : 'now'}" ${chk(!later)} />
        ${manual
          ? html`<span><b>Manualmente</b><small>Só quando alguém clicar em "Simular" ou "Executar" na lista de políticas.</small></span>`
          : html`<span><b>Agora</b><small>Uma análise, iniciada ao confirmar.</small></span>`}
      </label>
      <label>
        <input type="radio" name="when" value="schedule" ${chk(later)} />
        <span><b>Agendar</b><small>Uma vez ou com repetição, nos dias e horários escolhidos.</small></span>
      </label>
    </div>
    <div data-schedule-fields ${later ? '' : 'hidden'}>
      <div class="form-grid schedule-grid">
        <label class="field">
          <span>Repetição</span>
          <select name="frequency">
            <option value="once" ${sel(frequency, 'once')}>Uma vez</option>
            <option value="hourly" ${sel(frequency, 'hourly')}>A cada algumas horas</option>
            <option value="daily" ${sel(frequency, 'daily')}>Diariamente</option>
            <option value="weekly" ${sel(frequency, 'weekly')}>Semanalmente</option>
            <option value="monthly" ${sel(frequency, 'monthly')}>Mensalmente</option>
          </select>
        </label>
        <label class="field" data-for="hourly daily weekly monthly">
          <span data-interval-label>A cada quantos dias</span>
          <input type="number" name="interval" min="1" max="365" value="${r.interval || 1}" />
        </label>
        <label class="field">
          <span data-start-label>Data de início</span>
          <input type="date" name="startDate" value="${r.startDate || today(info?.timeZone)}" />
        </label>
        <label class="field">
          <span data-time-label>Horário</span>
          <input type="time" name="time" value="${r.time || '02:00'}" />
        </label>
        <label class="field" data-for="hourly">
          <span>Último horário do dia</span>
          <input type="time" name="untilTime" value="${r.untilTime || '18:00'}" />
        </label>
        <div class="field full" data-for="hourly weekly">
          <span class="field-label" id="weekdays-label">Dias da semana</span>
          <div class="weekdays" role="group" aria-labelledby="weekdays-label">
            ${WEEKDAYS.map(
              ([value, short, name]) => html`<label title="${name}">
                <input type="checkbox" name="weekdays" value="${value}" ${chk(weekdays.includes(value))} />
                <span aria-hidden="true">${short}</span><span class="sr-only">${name}</span>
              </label>`,
            )}
          </div>
        </div>
        <label class="check full" data-for="daily">
          <input type="checkbox" name="workdaysOnly" ${chk(r.workdaysOnly)} />
          <span>Somente em dias úteis (de segunda a sexta)</span>
        </label>
        <fieldset class="plain full" data-for="monthly">
          <legend>Dia do mês</legend>
          <div class="option-row">
            <label class="check"><input type="radio" name="monthlyMode" value="day" aria-label="No dia do mês informado ao lado" ${chk((r.monthlyMode || 'day') === 'day')} /><span>No dia</span></label>
            <input type="number" name="monthDay" min="1" max="31" value="${r.monthDay || 1}" aria-label="Dia do mês" data-selects="monthlyMode:day" />
          </div>
          <div class="option-row">
            <label class="check"><input type="radio" name="monthlyMode" value="last-day" ${chk(r.monthlyMode === 'last-day')} /><span>No último dia do mês</span></label>
          </div>
          <div class="option-row">
            <label class="check"><input type="radio" name="monthlyMode" value="weekday" aria-label="Num dia da semana do mês (semana e dia escolhidos ao lado)" ${chk(r.monthlyMode === 'weekday')} /><span data-article>${feminine(r.weekday ?? 1) ? 'Na' : 'No'}</span></label>
            <select name="weekOfMonth" aria-label="Qual semana do mês" data-selects="monthlyMode:weekday">
              ${ORDINALS.map(([v, stem]) => html`<option value="${v}" ${sel(r.weekOfMonth ?? 1, v)}>${ordinal(stem, r.weekday ?? 1)}</option>`)}
            </select>
            <select name="weekday" aria-label="Dia da semana" data-selects="monthlyMode:weekday">
              ${WEEKDAYS.map(([v, , name]) => html`<option value="${v}" ${sel(r.weekday ?? 1, v)}>${name}</option>`)}
            </select>
          </div>
        </fieldset>
        <fieldset class="plain full" data-for="hourly daily weekly monthly">
          <legend>Término</legend>
          <div class="option-row">
            <label class="check"><input type="radio" name="end" value="never" ${chk(!r.end || r.end === 'never')} /><span>Nunca</span></label>
          </div>
          <div class="option-row">
            <label class="check"><input type="radio" name="end" value="date" aria-label="Terminar na data informada ao lado" ${chk(r.end === 'date')} /><span>Em</span></label>
            <input type="date" name="endDate" value="${r.endDate || ''}" aria-label="Data de término" data-selects="end:date" />
          </div>
          <div class="option-row">
            <label class="check"><input type="radio" name="end" value="count" aria-label="Terminar depois do número de execuções informado ao lado" ${chk(r.end === 'count')} /><span>Depois de</span></label>
            <input type="number" name="count" min="1" max="1000" value="${r.count || 10}" aria-label="Número de execuções" data-selects="end:count" />
            <span>execuções</span>
          </div>
        </fieldset>
      </div>
      <div class="schedule-preview" data-preview aria-live="polite"></div>
      <p class="muted small">${zoneNote(info)} O CLEAN precisa estar em execução nesses horários (instale-o como serviço do Windows).</p>

      <div class="form-grid">
        ${withPeriod ? periodFields(p, mail, items) : ''}
        <label class="check full">
          <input type="checkbox" name="catchUp" ${chk(schedule ? schedule.catchUp : true)} />
          <span><b>Se o CLEAN estiver parado no horário, executar assim que ele voltar</b><br /><small class="muted">Uma única execução, mesmo que vários horários tenham sido perdidos. Desmarcado: o horário perdido fica só registrado no histórico.</small></span>
        </label>
        ${keep ? keepField(schedule) : ''}
      </div>
    </div>
  </fieldset>`;
}

/** Período de cada execução (todos, últimos dias ou incremental). */
function periodFields(p, mail, items) {
  const chk = (a) => (a ? 'checked' : '');
  return html`<fieldset class="plain full">
          <legend>${mail ? 'Mensagens analisadas' : 'Arquivos analisados'} em cada execução</legend>
          <div class="option-row">
            <label class="check"><input type="radio" name="periodType" value="all" ${chk(!p.type || p.type === 'all')} /><span>${mail ? 'Todas as mensagens' : 'Todos os arquivos'}</span></label>
          </div>
          <div class="option-row">
            <label class="check"><input type="radio" name="periodType" value="days" aria-label="Somente ${mail ? 'as recebidas' : 'os alterados'} nos últimos dias (número informado ao lado)" ${chk(p.type === 'days')} /><span>Somente ${mail ? 'as recebidas' : 'os alterados'} nos últimos</span></label>
            <input type="number" name="periodDays" min="1" max="3650" value="${p.days || 7}" aria-label="Número de dias" data-selects="periodType:days" />
            <span>dias</span>
          </div>
          <div class="option-row">
            <label class="check">
              <input type="radio" name="periodType" value="since-last" ${chk(p.type === 'since-last')} />
              <span>Somente ${mail ? 'as recebidas' : 'os alterados'} desde a execução anterior (incremental)</span>
            </label>
          </div>
          <div class="option-row indent" data-for-period="since-last">
            <span>Análise completa a cada</span>
            <input type="number" name="fullEvery" min="2" max="50" value="${p.fullEvery || 7}" aria-label="Análise completa a cada quantas execuções" />
            <span>execuções</span>
          </div>
          <small data-for-period="since-last">
            A partir da segunda execução, entram só ${mail ? 'as mensagens recebidas' : `os ${items} modificados, criados ou copiados para o repositório`} desde o início da última execução
            concluída (com 1 hora de margem); cada relatório mostra só o que foi encontrado nesse período. Mudanças nos locais, nas listas, nas
            opções ou na ação tornam a próxima execução completa. A análise completa periódica pega o que as incrementais não veem:
            ${mail ? 'mensagens movidas entre pastas ou importadas' : 'pastas movidas inteiras para o repositório'} e itens com erro de leitura.
          </small>
        </fieldset>`;
}

/** Regra, período e demais opções do agendamento, como a API espera. */
export function readSchedule(form) {
  const f = new FormData(form);
  const num = (name) => Number(f.get(name));
  const workdays = f.get('frequency') === 'daily' && f.get('workdaysOnly') === 'on';
  return {
    rule: {
      frequency: f.get('frequency'),
      interval: workdays ? 1 : Number(form.elements.interval.value),
      startDate: f.get('startDate'),
      time: f.get('time'),
      untilTime: f.get('untilTime'),
      weekdays: f.getAll('weekdays').map(Number),
      workdaysOnly: f.get('workdaysOnly') === 'on',
      monthlyMode: f.get('monthlyMode'),
      monthDay: num('monthDay'),
      weekOfMonth: num('weekOfMonth'),
      weekday: num('weekday'),
      end: f.get('end'),
      endDate: f.get('endDate'),
      count: num('count'),
    },
    period: { type: f.get('periodType'), days: num('periodDays'), fullEvery: f.get('fullEvery') === '' ? null : num('fullEvery') },
    catchUp: f.get('catchUp') === 'on',
    keepLast: num('keepLast'),
  };
}

/** A tela está no modo de agendamento? */
export const scheduling = (form) => form.elements.when?.value === 'schedule';

/**
 * Liga a seção ao formulário: mostra só os campos da repetição escolhida, atualiza a prévia e
 * avisa a tela quando o modo (agora/agendar) muda. Devolve a função de limpeza.
 */
export function bindSchedule(form, { onModeChange = () => {}, scheduleId = null } = {}) {
  const fields = form.querySelector('[data-schedule-fields]');
  const preview = form.querySelector('[data-preview]');
  let touchedDays = false;
  let lastFrequency = form.elements.frequency.value;
  let request = 0;
  const dayInputs = () => [...form.querySelectorAll('[name="weekdays"]')];
  // Dias escolhidos em cada repetição (hourly/weekly): voltar a uma repetição recupera os dias dela.
  const daysByFrequency = { [lastFrequency]: dayInputs().filter((el) => el.checked).map((el) => el.value) };
  const intervalByFrequency = { [lastFrequency]: form.elements.interval.value };

  const show = () => {
    const frequency = form.elements.frequency.value;
    fields.querySelectorAll('[data-for]').forEach((el) => {
      el.hidden = !el.dataset.for.split(' ').includes(frequency);
    });
    const period = form.elements.periodType?.value; // as políticas de retenção não têm período
    fields.querySelectorAll('[data-for-period]').forEach((el) => {
      el.hidden = el.dataset.forPeriod !== period;
    });
    const [label, max] = INTERVAL[frequency] || INTERVAL.daily;
    fields.querySelector('[data-interval-label]').textContent = label;
    form.elements.interval.max = String(max);
    const workdays = frequency === 'daily' && form.elements.workdaysOnly.checked;
    form.elements.interval.disabled = workdays;
    fields.querySelector('[data-start-label]').textContent = frequency === 'once' ? 'Data' : 'Data de início';
    fields.querySelector('[data-time-label]').textContent = frequency === 'hourly' ? 'Primeiro horário do dia' : 'Horário';
    // O intervalo muda de sentido com a repetição (horas, dias, semanas, meses): cada uma tem o seu
    // (1, até o usuário escolher).
    if (frequency !== lastFrequency) {
      intervalByFrequency[lastFrequency] = form.elements.interval.value;
      form.elements.interval.value = intervalByFrequency[frequency] || '1';
    }
    // Dias da semana: os já escolhidos nesta repetição ou, enquanto o usuário não escolher, uma
    // sugestão diferente para "a cada algumas horas" e "semanalmente".
    if (frequency !== lastFrequency) {
      daysByFrequency[lastFrequency] = dayInputs().filter((el) => el.checked).map((el) => el.value);
      const remembered = daysByFrequency[frequency];
      const days = remembered || (touchedDays ? null : frequency === 'hourly' ? ['1', '2', '3', '4', '5'] : ['1']);
      if (days) {
        for (const el of dayInputs()) el.checked = days.includes(el.value);
      }
    }
    lastFrequency = frequency;
    // "Na primeira segunda-feira" / "No último sábado".
    const weekday = form.elements.weekday.value;
    fields.querySelector('[data-article]').textContent = feminine(weekday) ? 'Na' : 'No';
    for (const option of form.elements.weekOfMonth.options) {
      const stem = ORDINALS.find(([v]) => String(v) === option.value)?.[1];
      if (stem) option.textContent = ordinal(stem, weekday);
    }
  };

  const drawPreview = async () => {
    if (!scheduling(form)) return;
    const token = ++request;
    try {
      const result = await post('/api/schedules/preview', { rule: readSchedule(form).rule, scheduleId });
      if (token !== request) return;
      const next = result.next || [];
      paint(
        preview,
        html`<div class="preview-title">${icon('clock')} <b>${result.description}</b></div>
          ${next.length
            ? html`<div class="muted small">${next.length === 1 ? 'Execução' : 'Próximas execuções'}:</div>
                <ol class="next-runs">${next.map((iso) => html`<li>${fmtServerDateTime(iso)}</li>`)}</ol>
                ${result.endsAt && result.rule.frequency !== 'once' ? html`<div class="muted small">Última execução${result.rule.end === 'count' ? ' prevista' : ''}: ${fmtServerDateTime(result.endsAt)}${result.rule.end === 'count' ? ` (${result.remaining === 1 ? 'falta 1 execução' : `faltam ${result.remaining} execuções`}; horários pulados ou perdidos não contam)` : ''}.</div>` : ''}`
            : html`<div class="danger-text small">${result.remaining === 0 ? 'Todas as execuções previstas já foram feitas: aumente o número de execuções ou mude a regra.' : 'Nenhuma execução futura: pela regra, a data e o horário já passaram.'}</div>`}`,
      );
    } catch (err) {
      if (token !== request) return;
      paint(preview, html`<div class="danger-text small">${icon('alert')} ${err.message}</div>`);
    }
  };
  const refresh = debounce(drawPreview, 250);

  const onChange = (event) => {
    const { name } = event.target;
    if (name === 'weekdays') touchedDays = true;
    // Campo ao lado de uma opção (ex.: o dia do mês) seleciona a opção.
    const selects = event.target.dataset?.selects;
    if (selects) {
      const [group, value] = selects.split(':');
      const radio = form.querySelector(`[name="${group}"][value="${value}"]`);
      if (radio) radio.checked = true;
    }
    if (name === 'when') {
      fields.hidden = !scheduling(form);
      onModeChange(scheduling(form));
    }
    show();
    if (fields.contains(event.target) || name === 'when') refresh();
  };
  form.addEventListener('change', onChange);
  form.addEventListener('input', onChange);
  show();
  onModeChange(scheduling(form));
  drawPreview();
  return () => {
    refresh.cancel();
    request++;
    form.removeEventListener('change', onChange);
    form.removeEventListener('input', onChange);
  };
}
