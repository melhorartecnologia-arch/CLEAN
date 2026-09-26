// Regras de recorrência dos agendamentos: validação, próximas execuções e descrição em português.
// As datas e os horários valem no fuso do servidor (a hora local da máquina que executa o CLEAN).
//
// Regra (depois de validada):
//   { frequency: 'once' | 'hourly' | 'daily' | 'weekly' | 'monthly',
//     startDate: 'AAAA-MM-DD', time: 'HH:MM',
//     interval,                         // a cada N horas, dias, semanas ou meses
//     untilTime: 'HH:MM',               // hourly: fim da janela do dia (o início é `time`)
//     weekdays: [0..6],                 // hourly e weekly (0 = domingo)
//     workdaysOnly,                     // daily: só de segunda a sexta
//     monthlyMode: 'day' | 'last-day' | 'weekday', monthDay, weekOfMonth (1..4 ou -1), weekday,
//     end: 'never' | 'date' | 'count', endDate, count }
//
// "Depois de N execuções" conta as execuções de fato iniciadas pelo agendamento (quem executa a
// regra informa quantas ainda faltam em `remaining`); horários pulados ou perdidos não contam.

export const FREQUENCIES = ['once', 'hourly', 'daily', 'weekly', 'monthly'];
const MAX_INTERVAL = { hourly: 23, daily: 365, weekly: 52, monthly: 24 };
export const MAX_COUNT = 1000;
const DAY = 86400000;
// Maior distância possível entre duas execuções de uma regra válida (a cada 24 meses): acima
// disso a busca para (proteção contra laços).
const MAX_GAP_DAYS = 800;

export class RuleError extends Error {
  constructor(message) {
    super(message);
    this.status = 400;
  }
}

// ---------- datas do calendário (dias numerados, sem fuso) ----------

const dayNumber = (y, m, d) => Math.round(Date.UTC(y, m, d) / DAY);

function fromDay(n) {
  const date = new Date(n * DAY);
  return { y: date.getUTCFullYear(), m: date.getUTCMonth(), d: date.getUTCDate(), wd: date.getUTCDay() };
}

const weekdayOf = (n) => fromDay(n).wd;
const mondayOf = (n) => n - ((weekdayOf(n) + 6) % 7);
const daysInMonth = (y, m) => new Date(Date.UTC(y, m + 1, 0)).getUTCDate();
const isWorkday = (n) => weekdayOf(n) >= 1 && weekdayOf(n) <= 5;

/** Dia (do calendário local) de um instante. */
const localDay = (date) => dayNumber(date.getFullYear(), date.getMonth(), date.getDate());

/** Instante do dia `n` no horário `minutes` (hora local do servidor). */
function localAt(n, minutes) {
  const { y, m, d } = fromDay(n);
  return new Date(y, m, d, Math.floor(minutes / 60), minutes % 60, 0, 0);
}

const pad = (n) => String(n).padStart(2, '0');

/** 'AAAA-MM-DD' de um instante, no calendário local. */
export function localDateText(date) {
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;
}

function parseDateText(value) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(value ?? '').trim());
  if (!m) return null;
  const [y, mo, d] = [Number(m[1]), Number(m[2]) - 1, Number(m[3])];
  if (y < 2000 || y > 2099 || mo > 11 || d < 1 || d > daysInMonth(y, mo)) return null;
  return dayNumber(y, mo, d);
}

const dayOfText = (value) => parseDateText(value);
const minutesOf = (value) => {
  const [h, m] = String(value).split(':').map(Number);
  return h * 60 + m;
};

// ---------- validação ----------

function dateField(value, field) {
  const text = String(value ?? '').trim();
  if (!text) throw new RuleError(`Informe ${field}.`);
  if (parseDateText(text) === null) throw new RuleError(`${field[0].toUpperCase()}${field.slice(1)} é inválida (use uma data entre 2000 e 2099).`);
  return text;
}

function timeField(value, field) {
  const m = /^(\d{1,2}):(\d{2})$/.exec(String(value ?? '').trim());
  if (!m || Number(m[1]) > 23 || Number(m[2]) > 59) throw new RuleError(`Informe ${field} no formato HH:MM (ex.: 08:30).`);
  return `${pad(Number(m[1]))}:${m[2]}`;
}

/** Inteiro de um número ou de um texto com número; vazio, nulo, listas e frações não valem. */
function toInt(value) {
  if (typeof value === 'number') return Number.isInteger(value) ? value : NaN;
  if (typeof value === 'string' && /^\s*-?\d+\s*$/.test(value)) return Number(value);
  return NaN;
}

function intField(value, min, max, message) {
  const n = toInt(value);
  if (!Number.isInteger(n) || n < min || n > max) throw new RuleError(message);
  return n;
}

function weekdaysField(value) {
  const days = (Array.isArray(value) ? value : []).map(toInt);
  if (days.some((d) => !Number.isInteger(d) || d < 0 || d > 6)) throw new RuleError('Dia da semana inválido.');
  if (days.length === 0) throw new RuleError('Escolha ao menos um dia da semana.');
  return [...new Set(days)].sort((a, b) => a - b);
}

/** Valor de uma lista de opções; ausente usa o padrão, desconhecido é recusado. */
function choiceField(value, choices, fallback, message) {
  if (value === undefined || value === null || value === '') return fallback;
  if (!choices.includes(value)) throw new RuleError(message);
  return value;
}

/** Valida e normaliza uma regra recebida da interface (lança RuleError com a explicação). */
export function validateRule(input) {
  if (!input || typeof input !== 'object' || Array.isArray(input)) throw new RuleError('Informe quando o agendamento deve ser executado.');
  const { frequency } = input;
  if (!FREQUENCIES.includes(frequency)) throw new RuleError('Escolha a repetição: uma vez, a cada algumas horas, diária, semanal ou mensal.');
  const rule = {
    frequency,
    startDate: dateField(input.startDate, frequency === 'once' ? 'a data da execução' : 'a data de início'),
    time: timeField(input.time, frequency === 'hourly' ? 'o horário inicial' : 'o horário'),
  };
  if (frequency === 'once') return rule;

  const max = MAX_INTERVAL[frequency];
  const unit = { hourly: 'horas', daily: 'dias', weekly: 'semanas', monthly: 'meses' }[frequency];
  const workdays = frequency === 'daily' && input.workdaysOnly === true;
  // "Somente em dias úteis" é sempre a cada 1 dia (o intervalo informado é ignorado).
  rule.interval = workdays ? 1 : intField(input.interval ?? 1, 1, max, `O intervalo deve ser de 1 a ${max} ${unit}.`);

  if (frequency === 'hourly') {
    rule.untilTime = timeField(input.untilTime || '23:59', 'o horário final');
    if (minutesOf(rule.untilTime) < minutesOf(rule.time)) {
      throw new RuleError('O horário final deve ser igual ou posterior ao inicial (a janela de execução não pode passar da meia-noite).');
    }
    rule.weekdays = weekdaysField(input.weekdays);
  }
  if (frequency === 'daily') rule.workdaysOnly = workdays;
  if (frequency === 'weekly') rule.weekdays = weekdaysField(input.weekdays);
  if (frequency === 'monthly') {
    rule.monthlyMode = choiceField(input.monthlyMode, ['day', 'last-day', 'weekday'], 'day', 'Escolha o dia do mês: um dia, o último dia ou um dia da semana.');
    if (rule.monthlyMode === 'day') rule.monthDay = intField(input.monthDay, 1, 31, 'O dia do mês deve ser de 1 a 31.');
    if (rule.monthlyMode === 'weekday') {
      rule.weekOfMonth = intField(input.weekOfMonth ?? 1, -1, 4, 'Escolha a semana do mês: da primeira à quarta, ou a última.');
      if (rule.weekOfMonth === 0) throw new RuleError('Escolha a semana do mês: da primeira à quarta, ou a última.');
      rule.weekday = intField(input.weekday, 0, 6, 'Escolha o dia da semana.');
    }
  }

  rule.end = choiceField(input.end, ['never', 'date', 'count'], 'never', 'Escolha o término: nunca, numa data ou depois de um número de execuções.');
  if (rule.end === 'date') {
    rule.endDate = dateField(input.endDate, 'a data de término');
    if (dayOfText(rule.endDate) < dayOfText(rule.startDate)) throw new RuleError('A data de término deve ser igual ou posterior à data de início.');
  }
  if (rule.end === 'count') rule.count = intField(input.count, 1, MAX_COUNT, `O número de execuções deve ser de 1 a ${MAX_COUNT}.`);
  return rule;
}

// ---------- próximas execuções ----------

/** Dia do mês em que a regra mensal cai (dia 31 em mês de 30 dias: o último dia). */
function monthlyDay(rule, y, m) {
  const last = daysInMonth(y, m);
  if (rule.monthlyMode === 'last-day') return last;
  if (rule.monthlyMode === 'weekday') {
    const first = 1 + ((rule.weekday - fromDay(dayNumber(y, m, 1)).wd + 7) % 7);
    if (rule.weekOfMonth !== -1) return first + 7 * (rule.weekOfMonth - 1);
    return first + 7 * Math.floor((last - first) / 7);
  }
  return Math.min(rule.monthDay, last);
}

/** Dias (em ordem) em que a regra se aplica, a partir do dia `from`. */
function* matchingDays(rule, from) {
  const start = dayOfText(rule.startDate);
  let n = Math.max(from, start);
  let guard = n + MAX_GAP_DAYS;
  const found = (day) => {
    guard = day + MAX_GAP_DAYS;
    return day;
  };
  switch (rule.frequency) {
    case 'once':
      if (n === start) yield start;
      return;
    case 'daily':
      if (rule.workdaysOnly) {
        for (; n <= guard; n++) if (isWorkday(n)) yield found(n);
        return;
      }
      for (n = start + Math.ceil((n - start) / rule.interval) * rule.interval; n <= guard; n += rule.interval) yield found(n);
      return;
    case 'hourly':
      for (; n <= guard; n++) if (rule.weekdays.includes(weekdayOf(n))) yield found(n);
      return;
    case 'weekly': {
      const base = mondayOf(start);
      while (n <= guard) {
        const week = Math.floor((mondayOf(n) - base) / 7);
        const skip = week % rule.interval;
        if (skip !== 0) {
          n = mondayOf(n) + 7 * (rule.interval - skip); // segunda-feira da próxima semana da regra
          continue;
        }
        if (rule.weekdays.includes(weekdayOf(n))) yield found(n);
        n++;
      }
      return;
    }
    case 'monthly': {
      const s = fromDay(start);
      const f = fromDay(n);
      let k = Math.ceil(((f.y - s.y) * 12 + (f.m - s.m)) / rule.interval) * rule.interval;
      for (;; k += rule.interval) {
        const y = s.y + Math.floor((s.m + k) / 12);
        const m = (s.m + k) % 12;
        const day = dayNumber(y, m, monthlyDay(rule, y, m));
        if (day > guard) return;
        if (day >= n) yield found(day);
      }
    }
    default:
  }
}

/** Horários (minutos do dia) em que a regra executa em cada dia. */
function timesOf(rule) {
  const first = minutesOf(rule.time);
  if (rule.frequency !== 'hourly') return [first];
  const last = minutesOf(rule.untilTime);
  const out = [];
  for (let t = first; t <= last; t += rule.interval * 60) out.push(t);
  return out;
}

/** Instantes das execuções, em ordem crescente e sem repetição, a partir do dia `from`. */
function* instants(rule, from) {
  const times = timesOf(rule);
  let last = -Infinity;
  for (const n of matchingDays(rule, from)) {
    for (const t of times) {
      const at = localAt(n, t);
      // Na mudança para o horário de verão, um horário inexistente vira o seguinte: sem repetir.
      if (+at <= last) continue;
      last = +at;
      yield at;
    }
  }
}

/** Fim do último dia permitido (término por data) ou null. */
function dateEnd(rule) {
  if (rule.frequency === 'once' || rule.end !== 'date') return null;
  const { y, m, d } = fromDay(dayOfText(rule.endDate));
  return new Date(y, m, d, 23, 59, 59, 999);
}

/**
 * As próximas `n` execuções depois de `after` (em ordem). remaining: quantas execuções ainda
 * faltam no término por número de execuções (sem ele, todas as da regra).
 */
export function upcoming(rule, after = new Date(), n = 5, { remaining = Infinity } = {}) {
  const out = [];
  const limit = Math.min(n, rule.end === 'count' && rule.frequency !== 'once' ? remaining : Infinity);
  if (!(limit > 0)) return out;
  const end = dateEnd(rule);
  for (const at of instants(rule, localDay(new Date(after)))) {
    if (+at <= +after) continue;
    if (end && +at > +end) break;
    out.push(at);
    if (out.length >= limit) break;
  }
  return out;
}

/** A próxima execução depois de `after`, ou null quando a regra já terminou. */
export function nextOccurrence(rule, after = new Date(), options = {}) {
  return upcoming(rule, after, 1, options)[0] || null;
}

/**
 * Última execução prevista depois de `after`: a do último dia (término por data) ou a última das
 * que faltam (término por número de execuções, sem contar horários pulados). null quando a regra
 * não termina ou já terminou.
 */
export function lastOccurrence(rule, { after = new Date(), remaining = Infinity } = {}) {
  if (rule.frequency === 'once') return nextOccurrence(rule, after);
  if (rule.end === 'count') return upcoming(rule, after, Math.min(rule.count, remaining)).at(-1) || null;
  const end = dateEnd(rule);
  if (!end || +end <= +after) return null;
  // Entre duas execuções há no máximo MAX_GAP_DAYS dias: a última fica nesse trecho antes do fim.
  let last = null;
  for (const at of instants(rule, Math.max(localDay(new Date(after)), localDay(end) - MAX_GAP_DAYS))) {
    if (+at > +end) break;
    if (+at > +after) last = at;
  }
  return last;
}

/** Quantos horários da regra houve no intervalo (from, to] (até `cap`). */
export function countBetween(rule, from, to, cap = 1000) {
  let count = 0;
  const end = dateEnd(rule);
  for (const at of instants(rule, localDay(new Date(from)))) {
    if (+at <= +from) continue;
    if (+at > +to || (end && +at > +end)) break;
    if (++count >= cap) break;
  }
  return count;
}

// ---------- descrição ----------

const WEEKDAYS = ['domingo', 'segunda-feira', 'terça-feira', 'quarta-feira', 'quinta-feira', 'sexta-feira', 'sábado'];
const SHORT = ['domingo', 'segunda', 'terça', 'quarta', 'quinta', 'sexta', 'sábado'];
const PLURAL = ['domingos', 'segundas', 'terças', 'quartas', 'quintas', 'sextas', 'sábados'];
const ORDINALS = { 1: 'primeir', 2: 'segund', 3: 'terceir', 4: 'quart', '-1': 'últim' };
const MONDAY_FIRST = [1, 2, 3, 4, 5, 6, 0];
const feminine = (wd) => wd >= 1 && wd <= 5; // "segunda-feira" ... "sexta-feira"

function joinPt(items) {
  return items.length <= 1 ? items.join('') : `${items.slice(0, -1).join(', ')} e ${items.at(-1)}`;
}

/** 'AAAA-MM-DD' → 'DD/MM/AAAA'. */
export function dateBr(text) {
  const [y, m, d] = String(text).split('-');
  return `${d}/${m}/${y}`;
}

const sameDays = (days, list) => days.length === list.length && list.every((d) => days.includes(d));

/** "às segundas e sextas e aos sábados", "de segunda a sexta", "todos os dias"... */
function daysPhrase(days) {
  if (days.length === 7) return 'todos os dias';
  if (sameDays(days, [1, 2, 3, 4, 5])) return 'de segunda a sexta';
  const ordered = MONDAY_FIRST.filter((d) => days.includes(d));
  const fem = ordered.filter(feminine);
  const masc = ordered.filter((d) => !feminine(d));
  const parts = [];
  if (fem.length) parts.push(`às ${fem.length === 1 ? `${PLURAL[fem[0]]}-feiras` : joinPt(fem.map((d) => PLURAL[d]))}`);
  if (masc.length) parts.push(`aos ${joinPt(masc.map((d) => PLURAL[d]))}`);
  return parts.join(' e ');
}

const capital = (text) => text[0].toUpperCase() + text.slice(1);

function monthlyText(rule) {
  if (rule.monthlyMode === 'last-day') return 'no último dia';
  if (rule.monthlyMode === 'weekday') {
    const fem = feminine(rule.weekday);
    return `${fem ? 'na' : 'no'} ${ORDINALS[rule.weekOfMonth]}${fem ? 'a' : 'o'} ${WEEKDAYS[rule.weekday]}`;
  }
  const note = rule.monthDay > 28 ? ` (nos meses mais curtos, no último dia)` : '';
  return `no dia ${rule.monthDay}${note}`;
}

/**
 * Descrição da regra em português, ex.: "Toda semana às segundas e quintas, às 22:00, até
 * 31/12/2026". `today` ('AAAA-MM-DD') omite "a partir de" quando o início já passou.
 */
export function describeRule(rule, { today = localDateText(new Date()) } = {}) {
  const at = `às ${rule.time}`;
  let text;
  switch (rule.frequency) {
    case 'once':
      return `Uma vez, em ${dateBr(rule.startDate)} ${at}`;
    case 'hourly': {
      const every = rule.interval === 1 ? 'A cada hora' : `A cada ${rule.interval} horas`;
      const window = rule.time === '00:00' && rule.untilTime === '23:59' ? ', o dia todo' : `, das ${rule.time} às ${rule.untilTime}`;
      text = `${every}${window}${rule.weekdays.length === 7 ? '' : `, ${daysPhrase(rule.weekdays)}`}`;
      break;
    }
    case 'daily':
      if (rule.workdaysOnly) text = `De segunda a sexta, ${at}`;
      else text = rule.interval === 1 ? `Todos os dias, ${at}` : `A cada ${rule.interval} dias, ${at}`;
      break;
    case 'weekly': {
      const days = rule.weekdays;
      if (rule.interval === 1 && days.length === 1) {
        const d = days[0];
        text = `${feminine(d) ? 'Toda' : 'Todo'} ${WEEKDAYS[d]}, ${at}`;
      } else if (rule.interval === 1) {
        text = `${capital(daysPhrase(days))}, ${at}`;
      } else {
        const names = joinPt(MONDAY_FIRST.filter((d) => days.includes(d)).map((d) => SHORT[d]));
        text = `A cada ${rule.interval} semanas (${names}), ${at}`;
      }
      break;
    }
    case 'monthly':
      text = `${rule.interval === 1 ? 'Todo mês' : `A cada ${rule.interval} meses`}, ${monthlyText(rule)}, ${at}`;
      break;
    default:
      return '';
  }
  if (rule.startDate > today) text += `, a partir de ${dateBr(rule.startDate)}`;
  if (rule.end === 'date') text += `, até ${dateBr(rule.endDate)}`;
  if (rule.end === 'count') text += rule.count === 1 ? ', uma única vez' : `, ${rule.count} execuções no total`;
  return text;
}
