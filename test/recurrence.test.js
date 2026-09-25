import { test } from 'node:test';
import assert from 'node:assert/strict';
import { validateRule, nextOccurrence, upcoming, endOf, countBetween, describeRule, dateBr, RuleError } from '../src/schedule/recurrence.js';

// As regras valem na hora local: as datas esperadas também são montadas na hora local.
const at = (y, m, d, h = 0, min = 0) => new Date(y, m - 1, d, h, min, 0, 0);
const rule = (input) => validateRule(input);
const list = (r, after, n) => upcoming(r, after, n).map((d) => d.getTime());
const times = (...dates) => dates.map((d) => d.getTime());

test('validação das regras', () => {
  const fails = (input, pattern) => assert.throws(() => validateRule(input), (err) => err instanceof RuleError && pattern.test(err.message));
  fails(null, /Informe quando/);
  fails({ frequency: 'yearly' }, /Escolha a repetição/);
  fails({ frequency: 'daily', startDate: '2026-02-30', time: '08:00' }, /data de início é inválida/);
  fails({ frequency: 'daily', startDate: '1999-12-31', time: '08:00' }, /inválida/);
  fails({ frequency: 'daily', startDate: '2026-09-25', time: '24:00' }, /HH:MM/);
  fails({ frequency: 'daily', startDate: '2026-09-25', time: '08:00', interval: 0 }, /de 1 a 365 dias/);
  fails({ frequency: 'hourly', startDate: '2026-09-25', time: '08:00', interval: 24, weekdays: [1] }, /de 1 a 23 horas/);
  fails({ frequency: 'hourly', startDate: '2026-09-25', time: '18:00', untilTime: '08:00', weekdays: [1] }, /não pode passar da meia-noite/);
  fails({ frequency: 'weekly', startDate: '2026-09-25', time: '08:00', weekdays: [] }, /ao menos um dia/);
  fails({ frequency: 'monthly', startDate: '2026-09-25', time: '08:00', monthDay: 32 }, /de 1 a 31/);
  fails({ frequency: 'daily', startDate: '2026-09-25', time: '08:00', end: 'date', endDate: '2026-09-24' }, /igual ou posterior/);
  fails({ frequency: 'daily', startDate: '2026-09-25', time: '08:00', end: 'count', count: 1001 }, /de 1 a 1000/);

  // Normalização: hora com um dígito, dias repetidos e em ordem, dias úteis sempre a cada 1 dia.
  assert.deepEqual(rule({ frequency: 'weekly', startDate: '2026-09-25', time: '8:05', weekdays: [5, 1, 1, '3', 9] }), {
    frequency: 'weekly',
    startDate: '2026-09-25',
    time: '08:05',
    interval: 1,
    weekdays: [1, 3, 5],
    end: 'never',
  });
  assert.equal(rule({ frequency: 'daily', startDate: '2026-09-25', time: '08:00', interval: 3, workdaysOnly: true }).interval, 1);
  assert.deepEqual(rule({ frequency: 'once', startDate: '2026-09-25', time: '08:00', interval: 5, end: 'count', count: 2 }), {
    frequency: 'once',
    startDate: '2026-09-25',
    time: '08:00',
  });
});

test('uma vez', () => {
  const r = rule({ frequency: 'once', startDate: '2026-10-01', time: '14:30' });
  assert.equal(+nextOccurrence(r, at(2026, 9, 25, 10)), +at(2026, 10, 1, 14, 30));
  assert.equal(+nextOccurrence(r, at(2026, 10, 1, 14, 29)), +at(2026, 10, 1, 14, 30));
  assert.equal(nextOccurrence(r, at(2026, 10, 1, 14, 30)), null, 'o horário exato já conta como passado');
  assert.equal(upcoming(r, at(2026, 9, 1), 5).length, 1);
  assert.equal(endOf(r), null);
});

test('a cada N horas, numa janela do dia e em dias escolhidos', () => {
  const r = rule({ frequency: 'hourly', interval: 2, startDate: '2026-09-25', time: '08:00', untilTime: '18:00', weekdays: [1, 2, 3, 4, 5] });
  // Sexta-feira, 25/09/2026, 17h: a próxima é às 18h; depois, só na segunda às 8h.
  assert.equal(+nextOccurrence(r, at(2026, 9, 25, 17)), +at(2026, 9, 25, 18));
  assert.deepEqual(list(r, at(2026, 9, 25, 18), 3), times(at(2026, 9, 28, 8), at(2026, 9, 28, 10), at(2026, 9, 28, 12)));
  // Antes do início: começa na data de início.
  assert.equal(+nextOccurrence(r, at(2026, 1, 1)), +at(2026, 9, 25, 8));
  // A cada 5 horas no dia todo: 0h, 5h, 10h, 15h, 20h e recomeça às 0h do dia seguinte.
  const five = rule({ frequency: 'hourly', interval: 5, startDate: '2026-09-25', time: '00:00', untilTime: '23:59', weekdays: [0, 1, 2, 3, 4, 5, 6] });
  assert.deepEqual(list(five, at(2026, 9, 25, 14), 4), times(at(2026, 9, 25, 15), at(2026, 9, 25, 20), at(2026, 9, 26, 0), at(2026, 9, 26, 5)));
  assert.equal(countBetween(five, at(2026, 9, 25, 0), at(2026, 9, 26, 0)), 5, 'de (0h, 0h do dia seguinte]: 5h, 10h, 15h, 20h e 0h');
});

test('diária: a cada N dias e só em dias úteis', () => {
  const every3 = rule({ frequency: 'daily', interval: 3, startDate: '2026-09-25', time: '02:00' });
  assert.deepEqual(list(every3, at(2026, 9, 24), 3), times(at(2026, 9, 25, 2), at(2026, 9, 28, 2), at(2026, 10, 1, 2)));
  assert.deepEqual(list(every3, at(2026, 9, 26), 2), times(at(2026, 9, 28, 2), at(2026, 10, 1, 2)), 'o ciclo conta a partir do início');
  const workdays = rule({ frequency: 'daily', startDate: '2026-09-25', time: '07:00', workdaysOnly: true });
  assert.deepEqual(list(workdays, at(2026, 9, 25, 8), 2), times(at(2026, 9, 28, 7), at(2026, 9, 29, 7)), 'pula o sábado e o domingo');
  // Virada do ano e do mês de fevereiro.
  const daily = rule({ frequency: 'daily', startDate: '2026-12-30', time: '23:59' });
  assert.deepEqual(list(daily, at(2026, 12, 31, 0), 2), times(at(2026, 12, 31, 23, 59), at(2027, 1, 1, 23, 59)));
});

test('semanal: dias escolhidos e a cada N semanas (semanas de segunda a domingo)', () => {
  const r = rule({ frequency: 'weekly', interval: 2, startDate: '2026-09-23', time: '22:00', weekdays: [1, 4] });
  // Início numa quarta: a segunda da mesma semana já passou; depois, semanas alternadas.
  assert.deepEqual(
    list(r, at(2026, 9, 1), 5),
    times(at(2026, 9, 24, 22), at(2026, 10, 5, 22), at(2026, 10, 8, 22), at(2026, 10, 19, 22), at(2026, 10, 22, 22)),
  );
  const weekend = rule({ frequency: 'weekly', startDate: '2026-09-25', time: '03:00', weekdays: [0, 6] });
  assert.deepEqual(list(weekend, at(2026, 9, 25), 3), times(at(2026, 9, 26, 3), at(2026, 9, 27, 3), at(2026, 10, 3, 3)));
});

test('mensal: dia do mês (com meses curtos), último dia e dia da semana', () => {
  const d31 = rule({ frequency: 'monthly', startDate: '2027-11-01', time: '03:00', monthlyMode: 'day', monthDay: 31 });
  assert.deepEqual(
    list(d31, at(2027, 10, 1), 5),
    times(at(2027, 11, 30, 3), at(2027, 12, 31, 3), at(2028, 1, 31, 3), at(2028, 2, 29, 3), at(2028, 3, 31, 3)),
    'dia 31 vira o último dia nos meses mais curtos (29/02 no ano bissexto)',
  );
  const last = rule({ frequency: 'monthly', startDate: '2026-01-15', time: '18:00', monthlyMode: 'last-day' });
  assert.deepEqual(list(last, at(2026, 1, 1), 3), times(at(2026, 1, 31, 18), at(2026, 2, 28, 18), at(2026, 3, 31, 18)));
  const firstMonday = rule({ frequency: 'monthly', startDate: '2026-09-01', time: '09:00', monthlyMode: 'weekday', weekOfMonth: 1, weekday: 1 });
  assert.deepEqual(list(firstMonday, at(2026, 9, 1), 3), times(at(2026, 9, 7, 9), at(2026, 10, 5, 9), at(2026, 11, 2, 9)));
  const lastFriday = rule({ frequency: 'monthly', startDate: '2026-09-01', time: '17:00', monthlyMode: 'weekday', weekOfMonth: -1, weekday: 5 });
  assert.deepEqual(list(lastFriday, at(2026, 9, 1), 3), times(at(2026, 9, 25, 17), at(2026, 10, 30, 17), at(2026, 11, 27, 17)));
  const fourthThursday = rule({ frequency: 'monthly', startDate: '2026-11-01', time: '10:00', monthlyMode: 'weekday', weekOfMonth: 4, weekday: 4 });
  assert.equal(+nextOccurrence(fourthThursday, at(2026, 11, 1)), +at(2026, 11, 26, 10));
  // A cada 3 meses a partir de setembro: o dia 10 de setembro já passou, então dezembro e março.
  const quarterly = rule({ frequency: 'monthly', interval: 3, startDate: '2026-09-25', time: '06:00', monthlyMode: 'day', monthDay: 10 });
  assert.deepEqual(list(quarterly, at(2026, 9, 25), 2), times(at(2026, 12, 10, 6), at(2027, 3, 10, 6)));
  // A cada 24 meses (o maior intervalo): a busca alcança a próxima.
  const biennial = rule({ frequency: 'monthly', interval: 24, startDate: '2026-09-25', time: '06:00', monthlyMode: 'day', monthDay: 1 });
  assert.equal(+nextOccurrence(biennial, at(2026, 9, 25)), +at(2028, 9, 1, 6));
});

test('término por data e por número de execuções', () => {
  const byDate = rule({ frequency: 'daily', startDate: '2026-09-25', time: '02:00', end: 'date', endDate: '2026-09-27' });
  assert.deepEqual(list(byDate, at(2026, 9, 1), 5), times(at(2026, 9, 25, 2), at(2026, 9, 26, 2), at(2026, 9, 27, 2)), 'o último dia entra');
  assert.equal(nextOccurrence(byDate, at(2026, 9, 27, 2)), null);
  const byCount = rule({ frequency: 'weekly', startDate: '2026-09-25', time: '02:00', weekdays: [1, 5], end: 'count', count: 3 });
  assert.equal(+endOf(byCount), +at(2026, 10, 2, 2), 'as execuções contam a partir do início, não de agora');
  assert.deepEqual(list(byCount, at(2026, 9, 28, 12), 5), times(at(2026, 10, 2, 2)));
  assert.equal(nextOccurrence(byCount, at(2026, 10, 2, 2)), null);
  assert.equal(countBetween(byCount, at(2026, 9, 1), at(2027, 1, 1)), 3);
});

test('descrição em português', () => {
  const today = '2026-09-25';
  const d = (input) => describeRule(rule(input), { today });
  assert.equal(d({ frequency: 'once', startDate: '2026-10-01', time: '14:30' }), 'Uma vez, em 01/10/2026 às 14:30');
  assert.equal(d({ frequency: 'hourly', startDate: today, time: '08:00', untilTime: '18:00', weekdays: [1, 2, 3, 4, 5] }), 'A cada hora, das 08:00 às 18:00, de segunda a sexta');
  assert.equal(d({ frequency: 'hourly', interval: 4, startDate: today, time: '00:00', untilTime: '23:59', weekdays: [0, 1, 2, 3, 4, 5, 6] }), 'A cada 4 horas, o dia todo');
  assert.equal(d({ frequency: 'daily', startDate: today, time: '02:00' }), 'Todos os dias, às 02:00');
  assert.equal(d({ frequency: 'daily', interval: 3, startDate: today, time: '02:00' }), 'A cada 3 dias, às 02:00');
  assert.equal(d({ frequency: 'daily', workdaysOnly: true, startDate: today, time: '07:00' }), 'De segunda a sexta, às 07:00');
  assert.equal(d({ frequency: 'weekly', startDate: today, time: '22:00', weekdays: [1] }), 'Toda segunda-feira, às 22:00');
  assert.equal(d({ frequency: 'weekly', startDate: today, time: '22:00', weekdays: [6] }), 'Todo sábado, às 22:00');
  assert.equal(d({ frequency: 'weekly', startDate: today, time: '22:00', weekdays: [5, 1, 3] }), 'Às segundas, quartas e sextas, às 22:00');
  assert.equal(d({ frequency: 'weekly', startDate: today, time: '22:00', weekdays: [0, 6] }), 'Aos sábados e domingos, às 22:00');
  assert.equal(d({ frequency: 'weekly', startDate: today, time: '22:00', weekdays: [1, 6] }), 'Às segundas-feiras e aos sábados, às 22:00');
  assert.equal(d({ frequency: 'weekly', interval: 2, startDate: today, time: '22:00', weekdays: [0, 4] }), 'A cada 2 semanas (quinta e domingo), às 22:00');
  assert.equal(d({ frequency: 'monthly', startDate: today, time: '03:00', monthlyMode: 'day', monthDay: 5 }), 'Todo mês, no dia 5, às 03:00');
  assert.equal(
    d({ frequency: 'monthly', interval: 3, startDate: today, time: '03:00', monthlyMode: 'day', monthDay: 31 }),
    'A cada 3 meses, no dia 31 (nos meses mais curtos, no último dia), às 03:00',
  );
  assert.equal(d({ frequency: 'monthly', startDate: today, time: '03:00', monthlyMode: 'last-day' }), 'Todo mês, no último dia, às 03:00');
  assert.equal(d({ frequency: 'monthly', startDate: today, time: '03:00', monthlyMode: 'weekday', weekOfMonth: 1, weekday: 1 }), 'Todo mês, na primeira segunda-feira, às 03:00');
  assert.equal(d({ frequency: 'monthly', startDate: today, time: '03:00', monthlyMode: 'weekday', weekOfMonth: -1, weekday: 6 }), 'Todo mês, no último sábado, às 03:00');
  assert.equal(
    d({ frequency: 'daily', startDate: '2026-10-01', time: '02:00', end: 'date', endDate: '2026-12-31' }),
    'Todos os dias, às 02:00, a partir de 01/10/2026, até 31/12/2026',
  );
  assert.equal(d({ frequency: 'daily', startDate: '2026-09-01', time: '02:00', end: 'count', count: 10 }), 'Todos os dias, às 02:00, 10 execuções no total');
  assert.equal(dateBr('2026-01-05'), '05/01/2026');
});
