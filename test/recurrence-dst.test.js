// Mudanças de horário de verão (fuso com horário de verão definido só neste arquivo de teste).
process.env.TZ = 'America/New_York';

const { test } = await import('node:test');
const assert = (await import('node:assert/strict')).default;
const { validateRule, upcoming } = await import('../src/schedule/recurrence.js');

const hours = (dates) => dates.map((d) => `${d.getDate()} ${d.getHours()}:${String(d.getMinutes()).padStart(2, '0')}`);

test('início do horário de verão: o horário que não existe passa para o seguinte, sem repetir', () => {
  const daily = validateRule({ frequency: 'daily', startDate: '2026-03-07', time: '02:30' });
  assert.deepEqual(hours(upcoming(daily, new Date(2026, 2, 7), 3)), ['7 2:30', '8 3:30', '9 2:30']);
  const hourly = validateRule({ frequency: 'hourly', startDate: '2026-03-08', time: '00:00', untilTime: '04:00', weekdays: [0] });
  const list = upcoming(hourly, new Date(2026, 2, 7, 12), 10).filter((d) => d.getDate() === 8);
  assert.deepEqual(hours(list), ['8 0:00', '8 1:00', '8 3:00', '8 4:00']);
  // Intervalos reais sempre positivos (nenhuma execução duplicada).
  for (let i = 1; i < list.length; i++) assert.ok(list[i] > list[i - 1]);
});

test('fim do horário de verão: a hora repetida executa uma vez só', () => {
  const hourly = validateRule({ frequency: 'hourly', startDate: '2026-11-01', time: '00:00', untilTime: '03:00', weekdays: [0] });
  const list = upcoming(hourly, new Date(2026, 9, 31, 12), 10).filter((d) => d.getDate() === 1 && d.getMonth() === 10);
  assert.deepEqual(hours(list), ['1 0:00', '1 1:00', '1 2:00', '1 3:00']);
  assert.equal((list[2] - list[1]) / 3600000, 2, 'da 1h (horário de verão) às 2h (horário padrão) passam 2 horas');
});
