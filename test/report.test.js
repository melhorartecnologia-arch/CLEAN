import { test } from 'node:test';
import assert from 'node:assert/strict';
import { unzipSync, strFromU8 } from 'fflate';
import { buildXlsx } from '../src/report/xlsx.js';
import { exportCsv } from '../src/report/exports.js';
import { filterRecords } from '../src/report/model.js';

test('abas maiores que o limite do Excel continuam em novas abas', async () => {
  const rows = Array.from({ length: 12 }, (_, i) => [`linha ${i + 1}`, i + 1]);
  const xlsx = await buildXlsx([{ name: 'Dados', header: ['Texto', 'Número'], rows }], { maxRows: 5 });
  const parts = unzipSync(new Uint8Array(xlsx));
  const workbook = strFromU8(parts['xl/workbook.xml']);
  assert.deepEqual(
    [...workbook.matchAll(/<sheet name="([^"]*)"/g)].map((m) => m[1]),
    ['Dados', 'Dados (2)', 'Dados (3)'],
  );
  const counts = [1, 2, 3].map((n) => (strFromU8(parts[`xl/worksheets/sheet${n}.xml`]).match(/<row /g) || []).length);
  assert.deepEqual(counts, [5, 5, 5]); // 4 + 4 + 4 linhas de dados, cada aba com cabeçalho
  assert.ok(strFromU8(parts['xl/worksheets/sheet3.xml']).includes('linha 12'));
  assert.ok(strFromU8(parts['xl/worksheets/sheet1.xml']).includes('<autoFilter ref="A1:B5"/>'));
});

test('exporta 150 mil linhas sem estourar a pilha', async () => {
  function* rows() {
    for (let i = 0; i < 150000; i++) yield [`arquivo ${i}.docx`, i, new Date(2026, 0, 1)];
  }
  const xlsx = await buildXlsx([{ name: 'Ocorrências', header: ['Arquivo', 'N', 'Data'], rows: rows() }]);
  const parts = unzipSync(new Uint8Array(xlsx), { filter: (f) => f.name === 'xl/workbook.xml' || f.name === '[Content_Types].xml' });
  assert.ok(strFromU8(parts['xl/workbook.xml']).includes('Ocorrências'));
  assert.ok(xlsx.length > 100000);
});

test('CSV: datas no formato do Excel em português e fórmulas neutralizadas', async () => {
  const chunks = [];
  const out = { write: (c) => (chunks.push(c), true) };
  const record = {
    id: 1,
    repositoryName: 'R',
    name: '=cmd.xlsx',
    path: 'C:\\R\\=cmd.xlsx',
    relativePath: '=cmd.xlsx',
    modified: new Date(2026, 8, 25, 13, 34, 21).toISOString(),
    lastUser: '@usuario',
    lastUserSource: 'owner',
    owner: '@usuario',
    metadata: {},
    matches: [{ term: 'CPF', list: 'L', location: 'content', count: 1, values: ['1'], samples: [] }],
    terms: ['CPF'],
    occurrences: 1,
  };
  await exportCsv([record], out);
  const csv = chunks.join('');
  const line = csv.split('\r\n')[1];
  assert.ok(line.includes('25/09/2026 13:34:21'), line);
  assert.ok(line.includes("'=cmd.xlsx"));
  assert.ok(line.includes("'@usuario"));
});

test('ordenação desconhecida ou maliciosa usa a padrão', () => {
  const records = [
    { path: 'b', name: 'b', terms: [], matches: [] },
    { path: 'a', name: 'a', terms: [], matches: [] },
  ];
  for (const sort of ['__proto__', 'hasOwnProperty', 'valueOf', 'constructor']) {
    assert.deepEqual(filterRecords(records, { sort }).map((r) => r.path), ['a', 'b']);
  }
});
