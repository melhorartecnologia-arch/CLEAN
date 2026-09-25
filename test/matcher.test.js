import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Matcher, foldText, validateTerm } from '../src/scan/matcher.js';
import { isValidCpf, isValidCnpj, isValidCard, isValidPis, PRESETS } from '../src/scan/presets.js';

const text = (value, extra = {}) => ({ id: value, value, type: 'text', ...extra });
const regex = (id, value, extra = {}) => ({ id, value, type: 'regex', ...extra });

test('foldText remove acentos, maiúsculas e espaços repetidos', () => {
  assert.equal(foldText('  AÇÃO   Salário\tJOÃO '), 'acao salario joao');
  assert.equal(foldText('İstanbul Ærø Ñandú'), 'istanbul ærø nandu');
});

test('termo de texto ignora acentos e maiúsculas', () => {
  const m = new Matcher([text('salario')]);
  const [hit] = m.match([{ text: 'Folha: SALÁRIO de março e salário de abril' }], 'content');
  assert.equal(hit.count, 2);
  assert.deepEqual(hit.values, ['SALÁRIO', 'salário']);
  assert.equal(hit.samples[0].match, 'SALÁRIO');
  assert.equal(hit.samples[0].before, 'Folha: ');
  assert.equal(hit.location, 'content');
});

test('termo com acento encontra texto sem acento e espaços variados', () => {
  const m = new Matcher([text('João da Silva')]);
  const [hit] = m.match([{ text: 'Responsável: joao  da\r\nsilva (RH)' }], 'content');
  assert.equal(hit.count, 1);
  assert.equal(hit.samples[0].match, 'joao da silva');
  assert.equal(hit.samples[0].after, ' (RH)');
});

test('texto em forma decomposta (NFD) é normalizado', () => {
  const m = new Matcher([text('café')]);
  const [hit] = m.match([{ text: 'café com leite' }], 'content');
  assert.equal(hit.count, 1);
});

test('palavra inteira respeita letras vizinhas, mas aceita _ e pontuação', () => {
  const m = new Matcher([text('ana', { wholeWord: true }), text('rh', { id: 'rh', wholeWord: false })]);
  assert.deepEqual(m.match([{ text: 'banana e mariana' }], 'name'), []);
  const hits = m.match([{ text: 'relatorio_ana.xlsx - Ana, RH' }], 'name');
  const ana = hits.find((h) => h.termId === 'ana');
  assert.equal(ana.count, 2);
  assert.equal(hits.find((h) => h.termId === 'rh').count, 1);
});

test('muitos termos e sobreposição (Aho-Corasick)', () => {
  const terms = ['he', 'she', 'his', 'hers'].map((v) => text(v));
  const m = new Matcher(terms);
  const hits = m.match([{ text: 'ushers' }], 'content');
  const counts = Object.fromEntries(hits.map((h) => [h.termId, h.count]));
  assert.deepEqual(counts, { she: 1, he: 1, hers: 1 });

  const many = Array.from({ length: 2000 }, (_, i) => text(`codigo-${i}`, { wholeWord: true }));
  const big = new Matcher(many);
  const found = big.match([{ text: 'itens codigo-15, CODIGO-1999 e codigo-150x' }], 'content');
  assert.deepEqual(found.map((h) => h.termId).sort(), ['codigo-15', 'codigo-1999']);
});

test('regex com validador de CPF descarta números inválidos', () => {
  const cpf = PRESETS.find((p) => p.id === 'cpf');
  const m = new Matcher([regex('cpf', cpf.value, { validator: 'cpf', label: 'CPF' })]);
  const [hit] = m.match([{ text: 'CPF 529.982.247-25, outro 111.111.111-11, 52998224726 e 52998224725.' }], 'content');
  assert.equal(hit.term, 'CPF');
  assert.equal(hit.count, 2);
  assert.deepEqual(hit.values, ['529.982.247-25', '52998224725']);
});

test('limite de ocorrências e número da linha nos exemplos', () => {
  const m = new Matcher([text('x', { wholeWord: true })], { maxCount: 5, maxSamples: 2 });
  const [hit] = m.match([{ text: 'a\nb\nx x x\nx x x x', lines: true, label: 'Planilha 1' }], 'content');
  assert.equal(hit.count, 5);
  assert.equal(hit.truncated, true);
  assert.equal(hit.samples.length, 2);
  assert.equal(hit.samples[0].where, 'Planilha 1, linha 3');
});

test('regex que casa vazio é rejeitada e inválida é reportada', () => {
  assert.match(validateTerm(regex('a', 'a*')), /vazio/);
  assert.match(validateTerm(regex('b', '([a-z')), /inválida/);
  assert.equal(validateTerm(regex('c', String.raw`\d+`)), null);
  const m = new Matcher([regex('bad', '(x'), text('ok')]);
  assert.equal(m.size, 1);
  assert.equal(m.invalid.length, 1);
});

test('segmentos múltiplos somam ocorrências e guardam o local', () => {
  const m = new Matcher([text('confidencial')]);
  const [hit] = m.match(
    [
      { text: 'nada aqui', label: 'Página 1' },
      { text: 'Documento CONFIDENCIAL', label: 'Página 2' },
      { text: 'confidencial de novo', label: 'Página 3' },
    ],
    'content',
  );
  assert.equal(hit.count, 2);
  assert.deepEqual(
    hit.samples.map((s) => s.where),
    ['Página 2', 'Página 3'],
  );
});

test('validadores de documentos', () => {
  assert.equal(isValidCpf('529.982.247-25'), true);
  assert.equal(isValidCpf('000.000.000-00'), false);
  assert.equal(isValidCnpj('11.222.333/0001-81'), true);
  assert.equal(isValidCnpj('11222333000182'), false);
  assert.equal(isValidCnpj('12.ABC.345/01DE-35'), true); // exemplo oficial de CNPJ alfanumérico
  assert.equal(isValidCnpj('12.abc.345/01de-35'), false);
  assert.equal(isValidCard('4111 1111 1111 1111'), true);
  assert.equal(isValidCard('4111 1111 1111 1112'), false);
  assert.equal(isValidPis('120.49826.65-5'), false);
  assert.equal(isValidPis('125.2851.620-4'.replace(/\D/g, '')), false);
  assert.equal(isValidPis('17033259504'), true);
});

test('modelos prontos compilam e encontram exemplos', () => {
  const m = new Matcher(PRESETS.map((p) => ({ ...p, id: p.id })));
  assert.equal(m.invalid.length, 0);
  const sample = [
    'CPF 529.982.247-25',
    'CNPJ 11.222.333/0001-81 e 12.ABC.345/01DE-35',
    'PIS 170.33259.50-4',
    'contato: maria.silva@empresa.com.br',
    'tel (24) 99999-1234',
    'cartão 4111-1111-1111-1111',
    'senha: Pa$$w0rd',
  ].join('\n');
  const found = Object.fromEntries(m.match([{ text: sample }], 'content').map((h) => [h.termId, h.count]));
  assert.deepEqual(found, { cpf: 1, cnpj: 2, pis: 1, email: 1, telefone: 1, cartao: 1, senha: 1 });
});
