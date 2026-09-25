import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { zipSync, strToU8 } from 'fflate';
import CFB from 'cfb';
import { extractFile } from '../src/scan/extractors/index.js';
import { decodeText, looksLikeText, binaryStrings } from '../src/scan/extractors/text.js';
import { pdfDate } from '../src/scan/extractors/pdf.js';

const FIXTURES = path.join(import.meta.dirname, 'fixtures');
let tmp;

before(() => {
  tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'clean-extract-'));
});
after(() => fs.rmSync(tmp, { recursive: true, force: true }));

async function extract(file, options = {}) {
  const full = path.isAbsolute(file) ? file : path.join(FIXTURES, file);
  return extractFile(full, { size: fs.statSync(full).size, ...options });
}
const allText = (r) => r.segments.map((s) => s.text).join('\n');

const DOC_TEXT = 'Funcionário: João da Silva, CPF 529.982.247-25.';

for (const file of ['doc.doc', 'doc.docx', 'doc.odt', 'doc.rtf', 'doc.pdf']) {
  test(`texto e metadados de ${file}`, async () => {
    const r = await extract(file);
    assert.equal(r.status, 'ok');
    assert.equal(r.type, path.extname(file).slice(1));
    const text = allText(r);
    assert.ok(text.includes(DOC_TEXT), text);
    assert.ok(text.includes('Célula B senha: abc123'));
    assert.equal(r.metadata.author, 'Joana Autora');
    if (file !== 'doc.pdf') {
      assert.equal(r.metadata.lastModifiedBy, 'Marcos Revisor');
      assert.equal(r.metadata.title, 'Relatório de Salários');
    }
  });
}

for (const file of ['plan.xls', 'plan.xlsx', 'plan.ods']) {
  test(`planilhas com nomes das abas e linhas: ${file}`, async () => {
    const r = await extract(file);
    assert.equal(r.status, 'ok');
    assert.deepEqual(
      r.segments.map((s) => s.label),
      ['Planilha "Folha de Pagamento"', 'Planilha "Resumo"'],
    );
    const lines = r.segments[0].text.split('\n');
    assert.equal(lines[1], 'José Ações\t52998224725\t4500.5'); // CPF gravado como número
    assert.equal(lines[2], 'Maria CONFIDENCIAL\t111.444.777-35\t12');
    assert.equal(r.segments[0].lines, true);
    assert.equal(r.metadata.lastModifiedBy, 'Carlos Financeiro');
    assert.equal(r.metadata.author, 'Ana Planilha');
  });
}

for (const file of ['gap.xls', 'gap.xlsx', 'gap.ods']) {
  test(`linhas vazias preservam a numeração: ${file}`, async () => {
    const r = await extract(file);
    const lines = r.segments[0].text.split('\n');
    assert.equal(lines[0], 'Cabeçalho');
    assert.ok(lines[6].startsWith('linha sete CONFIDENCIAL'), JSON.stringify(lines.slice(0, 8)));
    // Depois de 30 mil linhas vazias a numeração deixa de ser exata
    assert.equal(r.segments[0].lines, false);
    assert.ok(r.segments[0].text.includes('depois do vazio'));
  });
}

for (const file of ['apres.ppt', 'apres.pptx', 'apres.odp']) {
  test(`apresentações: ${file}`, async () => {
    const r = await extract(file);
    const text = allText(r);
    assert.ok(text.includes('Plano de Demissão 2025'));
    assert.ok(text.includes('fulano@empresa.com.br'));
    assert.ok(!text.includes('Click to edit'), 'não deve incluir textos do slide-mestre');
    assert.equal(r.metadata.lastModifiedBy, 'Rui Apresentador');
    if (file !== 'apres.ppt') assert.deepEqual(r.segments.map((s) => s.label), ['Slide 1', 'Slide 2']);
  });
}

test('Word 97: instruções de campo são removidas e notas/cabeçalhos incluídos', async () => {
  const text = allText(await extract('campos.doc'));
  assert.ok(text.includes('Veja o portal do RH agora.'));
  assert.ok(!text.includes('HYPERLINK'));
  assert.ok(text.includes('Rodapé com CPF 529.982.247-25'));
  assert.ok(text.includes('Cabeçalho SIGILOSO'));
});

test('RTF: nota de rodapé', async () => {
  const text = allText(await extract('nota.rtf'));
  assert.ok(text.includes('Rodapé com CPF 529.982.247-25'));
});

test('arquivos protegidos por senha', async () => {
  for (const file of ['senha.docx', 'senha.pdf']) {
    const r = await extract(file);
    assert.equal(r.status, 'encrypted', file);
    assert.equal(r.type, path.extname(file).slice(1));
    assert.match(r.note, /senha/);
  }
});

test('texto em UTF-16 (com e sem BOM), Windows-1252 e UTF-8', async () => {
  const text = 'Situação: salário de João\r\n';
  const files = {
    'utf16.txt': Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from(text, 'utf16le')]),
    'utf16-sem-bom.txt': Buffer.from(text, 'utf16le'),
    'latin.csv': latin1(text),
    'utf8.log': Buffer.from(text, 'utf8'),
  };
  for (const [name, data] of Object.entries(files)) {
    const file = path.join(tmp, name);
    fs.writeFileSync(file, data);
    const r = await extract(file);
    assert.equal(r.type, 'texto', name);
    assert.equal(r.segments[0].text, text, name);
    assert.equal(r.segments[0].lines, true);
  }
});

function latin1(text) {
  return Buffer.from([...text].map((c) => c.charCodeAt(0)));
}

test('HTML: remove marcas e decodifica entidades', async () => {
  const file = path.join(tmp, 'pagina.htm');
  fs.writeFileSync(file, '<html><head><title>RH &ndash; Intranet</title><style>p{}</style></head><body><p>Sal&aacute;rio de Jos&#233;</p></body></html>');
  const r = await extract(file);
  assert.equal(r.type, 'html');
  assert.ok(allText(r).includes('Salário de José'));
  assert.ok(allText(r).includes('RH – Intranet'));
  assert.ok(!allText(r).includes('p{}'));
});

test('ZIP comum: nomes dos arquivos compactados viram conteúdo', async () => {
  const file = path.join(tmp, 'backup.zip');
  fs.writeFileSync(file, zipSync({ 'pasta/salarios.xlsx': strToU8('x'), 'leia.txt': strToU8('y') }));
  const r = await extract(file);
  assert.equal(r.type, 'zip');
  assert.equal(r.segments[0].text, 'leia.txt\npasta/salarios.xlsx');
});

test('limites: arquivo grande, vazio, binário e texto parcial', async () => {
  const empty = path.join(tmp, 'vazio.txt');
  fs.writeFileSync(empty, '');
  assert.equal((await extract(empty)).status, 'empty');

  const bin = path.join(tmp, 'imagem.png');
  fs.writeFileSync(bin, Buffer.from([0x89, 0x50, 0x4e, 0x47, 0, 0, 0, 1, 2, 3]));
  assert.equal((await extract(bin)).status, 'unsupported');

  const limits = { maxBytes: 1000, maxChars: 20_000_000 };
  const big = await extract('doc.pdf', { limits });
  assert.equal(big.status, 'skipped-size');

  const log = path.join(tmp, 'grande.log');
  fs.writeFileSync(log, 'CPF 529.982.247-25\n'.repeat(200));
  const partial = await extract(log, { limits });
  assert.equal(partial.status, 'partial');
  assert.equal(partial.segments[0].text.length, 1000);

  const chars = await extract(log, { limits: { maxBytes: 1e6, maxChars: 50 } });
  assert.equal(chars.status, 'partial');
  assert.equal(chars.segments[0].text.length, 50);
});

test('somente metadados (sem ler o texto)', async () => {
  const r = await extract('doc.docx', { withText: false });
  assert.equal(r.segments.length, 0);
  assert.equal(r.metadata.lastModifiedBy, 'Marcos Revisor');
  const rtf = await extract('doc.rtf', { withText: false });
  assert.equal(rtf.metadata.lastModifiedBy, 'Marcos Revisor');
});

test('e-mail do Outlook (.msg): assunto, corpo, destinatários e anexos', async () => {
  const cfb = CFB.utils.cfb_new();
  const u16 = (s) => Buffer.from(`${s}\0`, 'utf16le');
  CFB.utils.cfb_add(cfb, '/__substg1.0_0037001F', u16('Folha de pagamento'));
  CFB.utils.cfb_add(cfb, '/__substg1.0_1000001F', u16('Segue a planilha com o CPF 529.982.247-25.'));
  CFB.utils.cfb_add(cfb, '/__substg1.0_0C1A001F', u16('Maria Remetente'));
  CFB.utils.cfb_add(cfb, '/__substg1.0_007D001F', u16('Received: from servidor-interno'));
  CFB.utils.cfb_add(cfb, '/__recip_version1.0_#00000000/__substg1.0_3001001F', u16('João Destinatário'));
  CFB.utils.cfb_add(cfb, '/__attach_version1.0_#00000000/__substg1.0_3707001F', u16('salarios_2025.xlsx'));
  const file = path.join(tmp, 'mensagem.msg');
  fs.writeFileSync(file, CFB.write(cfb, { type: 'buffer' }));
  const r = await extract(file);
  assert.equal(r.type, 'msg');
  const text = allText(r);
  for (const part of ['Folha de pagamento', 'CPF 529.982.247-25', 'João Destinatário', 'salarios_2025.xlsx']) {
    assert.ok(text.includes(part), part);
  }
  assert.ok(!text.includes('Received:'), 'cabeçalhos de transporte são ignorados');
  assert.equal(r.metadata.author, 'Maria Remetente');
  assert.equal(r.metadata.title, 'Folha de pagamento');
});

test('utilitários de texto', () => {
  assert.equal(decodeText(latin1('ação')), 'ação');
  assert.equal(looksLikeText(Buffer.from('abc\ndef')), true);
  assert.equal(looksLikeText(Buffer.from([1, 2, 3, 0, 5])), false);
  assert.equal(looksLikeText(Buffer.from([1, 0, 2, 0, 3, 0, 4, 0, 5, 0, 6, 0, 7, 0, 8, 0, 9, 0])), false);
  assert.equal(looksLikeText(Buffer.from('texto em UTF-16 sem BOM', 'utf16le')), true);
  const bin = Buffer.concat([Buffer.from([0, 1, 2]), Buffer.from('Relatório', 'latin1'), Buffer.from([0, 0]), Buffer.from('Confidencial', 'utf16le'), Buffer.from([0, 0, 1])]);
  const strings = binaryStrings(bin);
  assert.ok(strings.includes('Relatório'));
  assert.ok(strings.includes('Confidencial'));
  assert.equal(pdfDate("D:20240301103000-03'00'"), '2024-03-01T13:30:00.000Z');
  assert.equal(pdfDate('D:2024'), '2024-01-01T00:00:00.000Z');
  assert.equal(pdfDate('lixo'), null);
});

// ---------- Regressões encontradas na revisão ----------

test('objeto incorporado não muda o tipo nem os metadados do arquivo OLE', async () => {
  const r = await extract('embutido.xls');
  assert.equal(r.type, 'xls');
  assert.ok(allText(r).includes('111.444.777-35'), 'as células da planilha são lidas');
  assert.equal(r.metadata.lastModifiedBy, 'Carlos Financeiro');
});

test('caixa de texto do Word não é contada duas vezes', async () => {
  const text = allText(await extract('caixa-texto.docx'));
  assert.equal(text.match(/SEGREDO/gi).length, 1);
});

for (const file of ['multilinha.xlsx', 'multilinha.xls', 'multilinha.ods']) {
  test(`célula com várias linhas não desloca a numeração: ${file}`, async () => {
    const lines = (await extract(file)).segments[0].text.split('\n');
    assert.match(lines[3], /SEGREDO/i, JSON.stringify(lines));
  });
}

test('planilha com prefixo de namespace (Open XML SDK)', async () => {
  const text = allText(await extract('prefixo.xlsx'));
  assert.ok(text.includes('CPF 529.982.247-25 CONFIDENCIAL'), text);
});

test('ZIP do Windows Explorer: nomes na página de código 850', async () => {
  const r = await extract('explorer.zip');
  assert.equal(r.segments[0].text, 'Relatório de Salários.xlsx');
});

test('página .mht com quoted-printable', async () => {
  const r = await extract('pagina.mht');
  assert.equal(r.type, 'mht');
  assert.ok(allText(r).includes('Relação de salários: João da Silva CONFIDENCIAL'), allText(r));
});

test('e-mail .eml com corpo em base64, assunto codificado e anexo', async () => {
  const file = path.join(tmp, 'mensagem.eml');
  const body = Buffer.from('Segue a folha de pagamento. CPF 529.982.247-25', 'utf8').toString('base64');
  fs.writeFileSync(
    file,
    [
      'From: =?UTF-8?Q?Jo=C3=A3o_Silva?= <joao@empresa.com.br>',
      'To: rh@empresa.com.br',
      'Subject: =?UTF-8?B?U2Fsw6FyaW9zIGRlIHNldGVtYnJv?=',
      'Date: Thu, 25 Sep 2026 10:00:00 -0300',
      'MIME-Version: 1.0',
      'Content-Type: multipart/mixed; boundary="XYZ"',
      '',
      '--XYZ',
      'Content-Type: multipart/alternative; boundary="ALT"',
      '',
      '--ALT',
      'Content-Type: text/plain; charset=utf-8',
      'Content-Transfer-Encoding: base64',
      '',
      body,
      '--ALT',
      'Content-Type: text/html; charset=utf-8',
      'Content-Transfer-Encoding: base64',
      '',
      Buffer.from('<p>Segue a folha de pagamento. CPF 529.982.247-25</p>').toString('base64'),
      '--ALT--',
      '--XYZ',
      'Content-Type: application/vnd.ms-excel; name="folha.xls"',
      'Content-Disposition: attachment; filename*=UTF-8\'\'sal%C3%A1rios%202026.xls',
      'Content-Transfer-Encoding: base64',
      '',
      'AAAA',
      '--XYZ--',
      '',
    ].join('\r\n'),
  );
  const r = await extract(file);
  assert.equal(r.type, 'eml');
  const text = allText(r);
  assert.equal(text.match(/folha de pagamento/g).length, 1, 'multipart/alternative gera um único texto');
  assert.ok(text.includes('Assunto: Salários de setembro'));
  assert.ok(text.includes('salários 2026.xls'));
  assert.equal(r.metadata.title, 'Salários de setembro');
  assert.match(r.metadata.author, /João Silva/);
});

test('arquivo que não pode ser aberto vira erro de leitura (o nome continua valendo)', async () => {
  const r = await extractFile(path.join(tmp, 'nao-existe.docx'), { size: 10 });
  assert.equal(r.status, 'error');
  assert.match(r.note, /Não encontrado/);
});

test('texto grande cortado no meio de um caractere continua em UTF-8', async () => {
  const file = path.join(tmp, 'grande-utf8.txt');
  fs.writeFileSync(file, 'salário ação João '.repeat(100));
  // 1001 bytes: o corte cai no meio de um caractere de 2 bytes
  const r = await extract(file, { limits: { maxBytes: 1001, maxChars: 1e7 } });
  assert.equal(r.status, 'partial');
  assert.ok(r.segments[0].text.startsWith('salário ação João'), r.segments[0].text.slice(0, 40));
});

test('metadados de arquivos grandes demais para o conteúdo', async () => {
  const limits = { maxBytes: 1000, maxChars: 1e7, maxMetadataBytes: 1e8 };
  const full = await extract('doc.docx', { limits });
  assert.equal(full.status, 'skipped-size');
  const meta = await extract('doc.docx', { limits, withText: false });
  assert.equal(meta.metadata.lastModifiedBy, 'Marcos Revisor');
});
