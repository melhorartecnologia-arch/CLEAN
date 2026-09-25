import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import CFB from 'cfb';
import { parseMime, decodeHeader, parseAddresses, parseStructuredHeader, formatAddress } from '../src/scan/extractors/mime.js';
import { extractMessage, extractBuffer } from '../src/scan/extractors/index.js';

const FIXTURES = path.join(import.meta.dirname, 'fixtures');
const DOCX = fs.readFileSync(path.join(FIXTURES, 'doc.docx'));
const PDF = fs.readFileSync(path.join(FIXTURES, 'doc.pdf'));
const CPF = '529.982.247-25';

const b64 = (buf) =>
  Buffer.from(buf)
    .toString('base64')
    .replace(/.{76}/g, '$&\r\n');

function message({ headers = [], parts }) {
  return Buffer.from(
    [
      'From: =?UTF-8?Q?Jo=C3=A3o_Silva?= <joao@empresa.com.br>',
      'To: "Silva, Maria" <maria@empresa.com.br>, rh@empresa.com.br',
      'Subject: =?UTF-8?B?UmVsYXTDs3JpbyBkZSBzYWzDoXJpb3M=?=',
      'Date: Thu, 25 Sep 2026 10:00:00 -0300 (BRT)',
      'Message-ID: <abc@empresa.com.br>',
      'MIME-Version: 1.0',
      ...headers,
      'Content-Type: multipart/mixed; boundary="B1"',
      '',
      'Preâmbulo ignorado',
      ...parts.flatMap((p) => ['--B1', ...p]),
      '--B1--',
      'Epílogo ignorado',
      '',
    ].join('\r\n'),
    'utf8',
  );
}

const textPart = (text) => ['Content-Type: text/plain; charset=utf-8', 'Content-Transfer-Encoding: quoted-printable', '', text];

test('cabeçalhos codificados: palavras vizinhas, bytes UTF-8 crus e RFC 2231', () => {
  // "ção" dividido entre duas palavras codificadas (o "ç" ocupa 2 bytes, cortados ao meio)
  assert.equal(decodeHeader('=?UTF-8?B?w6c=?= =?UTF-8?B?w6Nv?='), 'ção');
  assert.equal(decodeHeader('Re: =?iso-8859-1?Q?Sal=E1rio?= de junho'), 'Re: Salário de junho');
  assert.equal(decodeHeader(Buffer.from('Relatório', 'utf8').toString('latin1')), 'Relatório');
  const ct = parseStructuredHeader("attachment; filename*0*=utf-8''rela%C3%A7%C3%A3o%20de; filename*1*=%20sal%C3%A1rios.pdf");
  assert.equal(ct.value, 'attachment');
  assert.equal(ct.params.filename, 'relação de salários.pdf');
  assert.equal(parseStructuredHeader('application/pdf; name="=?UTF-8?Q?or=C3=A7amento.pdf?="').params.name, 'orçamento.pdf');
  assert.equal(parseStructuredHeader('text/plain; charset="UTF-8"; format=flowed').params.charset, 'UTF-8');
});

test('endereços: nomes com vírgula, comentários, grupos e codificação', () => {
  const list = parseAddresses([
    '"Silva, Maria" <maria@x.com>, =?UTF-8?Q?Jo=C3=A3o?= <joao@x.com>, pedro@x.com (Pedro Souza), Equipe: ana@x.com, bia@x.com;',
  ]);
  assert.deepEqual(
    list.map((a) => [a.name, a.address]),
    [
      ['Silva, Maria', 'maria@x.com'],
      ['João', 'joao@x.com'],
      ['Pedro Souza', 'pedro@x.com'],
      ['', 'ana@x.com'],
      ['', 'bia@x.com'],
    ],
  );
  assert.equal(formatAddress(list[1]), 'João <joao@x.com>');
  assert.equal(formatAddress({ name: 'x@y.com', address: 'x@y.com' }), 'x@y.com');
});

test('estrutura: corpo alternativo, anexos, preâmbulo e epílogo ignorados', () => {
  const raw = message({
    parts: [
      [
        'Content-Type: multipart/alternative; boundary="B2"',
        '',
        '--B2',
        ...textPart('Veja o anexo.'),
        '--B2',
        'Content-Type: text/html; charset=utf-8',
        '',
        '<p>Veja o anexo, <b>por favor</b>.</p>',
        '--B2--',
      ],
      ['Content-Type: application/pdf; name="=?UTF-8?Q?sal=C3=A1rios.pdf?="', 'Content-Disposition: attachment', 'Content-Transfer-Encoding: base64', '', b64(PDF)],
      ['Content-Type: image/png', 'Content-ID: <logo@x>', 'Content-Transfer-Encoding: base64', '', 'iVBORw0KGgo='],
    ],
  });
  const m = parseMime(raw);
  assert.equal(m.subject, 'Relatório de salários');
  assert.equal(m.from[0].name, 'João Silva');
  assert.equal(m.to.length, 2);
  assert.equal(m.date, '2026-09-25T13:00:00.000Z');
  assert.equal(m.messageId, '<abc@empresa.com.br>');
  assert.equal(m.texts.length, 1, 'uma única versão do texto alternativo');
  assert.match(m.texts[0].text, /por favor/, 'fica a versão com mais conteúdo (HTML)');
  assert.deepEqual(
    m.attachments.map((a) => [a.name, a.inline, a.size]),
    [
      ['salários.pdf', false, PDF.length],
      ['imagem-embutida.png', true, 8],
    ],
  );
  assert.ok(!m.texts.some((t) => /Preâmbulo|Epílogo/.test(t.text)));
});

test('conteúdo dos anexos (Word e PDF) e mensagem encaminhada como anexo', async () => {
  const forwarded = [
    'From: Ana <ana@fornecedor.com>',
    'Subject: Planilha do fornecedor',
    'Content-Type: multipart/mixed; boundary="IN"',
    '',
    '--IN',
    'Content-Type: text/plain',
    '',
    'Mensagem interna com o CNPJ 11.222.333/0001-81.',
    '--IN',
    'Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document; name="contrato.docx"',
    'Content-Transfer-Encoding: base64',
    '',
    b64(DOCX),
    '--IN--',
  ].join('\r\n');
  const raw = message({
    parts: [textPart('Segue.'), ['Content-Type: message/rfc822', 'Content-Disposition: attachment', '', forwarded]],
  });
  const m = await extractMessage(raw);
  assert.equal(m.body, 'Segue.');
  assert.equal(m.attachments.length, 1);
  const fwd = m.attachments[0];
  assert.equal(fwd.name, 'Planilha do fornecedor.eml');
  assert.equal(fwd.status, 'ok');
  const texts = fwd.segments.map((s) => `${s.label}: ${s.text}`).join('\n');
  assert.match(texts, /Anexo "Planilha do fornecedor.eml" › Mensagem: Mensagem interna com o CNPJ 11\.222\.333\/0001-81/);
  assert.ok(fwd.segments.some((s) => s.label.startsWith('Anexo "Planilha do fornecedor.eml" › Anexo "contrato.docx"') && s.text.includes(CPF)), texts);
});

test('mensagem cortada no limite de tamanho: o anexo incompleto não é lido', async () => {
  const raw = message({
    parts: [textPart('Corpo completo.'), ['Content-Type: application/pdf; name="grande.pdf"', 'Content-Transfer-Encoding: base64', '', b64(PDF)]],
  });
  const cut = raw.subarray(0, raw.indexOf('grande.pdf') + 200);
  const m = await extractMessage(cut, { truncated: true });
  assert.equal(m.body, 'Corpo completo.');
  assert.equal(m.partial, true);
  assert.equal(m.attachments[0].status, 'skipped-size');
});

test('S/MIME criptografado e assinatura ignorada', async () => {
  const encrypted = Buffer.from(
    ['Subject: Sigiloso', 'Content-Type: application/pkcs7-mime; smime-type=enveloped-data; name=smime.p7m', 'Content-Transfer-Encoding: base64', '', 'MIAGCSqGSIb3DQEHA6CAMIACAQA='].join('\r\n'),
  );
  const e = await extractMessage(encrypted);
  assert.equal(e.encrypted, true);
  assert.equal(e.attachments.length, 0);
  const signed = Buffer.from(
    [
      'Subject: Assinado',
      'Content-Type: multipart/signed; protocol="application/pkcs7-signature"; boundary="S"',
      '',
      '--S',
      'Content-Type: text/plain',
      '',
      'Texto assinado com CPF 529.982.247-25',
      '--S',
      'Content-Type: application/pkcs7-signature; name=smime.p7s',
      'Content-Transfer-Encoding: base64',
      '',
      'MIAGCSqGSIb3DQEHAqCAMIACAQE=',
      '--S--',
    ].join('\r\n'),
  );
  const s = await extractMessage(signed);
  assert.match(s.body, /CPF 529/);
  assert.equal(s.attachments.length, 0, 'a assinatura não é listada como anexo');
});

test('arquivo .eml em disco (extractBuffer) inclui o conteúdo dos anexos', async () => {
  const raw = message({ parts: [textPart('Oi'), ['Content-Type: application/octet-stream; name="doc.docx"', 'Content-Transfer-Encoding: base64', '', b64(DOCX)]] });
  const r = await extractBuffer(raw, { name: 'mensagem.eml' });
  assert.equal(r.type, 'eml');
  assert.equal(r.status, 'ok');
  assert.ok(r.segments.some((s) => s.label === 'Anexo "doc.docx" › Documento' && s.text.includes(CPF)));
  assert.equal(r.attachments[0].status, 'ok');
  assert.equal(r.attachments[0].type, 'docx');
});

test('anexos binários de um .msg do Outlook são lidos', async () => {
  const cfb = CFB.utils.cfb_new();
  const u16 = (s) => Buffer.from(`${s}\0`, 'utf16le');
  CFB.utils.cfb_add(cfb, '/__substg1.0_0037001F', u16('Contrato'));
  CFB.utils.cfb_add(cfb, '/__substg1.0_1000001F', u16('Segue o contrato.'));
  CFB.utils.cfb_add(cfb, '/__attach_version1.0_#00000000/__substg1.0_3707001F', u16('contrato.docx'));
  CFB.utils.cfb_add(cfb, '/__attach_version1.0_#00000000/__substg1.0_37010102', DOCX);
  const r = await extractBuffer(CFB.write(cfb, { type: 'buffer' }), { name: 'contrato.msg' });
  assert.equal(r.type, 'msg');
  assert.ok(r.segments.some((s) => s.label === 'Anexo "contrato.docx" › Documento' && s.text.includes(CPF)), JSON.stringify(r.segments.map((s) => s.label)));
});

test('limite de texto por mensagem vale para a soma dos anexos', async () => {
  const raw = message({
    parts: [
      textPart('x'.repeat(50)),
      ['Content-Type: text/plain; name="a.txt"', '', 'y'.repeat(80)],
      ['Content-Type: text/plain; name="b.txt"', '', 'z'.repeat(80)],
    ],
  });
  const m = await extractMessage(raw, { limits: { maxChars: 100 } });
  assert.equal(m.attachments[0].status, 'partial');
  assert.equal(m.attachments[1].status, 'skipped-size');
});
