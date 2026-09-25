// Identifica o formato real do conteúdo (pelos primeiros bytes e pela extensão) e extrai texto e
// metadados, de arquivos em disco ou de conteúdos já em memória (anexos de e-mail).
// Resultado: { type, status, segments, metadata, note, attachments? }.
//   status: ok | partial | empty | unsupported | encrypted | skipped-size | error
import fs from 'node:fs/promises';
import path from 'node:path';
import { ZipReader } from './zip.js';
import { detectOoxml, ooxmlMetadata, docxSegments, xlsxSegments, pptxSegments, naturalCompare, compact } from './ooxml.js';
import { detectOdf, odfIsEncrypted, odfMetadata, odfSegments } from './odf.js';
import { isOle, oleExtract } from './ole.js';
import { isPdf, pdfExtract } from './pdf.js';
import { isRtf, rtfExtract } from './rtf.js';
import { decodeText, looksLikeText } from './text.js';
import { htmlToText } from './xml.js';
import { MIME_EXTENSIONS, parseMime, formatAddress } from './mime.js';
import { zipEntryNames } from './zip.js';
import { friendlyError } from '../errors.js';

export const DEFAULT_LIMITS = {
  maxBytes: 50 * 1024 * 1024, // arquivos maiores: só o nome é analisado (texto puro: lê o início)
  maxChars: 20_000_000, // texto máximo extraído por arquivo
  maxMetadataBytes: 200 * 1024 * 1024, // limite para ler apenas os metadados (último usuário)
};

// Anexos dentro de anexos (ex.: e-mail encaminhado que traz um documento) são lidos até esta profundidade.
export const MAX_NESTING = 3;

const HTML_EXT = new Set(['.htm', '.html', '.xhtml', '.mht', '.mhtml', '.hta']);
const WEB_ARCHIVE_EXT = new Set(['.mht', '.mhtml', '.mhtm']);
// Extensões sabidamente binárias sem texto útil (evita ler arquivos grandes à toa).
const BINARY_EXT = new Set(
  (
    '.exe .dll .sys .msi .msp .cab .iso .img .vhd .vhdx .vmdk .ova .jpg .jpeg .png .gif .bmp .tif .tiff ' +
    '.ico .webp .heic .psd .raw .cr2 .nef .mp3 .mp4 .m4a .m4v .avi .mov .mkv .wmv .wma .wav .flac .ogg ' +
    '.mpg .mpeg .3gp .7z .rar .gz .tgz .bz2 .xz .tar .zst .class .pyc .o .obj .lib .pdb .so .dylib .db ' +
    '.sqlite .mdb .accdb .ldb .laccdb .pst .ost .ttf .otf .woff .woff2 .eot .dwg .dxf .bak .tmp .swp .lnk'
  ).split(' '),
);

async function readHead(filePath, length) {
  const fh = await fs.open(filePath, 'r');
  try {
    const buf = Buffer.alloc(length);
    const { bytesRead } = await fh.read(buf, 0, length, 0);
    return buf.subarray(0, bytesRead);
  } finally {
    await fh.close();
  }
}

function isZip(buf) {
  return buf.length >= 4 && buf[0] === 0x50 && buf[1] === 0x4b && (buf[2] === 3 || buf[2] === 5 || buf[2] === 7);
}

/** Classifica o conteúdo pelos primeiros bytes (e pela extensão, em último caso). */
function classify(head, ext) {
  if (isZip(head)) return 'zip';
  if (isOle(head)) return 'ole';
  if (isPdf(head)) return 'pdf';
  if (isRtf(head)) return 'rtf';
  if (BINARY_EXT.has(ext)) return 'binary';
  if (looksLikeText(head)) {
    const start = head.toString('latin1', 0, 512).trimStart().toLowerCase();
    if (HTML_EXT.has(ext) || start.startsWith('<!doctype html') || start.startsWith('<html')) return 'html';
    return 'text';
  }
  return 'binary';
}

function limitSegments(segments, maxChars) {
  let total = 0;
  const out = [];
  for (const seg of segments) {
    if (!seg || !seg.text) continue;
    if (total + seg.text.length > maxChars) {
      const rest = maxChars - total;
      if (rest > 0) out.push({ ...seg, text: seg.text.slice(0, rest) });
      return { segments: out, truncated: true };
    }
    total += seg.text.length;
    out.push(seg);
  }
  return { segments: out, truncated: false };
}

const textLength = (segments) => (segments || []).reduce((sum, s) => sum + (s?.text?.length || 0), 0);

async function extractZip(buf, { withText }) {
  const zip = new ZipReader(buf);
  const ooxml = detectOoxml(zip);
  if (ooxml) {
    const metadata = ooxmlMetadata(zip);
    if (!withText) return { type: ooxml, metadata };
    const segments = ooxml === 'docx' ? docxSegments(zip) : ooxml === 'xlsx' ? xlsxSegments(zip) : pptxSegments(zip);
    return { type: ooxml, metadata, segments, partial: zip.truncated };
  }
  const odf = detectOdf(zip);
  if (odf) {
    const metadata = odfMetadata(zip);
    if (odfIsEncrypted(zip)) return { type: odf, metadata, encrypted: true };
    if (!withText) return { type: odf, metadata };
    return { type: odf, metadata, segments: odfSegments(zip, odf), partial: zip.truncated };
  }
  // ZIP comum: os nomes dos arquivos compactados entram como conteúdo.
  const names = (zipEntryNames(buf) || zip.names()).sort(naturalCompare);
  return { type: 'zip', metadata: {}, segments: [{ label: 'Arquivos compactados', text: names.join('\n'), lines: true }] };
}

const attachmentLabel = (name, segment) => `Anexo "${name}"${segment.label ? ` › ${segment.label}` : ''}`;

/** Resumo de um anexo para os relatórios (sem o texto). */
function attachmentSummary(info) {
  return compact({
    name: info.name,
    size: info.size,
    contentType: info.contentType,
    type: info.type,
    status: info.status,
    note: info.note,
    inline: info.inline || null,
  });
}

/**
 * Lê o conteúdo de cada anexo com os mesmos leitores usados para arquivos. O texto de todos os
 * anexos somado respeita limits.maxChars (descontado o que já foi extraído antes: `used`).
 * list: [{ name, data, size, contentType?, inline?, message?, incomplete? }]
 */
async function readAttachments(list, { limits, depth, used = 0 }) {
  const out = [];
  let total = used;
  for (const att of list) {
    const info = { name: att.name, size: att.size ?? att.data?.length ?? 0, contentType: att.contentType || null, inline: att.inline, type: null, status: null, note: null, metadata: {}, segments: [] };
    const ext = path.extname(att.name).toLowerCase();
    let r;
    if (att.incomplete) {
      r = { type: ext.slice(1) || null, status: 'skipped-size', note: 'O anexo ficou além do limite de tamanho da mensagem e não foi analisado.' };
    } else if (depth >= MAX_NESTING) {
      r = { type: ext.slice(1) || null, status: 'unsupported', note: 'Anexo dentro de outros anexos em muitos níveis: apenas o nome foi verificado.' };
    } else if (!att.data) {
      r = { type: ext.slice(1) || null, status: 'unsupported' };
    } else if (total >= limits.maxChars) {
      r = { type: ext.slice(1) || null, status: 'skipped-size', note: 'Limite de texto por mensagem atingido: anexo não analisado.' };
    } else {
      r = await extractBuffer(att.data, { name: att.name, limits: { ...limits, maxChars: limits.maxChars - total }, depth: depth + 1, mime: Boolean(att.message) });
    }
    info.type = r.type || null;
    info.status = r.status;
    info.note = r.note || null;
    info.metadata = r.metadata || {};
    info.segments = (r.segments || []).map((seg) => ({ ...seg, label: attachmentLabel(att.name, seg) }));
    total += textLength(info.segments);
    att.data = null; // libera a memória do anexo já lido
    out.push(info);
  }
  return out;
}

/**
 * Extrai uma mensagem MIME: remetente, destinatários, assunto, corpo e anexos (o conteúdo dos
 * anexos passa pelos leitores de arquivos, inclusive mensagens encaminhadas como anexo).
 * truncated: a mensagem foi cortada no limite de tamanho (a parte cortada não é lida).
 * attachments: false lista os anexos sem ler o conteúdo.
 * omittedToken: código das partes que o conector não baixou (anexos de mensagens acima do limite).
 */
export async function extractMessage(buf, { limits = DEFAULT_LIMITS, depth = 0, truncated = false, attachments = true, omittedToken = null } = {}) {
  const parsed = parseMime(buf, { truncated, decodeAttachments: attachments, omittedToken });
  let body = parsed.texts
    .map((t) => t.text.trim())
    .filter(Boolean)
    .join('\n\n');
  const maxChars = limits.maxChars ?? DEFAULT_LIMITS.maxChars;
  if (body.length > maxChars) body = body.slice(0, maxChars);
  const list = attachments
    ? await readAttachments(parsed.attachments, { limits: { ...DEFAULT_LIMITS, ...limits, maxChars }, depth, used: body.length })
    : parsed.attachments.map((a) => ({ name: a.name, size: a.size, contentType: a.contentType, inline: a.inline, type: null, status: 'not-requested', note: null, metadata: {}, segments: [] }));
  return {
    subject: parsed.subject,
    from: parsed.from,
    sender: parsed.sender,
    to: parsed.to,
    cc: parsed.cc,
    bcc: parsed.bcc,
    replyTo: parsed.replyTo,
    date: parsed.date,
    messageId: parsed.messageId,
    body,
    attachments: list,
    encrypted: parsed.encrypted,
    opaqueSigned: parsed.opaqueSigned,
    partial: truncated || parsed.texts.some((t) => t.incomplete),
  };
}

/** Arquivo .eml/.mht (ou mensagem anexada) como uma lista única de trechos de texto. */
async function extractMime(buf, { ext, limits, withText, partial, depth }) {
  const web = WEB_ARCHIVE_EXT.has(ext);
  const m = await extractMessage(buf, { limits, depth, truncated: partial, attachments: withText && !web });
  const addresses = (list) => list.map(formatAddress).join(', ');
  const header = [
    ['Assunto', m.subject],
    ['De', addresses(m.from)],
    ['Para', addresses(m.to)],
    ['Cc', addresses(m.cc)],
  ]
    .filter(([, v]) => v)
    .map(([k, v]) => `${k}: ${v}`)
    .join('\n');
  const listed = web ? [] : m.attachments;
  const segments = [];
  if (header) segments.push({ label: 'Cabeçalho', text: header });
  if (m.body) segments.push({ label: 'Mensagem', text: m.body });
  if (listed.length) segments.push({ label: 'Anexos', text: listed.map((a) => a.name).join('\n') });
  for (const att of listed) segments.push(...att.segments);
  const metadata = compact({ author: formatAddress(m.from[0]) || null, title: m.subject || null, created: m.date });
  const unreadable = (m.encrypted || m.opaqueSigned) && !m.body;
  return {
    type: ext.slice(1) || 'eml',
    metadata,
    segments: withText ? segments : [],
    encrypted: unreadable,
    partial: m.partial,
    attachments: listed.map(attachmentSummary),
  };
}

const bigNote = (partial) => (partial ? 'Arquivo grande: apenas o início foi analisado.' : null);

/** Extrai o conteúdo já carregado em memória, conforme o formato identificado. */
async function extractLoaded(kind, buf, { ext, limits, withText, partial, depth, mime }) {
  if (kind === 'rtf') {
    const { text, metadata } = rtfExtract(buf);
    return finish({ type: 'rtf', metadata, segments: withText ? [{ text }] : [] }, limits, bigNote(partial));
  }
  if (mime) return finish(await extractMime(buf, { ext, limits, withText, partial, depth }), limits, bigNote(partial));
  if (kind === 'html' || kind === 'text') {
    if (!withText) return { type: kind === 'html' ? 'html' : 'texto', status: 'ok', segments: [], metadata: {} };
    const text = decodeText(buf, { partial });
    const result = kind === 'html' ? { type: 'html', segments: [{ text: htmlToText(text) }] } : { type: 'texto', segments: [{ text, lines: true }] };
    return finish(result, limits, bigNote(partial));
  }
  if (kind === 'zip') return finish(await extractZip(buf, { withText }), limits);
  if (kind === 'ole') {
    const r = oleExtract(buf, { withText });
    const type = r.kind === 'encrypted' || r.kind === 'ole' ? ext.slice(1) || 'ole' : r.kind;
    const result = { type, metadata: r.metadata, segments: r.segments, encrypted: r.encrypted, unsupported: r.unsupported };
    if (withText && r.attachments?.length) {
      const infos = await readAttachments(r.attachments, { limits, depth, used: textLength(r.segments) });
      for (const info of infos) result.segments.push(...info.segments);
      result.attachments = infos.map(attachmentSummary);
    }
    return finish(result, limits);
  }
  if (kind === 'pdf') {
    const r = await pdfExtract(buf, { maxChars: limits.maxChars, withText });
    return finish({ type: 'pdf', metadata: r.metadata, segments: r.segments, encrypted: r.encrypted, partial: r.truncated }, limits);
  }
  return { type: kind, status: 'unsupported', segments: [], metadata: {} };
}

/**
 * Extrai o conteúdo (withText=true) ou apenas os metadados de um arquivo.
 * size: tamanho já obtido pelo stat (evita outra chamada).
 */
export async function extractFile(filePath, { size, limits = DEFAULT_LIMITS, withText = true } = {}) {
  const ext = path.extname(filePath).toLowerCase();
  if (size === 0) return { type: ext.slice(1) || 'arquivo', status: 'empty', segments: [], metadata: {} };
  limits = { ...DEFAULT_LIMITS, ...limits };
  const tooBig = size > limits.maxBytes;
  let kind = 'binary';
  try {
    kind = classify(await readHead(filePath, 8192), ext);
    if (kind === 'binary') return { type: ext.slice(1) || 'binário', status: 'unsupported', segments: [], metadata: {} };

    if (kind === 'text' || kind === 'html' || kind === 'rtf') {
      const mime = kind !== 'rtf' && MIME_EXTENSIONS.has(ext);
      if (!withText && kind !== 'rtf' && !mime) return { type: kind === 'html' ? 'html' : 'texto', status: 'ok', segments: [], metadata: {} };
      const partial = tooBig && withText;
      const buf = partial || !withText ? await readHead(filePath, withText ? limits.maxBytes : 256 * 1024) : await fs.readFile(filePath);
      return await extractLoaded(kind, buf, { ext, limits, withText, partial, depth: 0, mime });
    }

    // Formatos estruturados precisam do arquivo inteiro. Só para os metadados o limite é maior.
    if (withText ? tooBig : size > limits.maxMetadataBytes) {
      return {
        type: kind,
        status: 'skipped-size',
        segments: [],
        metadata: {},
        note: `Conteúdo não analisado: arquivo maior que ${Math.round(limits.maxBytes / 1048576)} MB.`,
      };
    }
    return await extractLoaded(kind, await fs.readFile(filePath), { ext, limits, withText, partial: false, depth: 0 });
  } catch (err) {
    // Sem permissão de leitura, arquivo bloqueado etc.: o nome continua sendo verificado.
    return { type: ext.slice(1) || kind, status: 'error', segments: [], metadata: {}, note: `Falha ao ler o conteúdo: ${friendlyError(err)}` };
  }
}

/**
 * Extrai o conteúdo de um arquivo já em memória (ex.: anexo de e-mail). name dá a extensão;
 * mime=true trata o conteúdo como mensagem de e-mail (anexo do tipo message/rfc822).
 */
export async function extractBuffer(buf, { name = '', limits = DEFAULT_LIMITS, withText = true, depth = 0, mime = false } = {}) {
  let ext = path.extname(name).toLowerCase();
  if (!buf || buf.length === 0) return { type: ext.slice(1) || 'arquivo', status: 'empty', segments: [], metadata: {} };
  limits = { ...DEFAULT_LIMITS, ...limits };
  let kind = 'binary';
  try {
    kind = mime ? 'text' : classify(buf.subarray(0, 8192), ext);
    if (kind === 'binary') return { type: ext.slice(1) || 'binário', status: 'unsupported', segments: [], metadata: {} };
    if (kind === 'text' || kind === 'html' || kind === 'rtf') {
      const isMime = mime || (kind !== 'rtf' && MIME_EXTENSIONS.has(ext));
      if (mime && !MIME_EXTENSIONS.has(ext)) ext = '.eml';
      const partial = withText && buf.length > limits.maxBytes;
      return await extractLoaded(kind, partial ? buf.subarray(0, limits.maxBytes) : buf, { ext, limits, withText, partial, depth, mime: isMime });
    }
    if (withText && buf.length > limits.maxBytes) {
      return { type: kind, status: 'skipped-size', segments: [], metadata: {}, note: `Conteúdo não analisado: maior que ${Math.round(limits.maxBytes / 1048576)} MB.` };
    }
    return await extractLoaded(kind, buf, { ext, limits, withText, partial: false, depth });
  } catch (err) {
    return { type: ext.slice(1) || kind, status: 'error', segments: [], metadata: {}, note: `Falha ao ler o conteúdo: ${friendlyError(err)}` };
  }
}

function finish(result, limits, note = null) {
  const out = { type: result.type, metadata: result.metadata || {}, segments: [], status: 'ok', note };
  if (result.attachments?.length) out.attachments = result.attachments;
  if (result.encrypted) {
    out.status = 'encrypted';
    out.note = 'Arquivo protegido por senha: conteúdo não analisado.';
    return out;
  }
  if (result.unsupported) {
    out.status = 'unsupported';
    return out;
  }
  const { segments, truncated } = limitSegments(result.segments || [], limits.maxChars);
  out.segments = segments;
  if (truncated || result.partial || note) {
    out.status = 'partial';
    out.note ||= 'Conteúdo muito extenso: apenas parte foi analisada.';
  }
  return out;
}
