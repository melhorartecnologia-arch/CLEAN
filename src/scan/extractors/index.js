// Identifica o formato real do arquivo (pelos primeiros bytes e pela extensão) e extrai texto e
// metadados. Resultado: { type, status, segments, metadata, note }.
//   status: ok | partial | empty | unsupported | encrypted | skipped-size | error
import fs from 'node:fs/promises';
import path from 'node:path';
import { ZipReader } from './zip.js';
import { detectOoxml, ooxmlMetadata, docxSegments, xlsxSegments, pptxSegments, naturalCompare } from './ooxml.js';
import { detectOdf, odfIsEncrypted, odfMetadata, odfSegments } from './odf.js';
import { isOle, oleExtract } from './ole.js';
import { isPdf, pdfExtract } from './pdf.js';
import { isRtf, rtfExtract } from './rtf.js';
import { decodeText, looksLikeText } from './text.js';
import { htmlToText } from './xml.js';
import { MIME_EXTENSIONS, mimeExtract } from './mime.js';
import { zipEntryNames } from './zip.js';
import { friendlyError } from '../errors.js';

export const DEFAULT_LIMITS = {
  maxBytes: 50 * 1024 * 1024, // arquivos maiores: só o nome é analisado (texto puro: lê o início)
  maxChars: 20_000_000, // texto máximo extraído por arquivo
  maxMetadataBytes: 200 * 1024 * 1024, // limite para ler apenas os metadados (último usuário)
};

const HTML_EXT = new Set(['.htm', '.html', '.xhtml', '.mht', '.mhtml', '.hta']);
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

/** Classifica o arquivo pelos primeiros bytes (e pela extensão, em último caso). */
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

/**
 * Extrai o conteúdo (withText=true) ou apenas os metadados de um arquivo.
 * size: tamanho já obtido pelo stat (evita outra chamada).
 */
export async function extractFile(filePath, { size, limits = DEFAULT_LIMITS, withText = true } = {}) {
  const ext = path.extname(filePath).toLowerCase();
  if (size === 0) return { type: ext.slice(1) || 'arquivo', status: 'empty', segments: [], metadata: {} };
  const maxBytes = limits.maxBytes ?? DEFAULT_LIMITS.maxBytes;
  const maxMetadataBytes = limits.maxMetadataBytes ?? DEFAULT_LIMITS.maxMetadataBytes;
  const tooBig = size > maxBytes;
  let kind = 'binary';
  try {
    kind = classify(await readHead(filePath, 8192), ext);
    if (kind === 'binary') return { type: ext.slice(1) || 'binário', status: 'unsupported', segments: [], metadata: {} };

    if (kind === 'text' || kind === 'html' || kind === 'rtf') {
      const mime = kind !== 'rtf' && MIME_EXTENSIONS.has(ext);
      if (!withText && kind !== 'rtf' && !mime) return { type: kind === 'html' ? 'html' : 'texto', status: 'ok', segments: [], metadata: {} };
      const partial = tooBig && withText;
      const buf = partial || !withText ? await readHead(filePath, withText ? maxBytes : 256 * 1024) : await fs.readFile(filePath);
      let result;
      if (kind === 'rtf') {
        const { text, metadata } = rtfExtract(buf);
        result = { type: 'rtf', metadata, segments: withText ? [{ text }] : [] };
      } else if (mime) {
        const r = mimeExtract(buf);
        result = { type: ext.slice(1), metadata: r.metadata, segments: withText ? r.segments : [] };
      } else if (kind === 'html') {
        result = { type: 'html', metadata: {}, segments: [{ text: htmlToText(decodeText(buf, { partial })) }] };
      } else {
        result = { type: 'texto', metadata: {}, segments: [{ text: decodeText(buf, { partial }), lines: true }] };
      }
      return finish(result, limits, partial ? 'Arquivo grande: apenas o início foi analisado.' : null);
    }

    // Formatos estruturados precisam do arquivo inteiro. Só para os metadados o limite é maior.
    if (withText ? tooBig : size > maxMetadataBytes) {
      return {
        type: kind,
        status: 'skipped-size',
        segments: [],
        metadata: {},
        note: `Conteúdo não analisado: arquivo maior que ${Math.round(maxBytes / 1048576)} MB.`,
      };
    }

    const buf = await fs.readFile(filePath);
    if (kind === 'zip') return finish(await extractZip(buf, { withText }), limits);
    if (kind === 'ole') {
      const r = oleExtract(buf, { withText });
      const type = r.kind === 'encrypted' || r.kind === 'ole' ? ext.slice(1) || 'ole' : r.kind;
      return finish({ type, metadata: r.metadata, segments: r.segments, encrypted: r.encrypted, unsupported: r.unsupported }, limits);
    }
    if (kind === 'pdf') {
      const r = await pdfExtract(buf, { maxChars: limits.maxChars, withText });
      return finish({ type: 'pdf', metadata: r.metadata, segments: r.segments, encrypted: r.encrypted, partial: r.truncated }, limits);
    }
    return { type: kind, status: 'unsupported', segments: [], metadata: {} };
  } catch (err) {
    // Sem permissão de leitura, arquivo bloqueado etc.: o nome continua sendo verificado.
    return { type: ext.slice(1) || kind, status: 'error', segments: [], metadata: {}, note: `Falha ao ler o conteúdo: ${friendlyError(err)}` };
  }
}

function finish(result, limits, note = null) {
  const out = { type: result.type, metadata: result.metadata || {}, segments: [], status: 'ok', note };
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
