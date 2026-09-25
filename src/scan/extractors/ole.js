// Formatos binários do Office 97-2003 (OLE / Compound File): .doc, .xls, .ppt e e-mails .msg do Outlook.
// Também lê o fluxo SummaryInformation, que guarda o autor e o "salvo por último por".
import CFB from 'cfb';
import { decoderFor, codePageLabel, binaryStrings, decodeText } from './text.js';
import { htmlToText } from './xml.js';
import { compact } from './ooxml.js';

export function isOle(buf) {
  return buf.length >= 512 && buf.readUInt32LE(0) === 0xe011cfd0 && buf.readUInt32LE(4) === 0xe11ab1a1;
}

export function openOle(buf) {
  return CFB.read(buf, { type: 'buffer' });
}

function stream(cfb, name) {
  const entry = CFB.find(cfb, name);
  if (!entry || entry.type !== 2 || !entry.content || !entry.content.length) return null;
  return Buffer.isBuffer(entry.content) ? entry.content : Buffer.from(entry.content);
}

/** Identifica o conteúdo do contêiner OLE. */
export function detectOle(cfb) {
  const names = new Set(cfb.FileIndex.map((e) => e.name));
  if (names.has('EncryptedPackage') || names.has('EncryptionInfo')) return 'encrypted';
  if (names.has('WordDocument')) return 'doc';
  if (names.has('Workbook') || names.has('Book')) return 'xls';
  if (names.has('PowerPoint Document')) return 'ppt';
  if ([...names].some((n) => n.startsWith('__substg1.0_'))) return 'msg';
  return 'ole';
}

// ---------------------------------------------------------------------------------------------
// Conjuntos de propriedades (SummaryInformation / DocumentSummaryInformation) — [MS-OLEPS]

const FILETIME_EPOCH_DIFF = 11644473600000;

export function parsePropertySet(buf) {
  const props = new Map();
  if (!buf || buf.length < 48 || buf.readUInt16LE(0) !== 0xfffe || buf.readUInt32LE(24) < 1) return props;
  const base = buf.readUInt32LE(44);
  if (base + 8 > buf.length) return props;
  const count = Math.min(buf.readUInt32LE(base + 4), 1000);
  const entries = [];
  for (let k = 0; k < count; k++) {
    const p = base + 8 + k * 8;
    if (p + 8 > buf.length) break;
    entries.push([buf.readUInt32LE(p), base + buf.readUInt32LE(p + 4)]);
  }
  let codePage = 1252;
  for (const [id, off] of entries) {
    if (id === 1 && off + 6 <= buf.length && buf.readUInt16LE(off) === 2) codePage = buf.readUInt16LE(off + 4);
  }
  for (const [id, off] of entries) {
    if (off + 8 > buf.length) continue;
    const type = buf.readUInt16LE(off);
    let value = null;
    if (type === 0x1e) {
      const size = buf.readUInt32LE(off + 4);
      const bytes = buf.subarray(off + 8, Math.min(buf.length, off + 8 + size));
      value = codePage === 1200 ? bytes.toString('utf16le') : decoderFor(codePageLabel(codePage)).decode(bytes);
    } else if (type === 0x1f) {
      const chars = buf.readUInt32LE(off + 4);
      value = buf.toString('utf16le', off + 8, Math.min(buf.length, off + 8 + chars * 2));
    } else if (type === 0x40 && off + 12 <= buf.length) {
      const ticks = buf.readUInt32LE(off + 8) * 2 ** 32 + buf.readUInt32LE(off + 4);
      const ms = ticks / 10000 - FILETIME_EPOCH_DIFF;
      if (ms > Date.UTC(1980, 0, 1) && ms < Date.UTC(2200, 0, 1)) value = new Date(ms).toISOString();
    }
    if (typeof value === 'string') value = value.replace(/\0+$/g, '').trim();
    if (value) props.set(id, value);
  }
  return props;
}

export function oleMetadata(cfb) {
  const si = parsePropertySet(stream(cfb, '\u0005SummaryInformation'));
  const dsi = parsePropertySet(stream(cfb, '\u0005DocumentSummaryInformation'));
  return compact({
    title: si.get(2),
    author: si.get(4),
    lastModifiedBy: si.get(8),
    created: si.get(12),
    modified: si.get(13),
    application: si.get(18),
    company: dsi.get(15),
  });
}

// ---------------------------------------------------------------------------------------------
// Word 97-2003: tabela de peças (CLX) do fluxo WordDocument — [MS-DOC]

function wordText(cfb) {
  const wd = stream(cfb, 'WordDocument');
  if (!wd || wd.length < 0x1a0) return null;
  const ident = wd.readUInt16LE(0);
  if (ident !== 0xa5ec) return { text: binaryStrings(wd) }; // Word 6/95 ou desconhecido
  const flags = wd.readUInt16LE(0x0a);
  if (flags & 0x0100) return { encrypted: true };
  const table = stream(cfb, flags & 0x0200 ? '1Table' : '0Table');
  if (!table) return { text: binaryStrings(wd) };
  let pos = 32;
  pos += 2 + wd.readUInt16LE(pos) * 2; // csw + fibRgW
  pos += 2 + wd.readUInt16LE(pos) * 4; // cslw + fibRgLw
  const cbRgFcLcb = wd.readUInt16LE(pos);
  pos += 2;
  if (cbRgFcLcb < 34 || pos + 34 * 8 > wd.length) return { text: binaryStrings(wd) };
  const fcClx = wd.readUInt32LE(pos + 33 * 8);
  const lcbClx = wd.readUInt32LE(pos + 33 * 8 + 4);
  if (!lcbClx || fcClx + lcbClx > table.length) return { text: binaryStrings(wd) };

  let p = fcClx;
  const end = fcClx + lcbClx;
  while (p < end && table[p] === 0x01) p += 3 + table.readUInt16LE(p + 1); // Prc (ignorado)
  if (p + 5 > end || table[p] !== 0x02) return { text: binaryStrings(wd) };
  const lcb = table.readUInt32LE(p + 1);
  const plc = p + 5;
  const n = (lcb - 4) / 12;
  if (!Number.isInteger(n) || n <= 0 || plc + lcb > table.length) return { text: binaryStrings(wd) };
  const cp1252 = decoderFor('windows-1252');
  const pieces = [];
  for (let k = 0; k < n; k++) {
    const count = table.readUInt32LE(plc + (k + 1) * 4) - table.readUInt32LE(plc + k * 4);
    if (count <= 0) continue;
    const fcRaw = table.readUInt32LE(plc + (n + 1) * 4 + k * 8 + 2);
    const fc = fcRaw & 0x3fffffff;
    if (fcRaw & 0x40000000) {
      const start = fc / 2;
      pieces.push(cp1252.decode(wd.subarray(start, Math.min(wd.length, start + count))));
    } else {
      pieces.push(wd.toString('utf16le', fc, Math.min(wd.length, fc + count * 2)));
    }
  }
  return { text: cleanWordText(pieces.join('')) };
}

/** Converte os caracteres especiais do Word e remove instruções de campos (mantém o resultado). */
function cleanWordText(text) {
  let out = '';
  const fields = []; // true = lendo instrução do campo
  for (let i = 0; i < text.length; i++) {
    const c = text.charCodeAt(i);
    if (c === 0x13) {
      fields.push(true);
      continue;
    }
    if (c === 0x14) {
      if (fields.length) fields[fields.length - 1] = false;
      continue;
    }
    if (c === 0x15) {
      fields.pop();
      continue;
    }
    if (fields.length && fields.includes(true)) continue;
    if (c === 0x0d || c === 0x0b || c === 0x0c || c === 0x0e) out += '\n';
    else if (c === 0x07) out += '\t';
    else if (c === 0x1e) out += '-';
    else if (c === 0x09 || c >= 0x20) out += text[i];
  }
  return out;
}

// ---------------------------------------------------------------------------------------------
// Excel 97-2003 (BIFF8) — [MS-XLS]

/** Leitor que atravessa registros CONTINUE (necessário para a tabela de strings SST). */
class ChunkReader {
  constructor(chunks) {
    this.chunks = chunks;
    this.ci = 0;
    this.pos = 0;
  }

  #ensure() {
    while (this.ci < this.chunks.length && this.pos >= this.chunks[this.ci].length) {
      this.ci++;
      this.pos = 0;
    }
    if (this.ci >= this.chunks.length) throw new RangeError('fim dos dados');
  }

  u8() {
    this.#ensure();
    return this.chunks[this.ci][this.pos++];
  }

  u16() {
    return this.u8() | (this.u8() << 8);
  }

  u32() {
    return (this.u16() | (this.u16() << 16)) >>> 0;
  }

  skip(n) {
    while (n > 0) {
      this.#ensure();
      const take = Math.min(n, this.chunks[this.ci].length - this.pos);
      this.pos += take;
      n -= take;
    }
  }

  chars(cch, high) {
    let out = '';
    while (cch > 0) {
      if (this.pos >= this.chunks[this.ci].length) {
        // A string continua no próximo CONTINUE, que começa com um novo byte de opções.
        this.ci++;
        this.pos = 0;
        if (this.ci >= this.chunks.length) throw new RangeError('fim dos dados');
        high = (this.chunks[this.ci][this.pos++] & 1) === 1;
      }
      const chunk = this.chunks[this.ci];
      const avail = high ? (chunk.length - this.pos) >> 1 : chunk.length - this.pos;
      const take = Math.min(cch, avail);
      if (take === 0) {
        this.pos = chunk.length;
        continue;
      }
      const bytes = high ? take * 2 : take;
      out += chunk.toString(high ? 'utf16le' : 'latin1', this.pos, this.pos + bytes);
      this.pos += bytes;
      cch -= take;
    }
    return out;
  }
}

function readSst(chunks) {
  const r = new ChunkReader(chunks);
  const list = [];
  try {
    r.u32(); // total
    const unique = r.u32();
    for (let k = 0; k < unique; k++) {
      const cch = r.u16();
      const flags = r.u8();
      const runs = flags & 0x08 ? r.u16() : 0;
      const ext = flags & 0x04 ? r.u32() : 0;
      list.push(r.chars(cch, (flags & 0x01) === 1));
      if (runs) r.skip(runs * 4);
      if (ext) r.skip(ext);
    }
  } catch {
    // SST truncada: mantém o que foi lido
  }
  return list;
}

function xlString(data, offset) {
  if (offset + 3 > data.length) return '';
  const cch = data.readUInt16LE(offset);
  const high = (data[offset + 2] & 1) === 1;
  const start = offset + 3;
  return high
    ? data.toString('utf16le', start, Math.min(data.length, start + cch * 2))
    : data.toString('latin1', start, Math.min(data.length, start + cch));
}

const rkBuf = Buffer.alloc(8);
function rkNumber(rk) {
  let num;
  if (rk & 0x02) {
    num = (rk | 0) >> 2;
  } else {
    rkBuf.writeUInt32LE(0, 0);
    rkBuf.writeUInt32LE((rk & 0xfffffffc) >>> 0, 4);
    num = rkBuf.readDoubleLE(0);
  }
  return rk & 0x01 ? num / 100 : num;
}

function excelText(cfb) {
  const wb = stream(cfb, 'Workbook') || stream(cfb, 'Book');
  if (!wb || wb.length < 8) return null;
  if (wb.readUInt16LE(0) !== 0x0809 || wb.readUInt16LE(4) !== 0x0600) return { text: binaryStrings(wb) }; // BIFF5/7
  const sheetNames = new Map();
  let sst = [];
  const segments = [];
  let sheet = null;
  let p = 0;

  const put = (row, value) => {
    if (!sheet || value === '' || value === null || value === undefined) return;
    const line = row + 1;
    if (line > sheet.line) {
      if (line - sheet.line <= 20000) {
        sheet.parts.push('\n'.repeat(line - sheet.line));
      } else {
        sheet.parts.push('\n');
        sheet.exact = false;
      }
      sheet.line = line;
      sheet.first = true;
    }
    sheet.parts.push(sheet.first ? String(value) : `\t${value}`);
    sheet.first = false;
  };

  while (p + 4 <= wb.length) {
    const type = wb.readUInt16LE(p);
    const len = wb.readUInt16LE(p + 2);
    const data = wb.subarray(p + 4, p + 4 + len);
    const recordStart = p;
    p += 4 + len;
    switch (type) {
      case 0x0809: // BOF
        if (sheetNames.has(recordStart)) sheet = { name: sheetNames.get(recordStart), parts: [], line: 1, first: true, exact: true };
        break;
      case 0x000a: // EOF
        if (sheet) {
          segments.push({ label: `Planilha "${sheet.name}"`, text: sheet.parts.join(''), lines: sheet.exact });
          sheet = null;
        }
        break;
      case 0x002f: // FILEPASS: pasta de trabalho criptografada
        if (!sheet) return { encrypted: true };
        break;
      case 0x0085: // BOUNDSHEET
        if (data.length >= 8) {
          const cch = data[6];
          const high = (data[7] & 1) === 1;
          const name = high ? data.toString('utf16le', 8, 8 + cch * 2) : data.toString('latin1', 8, 8 + cch);
          sheetNames.set(data.readUInt32LE(0), name);
        }
        break;
      case 0x00fc: {
        // SST + CONTINUE
        const chunks = [data];
        while (p + 4 <= wb.length && wb.readUInt16LE(p) === 0x003c) {
          const clen = wb.readUInt16LE(p + 2);
          chunks.push(wb.subarray(p + 4, p + 4 + clen));
          p += 4 + clen;
        }
        sst = readSst(chunks);
        break;
      }
      case 0x00fd: // LABELSST
        if (data.length >= 10) put(data.readUInt16LE(0), sst[data.readUInt32LE(6)]);
        break;
      case 0x0204: // LABEL
        if (data.length >= 9) put(data.readUInt16LE(0), xlString(data, 6));
        break;
      case 0x0203: // NUMBER
        if (data.length >= 14) put(data.readUInt16LE(0), data.readDoubleLE(6));
        break;
      case 0x027e: // RK
        if (data.length >= 10) put(data.readUInt16LE(0), rkNumber(data.readUInt32LE(6)));
        break;
      case 0x00bd: {
        // MULRK
        if (data.length < 12) break;
        const row = data.readUInt16LE(0);
        for (let o = 4; o + 6 <= data.length - 2; o += 6) put(row, rkNumber(data.readUInt32LE(o + 2)));
        break;
      }
      case 0x0006: // FORMULA (valor numérico; strings vêm no registro STRING seguinte)
        if (data.length >= 14) {
          const row = data.readUInt16LE(0);
          if (data.readUInt16LE(12) !== 0xffff) put(row, data.readDoubleLE(6));
          else if (data[6] === 0 && sheet) sheet.pendingRow = row;
        }
        break;
      case 0x0207: // STRING (resultado de fórmula)
        if (sheet && sheet.pendingRow !== undefined) {
          put(sheet.pendingRow, xlString(data, 0));
          sheet.pendingRow = undefined;
        }
        break;
      default:
        break;
    }
  }
  return { segments };
}

// ---------------------------------------------------------------------------------------------
// PowerPoint 97-2003 — [MS-PPT]

function powerPointText(cfb) {
  if (CFB.find(cfb, 'EncryptedSummary')) return { encrypted: true };
  const doc = stream(cfb, 'PowerPoint Document');
  if (!doc) return null;
  const parts = [];
  let p = 0;
  while (p + 8 <= doc.length) {
    const recVer = doc.readUInt16LE(p) & 0x0f;
    const type = doc.readUInt16LE(p + 2);
    const len = doc.readUInt32LE(p + 4);
    if (type === 0x03f8 || type === 0x0fc9) {
      p += 8 + len; // slide-mestre e folheto: só textos de modelo ("Clique para editar...")
      continue;
    }
    if (recVer === 0x0f) {
      p += 8; // contêiner: entra nos registros filhos
      continue;
    }
    const start = p + 8;
    const end = Math.min(doc.length, start + len);
    if (type === 0x0fa0) parts.push(doc.toString('utf16le', start, end - ((end - start) % 2)));
    else if (type === 0x0fa8) parts.push(doc.toString('latin1', start, end));
    p = start + len;
  }
  return { text: parts.join('\n').replace(/[\r\v]/g, '\n') };
}

// ---------------------------------------------------------------------------------------------
// Outlook .msg — [MS-OXMSG]

const MSG_SKIP = new Set(['007D', '1035', '1042', '0E1D', '0064', '0C1E', '3002', '0070', '5D0A', '5D0B']);

function msgText(cfb) {
  const parts = [];
  const metadata = {};
  let hasBody = false;
  let htmlBody = null;
  cfb.FileIndex.forEach((entry, idx) => {
    const m = /^__substg1\.0_([0-9A-F]{4})([0-9A-F]{4})$/i.exec(entry.name);
    if (!m || entry.type !== 2 || !entry.content || !entry.content.length) return;
    const prop = m[1].toUpperCase();
    const type = m[2].toUpperCase();
    const content = Buffer.from(entry.content);
    if (prop === '1013' && type === '0102') {
      htmlBody = content;
      return;
    }
    if (MSG_SKIP.has(prop) || (type !== '001F' && type !== '001E')) return;
    const value = (type === '001F' ? content.toString('utf16le') : decoderFor('windows-1252').decode(content)).replace(/\0+$/, '');
    if (!value.trim()) return;
    if (prop === '1000') hasBody = true;
    parts.push(value);
    const full = cfb.FullPaths[idx] || '';
    if (!/__(attach|recip)_version1\.0_/i.test(full)) {
      if (prop === '0C1A') metadata.author ??= value.trim();
      if (prop === '3FFA') metadata.lastModifiedBy ??= value.trim();
      if (prop === '0037') metadata.title ??= value.trim();
    }
  });
  if (!hasBody && htmlBody) parts.push(htmlToText(decodeText(htmlBody)));
  return { text: parts.join('\n'), metadata };
}

// ---------------------------------------------------------------------------------------------

/**
 * Extrai texto e metadados de um arquivo OLE.
 * Retorna { kind, segments, metadata, encrypted? }.
 */
export function oleExtract(buf, { withText = true } = {}) {
  const cfb = openOle(buf);
  const kind = detectOle(cfb);
  if (kind === 'encrypted') return { kind, segments: [], metadata: {}, encrypted: true };
  let metadata = oleMetadata(cfb);
  if (!withText) {
    if (kind === 'msg') metadata = { ...msgText(cfb).metadata, ...metadata };
    return { kind, segments: [], metadata };
  }
  const parsers = { doc: wordText, xls: excelText, ppt: powerPointText, msg: msgText };
  const mainStream = { doc: 'WordDocument', xls: 'Workbook', ppt: 'PowerPoint Document' };
  let result = null;
  if (parsers[kind]) {
    try {
      result = parsers[kind](cfb);
    } catch {
      // Estrutura inesperada: recorre à extração de sequências legíveis do fluxo principal.
      const main = mainStream[kind] && (stream(cfb, mainStream[kind]) || stream(cfb, 'Book'));
      result = main ? { text: binaryStrings(main) } : null;
    }
  }
  if (!result) return { kind, segments: [], metadata, unsupported: true };
  if (result.encrypted) return { kind, segments: [], metadata, encrypted: true };
  if (result.metadata) metadata = { ...result.metadata, ...metadata };
  const segments = result.segments || [{ text: result.text || '' }];
  return { kind, segments, metadata };
}
