// Leitura de arquivos ZIP (base dos formatos Office Open XML e OpenDocument) com limites de tamanho.
import path from 'node:path';
import { unzipSync, strFromU8 } from 'fflate';

export const ZIP_LIMITS = {
  maxEntryBytes: 120 * 1024 * 1024, // tamanho descompactado máximo de uma parte
  maxTotalBytes: 400 * 1024 * 1024, // total descompactado por arquivo
};

export class ZipReader {
  constructor(buf, limits = ZIP_LIMITS) {
    this.buf = buf;
    this.limits = limits;
    this.total = 0;
    this.truncated = false;
    this.entries = new Map();
    unzipSync(buf, {
      filter: (f) => {
        this.entries.set(f.name, f.originalSize);
        return false;
      },
    });
  }

  has(name) {
    return this.entries.has(name);
  }

  names() {
    return [...this.entries.keys()];
  }

  /** Conteúdo de uma parte como texto UTF-8 (null se não existir ou exceder os limites). */
  text(name) {
    const bytes = this.bytes(name);
    return bytes ? strFromU8(bytes) : null;
  }

  bytes(name) {
    const size = this.entries.get(name);
    if (size === undefined) return null;
    if (size > this.limits.maxEntryBytes || this.total + size > this.limits.maxTotalBytes) {
      this.truncated = true;
      return null;
    }
    this.total += size;
    const out = unzipSync(this.buf, { filter: (f) => f.name === name });
    return out[name] ?? null;
  }
}

// Página de código 850 (OEM do Windows em português), usada pela "Pasta compactada" do Explorer
// nos nomes de arquivos, sem o indicador de UTF-8.
const CP850 =
  'ÇüéâäàåçêëèïîìÄÅÉæÆôöòûùÿÖÜø£Ø×ƒáíóúñÑªº¿®¬½¼¡«»░▒▓│┤ÁÂÀ©╣║╗╝¢¥┐└┴┬├─┼ãÃ╚╔╩╦╠═╬¤ðÐÊËÈıÍÎÏ┘┌█▄¦Ì▀ÓßÔÒõÕµþÞÚÛÙýÝ¯´\u00ad±‗¾¶§÷¸°¨·¹³²■\u00a0';

export function decodeCp850(bytes) {
  let out = '';
  for (const b of bytes) out += b < 0x80 ? String.fromCharCode(b) : CP850[b - 0x80];
  return out;
}

const utf8Strict = new TextDecoder('utf-8', { fatal: true });

function decodeEntryName(raw, flags, extra) {
  if (flags & 0x800) return raw.toString('utf8');
  // Campo extra "Info-ZIP Unicode Path" (0x7075), gravado por 7-Zip, WinZip e outros
  for (let p = 0; p + 4 <= extra.length; ) {
    const id = extra.readUInt16LE(p);
    const size = extra.readUInt16LE(p + 2);
    if (id === 0x7075 && size > 5 && extra[p + 4] === 1) return extra.toString('utf8', p + 9, p + 4 + size);
    p += 4 + size;
  }
  if (raw.every((b) => b < 0x80)) return raw.toString('latin1');
  try {
    return utf8Strict.decode(raw);
  } catch {
    return decodeCp850(raw);
  }
}

/**
 * Nomes das entradas lidos diretamente do diretório central, respeitando a codificação usada
 * por quem criou o ZIP. Retorna null se o arquivo não puder ser lido assim (ex.: ZIP64).
 */
export function zipEntryNames(buf) {
  let eocd = -1;
  for (let i = buf.length - 22; i >= Math.max(0, buf.length - 65557); i--) {
    if (buf.readUInt32LE(i) === 0x06054b50) {
      eocd = i;
      break;
    }
  }
  if (eocd === -1) return null;
  const count = buf.readUInt16LE(eocd + 10);
  let p = buf.readUInt32LE(eocd + 16);
  if (count === 0xffff || p === 0xffffffff) return null;
  const names = [];
  for (let k = 0; k < count; k++) {
    if (p + 46 > buf.length || buf.readUInt32LE(p) !== 0x02014b50) return null;
    const flags = buf.readUInt16LE(p + 8);
    const nameLength = buf.readUInt16LE(p + 28);
    const extraLength = buf.readUInt16LE(p + 30);
    const commentLength = buf.readUInt16LE(p + 32);
    const raw = buf.subarray(p + 46, p + 46 + nameLength);
    const extra = buf.subarray(p + 46 + nameLength, p + 46 + nameLength + extraLength);
    names.push(decodeEntryName(raw, flags, extra));
    p += 46 + nameLength + extraLength + commentLength;
  }
  return names;
}

/** Resolve o alvo de um relacionamento (.rels) a partir da pasta da parte de origem. */
export function resolveTarget(baseDir, target) {
  if (target.startsWith('/')) return path.posix.normalize(target.slice(1));
  return path.posix.normalize(path.posix.join(baseDir, target));
}

/** Lê um arquivo .rels e devolve Map(Id -> { target, type }). */
export function readRels(zip, relsPath, baseDir) {
  const xml = zip.text(relsPath);
  const map = new Map();
  if (!xml) return map;
  const re = /<Relationship\b([^>]*)\/?>/g;
  let m;
  while ((m = re.exec(xml)) !== null) {
    const id = /\bId="([^"]*)"/.exec(m[1])?.[1];
    const target = /\bTarget="([^"]*)"/.exec(m[1])?.[1];
    const type = /\bType="([^"]*)"/.exec(m[1])?.[1] || '';
    const external = /\bTargetMode="External"/.test(m[1]);
    if (id && target && !external) map.set(id, { target: resolveTarget(baseDir, target), type });
  }
  return map;
}
