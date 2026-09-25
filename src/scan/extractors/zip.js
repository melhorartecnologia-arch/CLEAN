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
