// Exclusão dos arquivos encontrados (automática, durante a análise, ou manual, pelo relatório).
// A exclusão é definitiva: arquivos apagados por programas (inclusive em compartilhamentos de rede)
// não vão para a Lixeira do Windows.
import fs from 'node:fs/promises';
import path from 'node:path';
import { friendlyError } from './errors.js';

/** Indica se `target` está dentro da pasta `root` (sem sair dela com "..", nem em outra unidade). */
export function isInside(root, target) {
  if (!root || !target) return false;
  const rel = path.relative(root, target);
  return Boolean(rel) && !rel.startsWith('..') && !path.isAbsolute(rel);
}

/**
 * Exclui um arquivo. Retorna { status: 'deleted' | 'missing' | 'changed' | 'failed', error? }.
 * root: pasta do repositório (o arquivo precisa estar dentro dela).
 * expected: { size, modified } registrados na análise; se o arquivo mudou depois disso, a exclusão
 * só acontece com force = true (status 'changed').
 */
export async function deleteFile(filePath, { root, expected = null, force = false } = {}) {
  if (!isInside(root, filePath)) return { status: 'failed', error: 'O arquivo não está dentro da pasta do repositório.' };
  let st;
  try {
    st = await fs.lstat(filePath);
  } catch (err) {
    if (err.code === 'ENOENT') return { status: 'missing', error: 'O arquivo não existe mais (já excluído ou movido).' };
    return { status: 'failed', error: friendlyError(err) };
  }
  if (!st.isFile()) return { status: 'failed', error: 'O caminho não é um arquivo.' };
  if (expected && !force) {
    const modified = expected.modified ? Date.parse(expected.modified) : NaN;
    const sameSize = expected.size === undefined || st.size === expected.size;
    const sameDate = Number.isNaN(modified) || Math.abs(st.mtimeMs - modified) < 2000;
    if (!sameSize || !sameDate) {
      return { status: 'changed', error: 'O arquivo foi alterado depois da análise (tamanho ou data de modificação diferentes).' };
    }
  }
  try {
    await fs.unlink(filePath);
  } catch (err) {
    // No Windows, arquivos com o atributo "somente leitura" só podem ser excluídos sem ele.
    if (err.code === 'EPERM' && process.platform === 'win32') {
      try {
        await fs.chmod(filePath, 0o666);
        await fs.unlink(filePath);
        return { status: 'deleted' };
      } catch (retry) {
        return { status: 'failed', error: friendlyError(retry) };
      }
    }
    if (err.code === 'ENOENT') return { status: 'missing', error: 'O arquivo não existe mais (já excluído ou movido).' };
    return { status: 'failed', error: friendlyError(err) };
  }
  return { status: 'deleted' };
}

/** Evento de exclusão gravado no registro da análise (deletions.ndjson). */
export function deletionEvent(recordId, result, { mode, method, by }) {
  return {
    recordId,
    status: result.status,
    error: result.error || null,
    mode, // 'auto' (durante a análise) ou 'manual' (pelo relatório)
    method, // 'file', 'permanent' ou 'trash'
    by,
    at: new Date().toISOString(),
  };
}
