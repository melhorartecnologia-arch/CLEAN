// Exclusão dos arquivos encontrados (automática, durante a análise, ou manual, pelo relatório).
// A exclusão é definitiva: arquivos apagados por programas (inclusive em compartilhamentos de rede)
// não vão para a Lixeira do Windows.
import fs from 'node:fs/promises';
import path from 'node:path';
import { friendlyError } from './errors.js';

const MISSING = 'O arquivo não existe mais (já excluído ou movido).';
const OUTSIDE =
  'Uma pasta no caminho do arquivo agora aponta para fora do repositório (atalho, link simbólico ou junção); por segurança, o arquivo não foi excluído.';

/** Repositório no OneDrive ou no SharePoint (os demais são pastas do Windows). */
export const isCloudRepo = (repo) => repo?.type === 'onedrive' || repo?.type === 'sharepoint';

/** O que a exclusão alcança: caminho da pasta ou tipo, locatário e contas/sites do repositório. */
export function deletionScope(repo) {
  if (!isCloudRepo(repo)) return `local|${repo.path}`;
  const c = repo.cloud || {};
  const list = [...(c.accounts || []), ...(c.sites || [])].map((v) => String(v).toLowerCase()).sort();
  return `${repo.type}|${String(repo.graph?.tenantId || '').toLowerCase()}|${c.scope}|${list.join(',')}`;
}

/** O que a exclusão de e-mails alcança: tipo, conta ou servidor, alcance e caixas da conexão. */
export function mailDeletionScope(source) {
  const s = source || {};
  const account = s.type === 'graph' ? s.graph?.tenantId : s.type === 'gmail' ? `${s.gmail?.clientEmail}|${s.gmail?.adminEmail}` : `${s.imap?.host}:${s.imap?.port}`;
  const boxes = (s.mailboxes || []).map((m) => `${m.address}${m.login ? `>${m.login}` : ''}`.toLowerCase()).sort();
  return `${s.type}|${String(account || '').toLowerCase()}|${s.scope}|${boxes.join(',')}`;
}

/** Indica se `target` está dentro da pasta `root` (sem sair dela com "..", nem em outra unidade). */
export function isInside(root, target) {
  if (!root || !target) return false;
  const rel = path.relative(root, target);
  return Boolean(rel) && rel !== '..' && !rel.startsWith(`..${path.sep}`) && !path.isAbsolute(rel);
}

/** A própria pasta `root` ou algo dentro dela. */
export function isWithin(root, target) {
  return Boolean(root && target) && (path.relative(root, target) === '' || isInside(root, target));
}

/**
 * Pastas que nunca têm arquivos excluídos: os dados e a instalação do CLEAN (menos a pasta de
 * demonstração). Cada item: { path, except?, error }.
 */
export function cleanPaths({ dataDir, appDir }) {
  return [
    { path: dataDir, data: true, error: 'Arquivo de dados do CLEAN: não é excluído.' },
    { path: appDir, except: [path.join(appDir, 'demo')], error: 'Arquivo da instalação do CLEAN: não é excluído.' },
  ].filter((p) => p.path);
}

/** Repositórios cadastrados dentro de `repo` que não permitem exclusão (o mais específico vale). */
export function keptPaths(repo, repositories) {
  return repositories
    .filter((other) => other.id !== repo.id && !other.allowDelete && isWithin(repo.path, other.path))
    .map((other) => ({ path: other.path, error: `O arquivo está dentro do repositório "${other.name}", que não permite exclusão.` }));
}

function guardFor(guards, target) {
  return guards.find((g) => isWithin(g.path, target) && !(g.except || []).some((e) => isWithin(e, target))) || null;
}

async function realOrSame(p) {
  try {
    return await fs.realpath(p);
  } catch {
    return p;
  }
}

/** Confere pasta por pasta (servidores que não informam o caminho final) que nenhuma é um atalho. */
async function checkFolders(root, dir) {
  let current = root;
  const rel = path.relative(root, dir);
  for (const part of rel ? rel.split(path.sep) : []) {
    current = path.join(current, part);
    let st;
    try {
      st = await fs.lstat(current);
    } catch (err) {
      if (err.code === 'ENOENT' || err.code === 'ENOTDIR') return { status: 'missing', error: MISSING };
      return { status: 'failed', error: friendlyError(err) };
    }
    if (st.isSymbolicLink() || !st.isDirectory()) return { status: 'failed', error: OUTSIDE };
  }
  return null;
}

/**
 * Confere pelo caminho real (resolvendo links, junções e unidades mapeadas) que a pasta do arquivo
 * continua dentro do repositório e fora das pastas protegidas: uma pasta trocada depois da análise
 * por um atalho levaria a exclusão para outro lugar.
 */
async function checkLocation(root, filePath, guards) {
  const dir = path.dirname(filePath);
  let realRoot;
  let realDir;
  try {
    [realRoot, realDir] = await Promise.all([fs.realpath(root), fs.realpath(dir)]);
  } catch (err) {
    if (err.code === 'ENOENT' || err.code === 'ENOTDIR') return { status: 'missing', error: MISSING };
    return checkFolders(root, dir);
  }
  if (!isWithin(realRoot, realDir)) return { status: 'failed', error: OUTSIDE };
  const realFile = path.join(realDir, path.basename(filePath));
  for (const g of guards) {
    const real = { ...g, path: await realOrSame(g.path), except: await Promise.all((g.except || []).map(realOrSame)) };
    if (guardFor([real], realFile)) return { status: 'failed', error: g.error };
  }
  return null;
}

/**
 * Exclui um arquivo. Retorna { status: 'deleted' | 'missing' | 'changed' | 'failed', error? }.
 * root: pasta do repositório (o arquivo precisa estar dentro dela, também pelo caminho real).
 * expected: { size, modified } registrados na análise; se o arquivo mudou depois disso, a exclusão
 * só acontece com force = true (status 'changed').
 * protect: pastas cujos arquivos nunca são excluídos ({ path, except?, error }).
 */
export async function deleteFile(filePath, { root, expected = null, force = false, protect = [] } = {}) {
  if (!isInside(root, filePath)) return { status: 'failed', error: 'O arquivo não está dentro da pasta do repositório.' };
  const guard = guardFor(protect, filePath);
  if (guard) return { status: 'failed', error: guard.error };
  const location = await checkLocation(root, filePath, protect);
  if (location) return location;
  let st;
  try {
    st = await fs.lstat(filePath);
  } catch (err) {
    if (err.code === 'ENOENT') return { status: 'missing', error: MISSING };
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
    if (err.code === 'ENOENT') return { status: 'missing', error: MISSING };
    return { status: 'failed', error: friendlyError(err) };
  }
  return { status: 'deleted' };
}

/**
 * Evento de exclusão gravado no registro da análise (deletions.ndjson) e no registro geral
 * (exclusoes.ndjson). `item` descreve o que foi excluído (caminho do arquivo ou caixa, pasta e
 * assunto da mensagem), para que o registro não dependa do relatório.
 */
export function deletionEvent(recordId, result, { mode, method, by, item }) {
  return {
    recordId,
    status: result.status,
    error: result.error || null,
    note: result.note || null,
    mode, // 'auto' (durante a análise) ou 'manual' (pelo relatório)
    method, // 'file', 'permanent' ou 'trash'
    by: by || null,
    item: item || null,
    at: new Date().toISOString(),
  };
}
