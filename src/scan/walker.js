// Percorre as pastas de um repositório, sem seguir atalhos/junções (evita laços e duplicidades).
import fs from 'node:fs/promises';
import path from 'node:path';

/** Padrões ignorados em todos os repositórios. */
export const DEFAULT_EXCLUDES = ['$RECYCLE.BIN', 'System Volume Information', '~$*', 'Thumbs.db', 'desktop.ini', '.DS_Store', '~*.tmp'];

function wildcardToRegExp(pattern) {
  const source = pattern.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.');
  return new RegExp(`^${source}$`, 'i');
}

/**
 * Compila padrões de exclusão. Cada padrão (com curingas * e ?) é comparado ao nome do arquivo ou
 * da pasta; padrões com barra (ex.: "Financeiro\Antigo") são comparados ao caminho relativo.
 */
export function compileExclusions(patterns = []) {
  const byName = [];
  const byPath = [];
  for (const raw of patterns) {
    const p = String(raw || '').trim().replace(/[\\/]+$/, '');
    if (!p) continue;
    if (/[\\/]/.test(p)) byPath.push(wildcardToRegExp(p.replace(/\//g, '\\').replace(/^\\+/, '')));
    else byName.push(wildcardToRegExp(p));
  }
  return (name, relativePath) => {
    if (byName.some((re) => re.test(name))) return true;
    if (byPath.length === 0) return false;
    const rel = relativePath.replace(/\//g, '\\');
    return byPath.some((re) => re.test(rel));
  };
}

/**
 * Gera { type: 'file', path, name, relativePath } para cada arquivo e { type: 'error', path, error }
 * para pastas inacessíveis. `shouldStop` interrompe a varredura (cancelamento); `skipDir(caminho)`
 * pula uma pasta inteira.
 */
export async function* walk(root, { isExcluded = () => false, skipDir = () => false, shouldStop = () => false } = {}) {
  const stack = [''];
  while (stack.length > 0) {
    if (shouldStop()) return;
    const rel = stack.pop();
    const dir = rel ? path.join(root, rel) : root;
    let handle;
    try {
      handle = await fs.opendir(dir, { bufferSize: 128 });
    } catch (error) {
      yield { type: 'error', path: dir, error };
      continue;
    }
    const subdirs = [];
    try {
      for await (const entry of handle) {
        const childRel = rel ? path.join(rel, entry.name) : entry.name;
        if (isExcluded(entry.name, childRel)) continue;
        let isFile = entry.isFile();
        if (entry.isSymbolicLink()) {
          // Pontos de nova análise (reparse points): arquivos deduplicados ou em camadas de
          // armazenamento continuam sendo arquivos; pastas vinculadas (junções) são ignoradas.
          const target = await fs.stat(path.join(root, childRel)).catch(() => null);
          isFile = Boolean(target?.isFile());
        } else if (entry.isDirectory()) {
          if (!skipDir(path.join(root, childRel))) subdirs.push(childRel);
        }
        if (isFile) yield { type: 'file', path: path.join(root, childRel), name: entry.name, relativePath: childRel };
        if (shouldStop()) return;
      }
    } catch (error) {
      yield { type: 'error', path: dir, error };
    }
    yield { type: 'dir', path: dir };
    // Ordem alfabética na saída: empilha em ordem inversa.
    subdirs.sort((a, b) => b.localeCompare(a));
    stack.push(...subdirs);
  }
}
