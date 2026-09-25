// Utilitários comuns aos conectores de e-mail.
import { foldText } from '../scan/matcher.js';

/** Caixa que não pode ser analisada por um motivo esperado (ex.: usuário sem licença de e-mail). */
export class SkipMailboxError extends Error {
  constructor(message) {
    super(message);
    this.name = 'SkipMailboxError';
    this.skipMailbox = true;
  }
}

function wildcardRegExp(pattern) {
  const escaped = foldText(pattern).replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.');
  return new RegExp(`^${escaped}$`);
}

/**
 * Pastas ignoradas: padrões com curingas (* e ?), sem diferenciar maiúsculas nem acentos. Um nome
 * simples vale para a pasta em qualquer nível ("Lixo Eletrônico"); com barra, compara o caminho
 * inteiro ("Caixa de Entrada/Pessoal"). Subpastas de uma pasta ignorada também são ignoradas.
 */
export function folderMatcher(patterns = []) {
  const rules = patterns
    .map((p) => String(p || '').trim().replace(/\\/g, '/').replace(/^\/+|\/+$/g, ''))
    .filter(Boolean)
    .map((p) => ({ path: p.includes('/'), re: wildcardRegExp(p) }));
  if (rules.length === 0) return () => false;
  return (folderPath) => {
    const segments = String(folderPath || '').split('/').filter(Boolean);
    for (let i = 1; i <= segments.length; i++) {
      const prefix = foldText(segments.slice(0, i).join('/'));
      const name = foldText(segments[i - 1]);
      if (rules.some((r) => r.re.test(r.path ? prefix : name))) return true;
    }
    return false;
  };
}

/** Endereços ignorados (curingas permitidos, ex.: "noreply@*"). */
export function addressMatcher(patterns = []) {
  const rules = patterns.map((p) => String(p || '').trim()).filter(Boolean).map(wildcardRegExp);
  return (address) => rules.some((re) => re.test(foldText(address)));
}

export function normalizeAddress(value) {
  return String(value || '').trim().toLowerCase();
}

/** Data ISO válida ou null. */
export function validDate(value) {
  if (!value) return null;
  const d = new Date(value);
  return Number.isNaN(d.getTime()) ? null : d;
}
