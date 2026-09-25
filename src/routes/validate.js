// Validação dos dados recebidos pela API.
import path from 'node:path';
import crypto from 'node:crypto';
import { validateTerm, foldText } from '../scan/matcher.js';
import { VALIDATORS } from '../scan/presets.js';

export class HttpError extends Error {
  /** code: identificador opcional devolvido à interface junto com a mensagem (ex.: 'method-changed'). */
  constructor(status, message, code = '') {
    super(message);
    this.status = status;
    this.code = code;
  }
}

export const bad = (message) => new HttpError(400, message);

export function text(value, field, { required = false, max = 500 } = {}) {
  const v = typeof value === 'string' ? value.trim() : value === undefined || value === null ? '' : String(value).trim();
  if (required && !v) throw bad(`Informe ${field}.`);
  if (v.length > max) throw bad(`${field} deve ter no máximo ${max} caracteres.`);
  return v;
}

export function lines(value, max = 500) {
  const items = Array.isArray(value) ? value : String(value || '').split(/\r?\n/);
  return [...new Set(items.map((v) => String(v).trim()).filter(Boolean))].slice(0, max);
}

/** Caminho absoluto de pasta: C:\..., \\servidor\compartilhamento\... (ou /... fora do Windows). */
export function normalizeRepoPath(value) {
  let p = text(value, 'o caminho da pasta', { required: true, max: 1000 });
  const isWindowsStyle = /^[a-z]:[\\/]/i.test(p) || /^\\\\[^\\]+\\[^\\]+/.test(p);
  if (process.platform === 'win32' || isWindowsStyle) {
    if (!isWindowsStyle) throw bad('Informe um caminho completo, como D:\\Dados ou \\\\servidor\\compartilhamento.');
    p = p.replace(/\//g, '\\');
    if (!/^[a-z]:\\$/i.test(p)) p = p.replace(/\\+$/, '');
    return p;
  }
  if (!path.isAbsolute(p)) throw bad('Informe um caminho absoluto para a pasta.');
  return p.length > 1 ? p.replace(/\/+$/, '') : p;
}

export function parseRepository(body = {}) {
  const audit = body.audit || {};
  const computer = text(audit.computer, 'o computador da auditoria', { max: 255 });
  if (computer && !/^[A-Za-z0-9._-]+$/.test(computer)) throw bad('Nome de computador inválido para a auditoria.');
  const localPath = text(audit.localPath, 'o caminho local da auditoria', { max: 1000 });
  if (localPath && !/^[a-z]:\\/i.test(localPath.replace(/\//g, '\\'))) throw bad('O caminho local no servidor deve começar com a letra da unidade (ex.: E:\\Compartilhamentos).');
  const days = Number(audit.days) || 30;
  const maxEvents = Number(audit.maxEvents) || 200000;
  return {
    name: text(body.name, 'o nome do repositório', { required: true, max: 200 }),
    path: normalizeRepoPath(body.path),
    description: text(body.description, 'a descrição', { max: 1000 }),
    exclude: lines(body.exclude),
    // Permite excluir os arquivos encontrados (automaticamente na análise ou pelo relatório).
    allowDelete: body.allowDelete === true,
    audit: {
      enabled: Boolean(audit.enabled),
      computer,
      localPath: localPath.replace(/\//g, '\\'),
      days: Math.min(Math.max(Math.round(days), 1), 365),
      maxEvents: Math.min(Math.max(Math.round(maxEvents), 100), 5_000_000),
      ignoreUsers: lines(audit.ignoreUsers, 100).map((u) => u.replace(/;/g, '')),
    },
  };
}

const MAX_TERMS = 50000;

/** Valida os termos de uma lista, preservando os identificadores existentes e removendo duplicados. */
export function parseTerms(input) {
  if (!Array.isArray(input)) throw bad('A lista de termos é inválida.');
  if (input.length > MAX_TERMS) throw bad(`Cada lista pode ter no máximo ${MAX_TERMS} termos.`);
  const seen = new Set();
  const terms = [];
  input.forEach((raw, index) => {
    const type = raw?.type === 'regex' ? 'regex' : 'text';
    const term = {
      id: typeof raw?.id === 'string' && /^[\w-]{1,64}$/.test(raw.id) ? raw.id : crypto.randomUUID(),
      type,
      value: text(raw?.value, `o termo ${index + 1}`, { max: 2000 }),
      wholeWord: type === 'text' ? Boolean(raw?.wholeWord) : false,
      validator: type === 'regex' && raw?.validator && VALIDATORS[raw.validator] ? raw.validator : null,
      label: text(raw?.label, `o rótulo do termo ${index + 1}`, { max: 200 }),
    };
    if (!term.value) return;
    const error = validateTerm(term);
    if (error) throw bad(`Termo ${index + 1} ("${term.value.slice(0, 60)}"): ${error}`);
    const key = `${type}|${type === 'text' ? foldText(term.value) : term.value}|${term.wholeWord}|${term.validator}`;
    if (seen.has(key)) return;
    seen.add(key);
    terms.push(term);
  });
  return terms;
}

export function parseList(body = {}) {
  return {
    name: text(body.name, 'o nome da lista', { required: true, max: 200 }),
    description: text(body.description, 'a descrição', { max: 1000 }),
    terms: parseTerms(body.terms || []),
  };
}
