// Políticas de retenção: arquivos e mensagens mais antigos que a idade máxima, pelo critério de data
// escolhido, são listados no relatório e — se a política excluir — eliminados.
//
// Retenção (depois de validada):
//   { criterion, amount, unit: 'days' | 'months' | 'years',
//     patterns: ['*.tmp', ...]   // arquivos: só estes nomes (vazio: todos)
//     includeTrash, includeJunk  // e-mail: inclui a Lixeira e o Lixo Eletrônico
//     maxDeletions,              // limite de exclusões por execução (0: sem limite)
//     deleteMode: 'permanent' | 'trash' }
// A data de corte ("cutoff") é calculada quando a análise é criada: itens com a data do critério
// anterior a ela estão expirados.

export class RetentionError extends Error {
  constructor(message) {
    super(message);
    this.status = 400;
  }
}

/** Critérios de data dos arquivos (a ordem é a da tela). */
export const FILE_CRITERIA = {
  used: {
    label: 'Sem uso (nenhuma data recente)',
    phrase: 'sem uso',
    hint: 'O mais seguro: o arquivo só expira se a modificação, o último acesso e a criação forem todos mais antigos que o limite.',
  },
  modified: { label: 'Última modificação', phrase: 'sem modificação', hint: 'Data em que o conteúdo foi alterado pela última vez (confiável em qualquer servidor).' },
  accessed: {
    label: 'Último acesso (abertura)',
    phrase: 'sem acesso',
    hint: 'Só é confiável se o Windows registrar o último acesso no servidor de arquivos (muitas vezes desligado). Não disponível no OneDrive e no SharePoint.',
  },
  created: { label: 'Criação', phrase: 'criados', hint: 'Data em que o arquivo foi criado ou copiado para o repositório.' },
};

export const MAIL_CRITERIA = {
  received: { label: 'Data de recebimento', phrase: 'recebidas', hint: 'Data em que a mensagem chegou à caixa (nos Itens Enviados, a data do envio).' },
};

export const UNITS = {
  days: { one: 'dia', many: 'dias', max: 36500 },
  months: { one: 'mês', many: 'meses', max: 1200 },
  years: { one: 'ano', many: 'anos', max: 100 },
};

const MAX_PATTERNS = 100;
export const DEFAULT_MAX_DELETIONS = 1000;
const DAY = 86400000;

const toInt = (value) => (typeof value === 'number' ? value : typeof value === 'string' && /^\s*\d+\s*$/.test(value) ? Number(value) : NaN);

/**
 * Valida a retenção de uma política. kind: 'files' | 'mail'; cloud: algum repositório do OneDrive
 * ou do SharePoint (sem data de último acesso).
 */
export function sanitizeRetention(input, kind, { cloud = false } = {}) {
  const r = input && typeof input === 'object' && !Array.isArray(input) ? input : {};
  const mail = kind === 'mail';
  const criteria = mail ? MAIL_CRITERIA : FILE_CRITERIA;
  const criterion = r.criterion === undefined || r.criterion === null || r.criterion === '' ? (mail ? 'received' : 'used') : r.criterion;
  if (!Object.hasOwn(criteria, criterion)) throw new RetentionError('Escolha o critério de data da política.');
  if (criterion === 'accessed' && cloud) {
    throw new RetentionError('O OneDrive e o SharePoint não informam o último acesso de cada arquivo: use "Sem uso", "Última modificação" ou "Criação".');
  }
  const unit = r.unit === undefined || r.unit === null || r.unit === '' ? 'years' : r.unit;
  if (!Object.hasOwn(UNITS, unit)) throw new RetentionError('Escolha a unidade da idade máxima: dias, meses ou anos.');
  const amount = toInt(r.amount);
  if (!Number.isInteger(amount) || amount < 1 || amount > UNITS[unit].max) {
    throw new RetentionError(`Informe a idade máxima: de 1 a ${UNITS[unit].max} ${UNITS[unit].many}.`);
  }
  const maxDeletions = r.maxDeletions === undefined || r.maxDeletions === null || r.maxDeletions === '' ? DEFAULT_MAX_DELETIONS : toInt(r.maxDeletions);
  if (!Number.isInteger(maxDeletions) || maxDeletions < 0 || maxDeletions > 10_000_000) {
    throw new RetentionError('Informe o limite de exclusões por execução (0 = sem limite).');
  }
  const out = { criterion, amount, unit, maxDeletions, deleteMode: r.deleteMode === 'trash' ? 'trash' : 'permanent' };
  if (mail) {
    out.includeTrash = r.includeTrash !== false;
    out.includeJunk = r.includeJunk !== false;
  } else {
    const items = Array.isArray(r.patterns) ? r.patterns : String(r.patterns || '').split(/[\r\n]+/);
    const patterns = [...new Set(items.map((p) => String(p).trim()).filter(Boolean))];
    if (patterns.length > MAX_PATTERNS) throw new RetentionError(`Informe no máximo ${MAX_PATTERNS} padrões de nomes.`);
    if (patterns.some((p) => p.length > 200 || /[\\/]/.test(p))) throw new RetentionError('Os padrões de nomes valem para o nome do arquivo (sem pastas), com até 200 caracteres. Ex.: *.tmp');
    out.patterns = patterns;
  }
  return out;
}

/** Data de corte: agora menos a idade máxima (meses e anos pelo calendário, na hora local). */
export function cutoffDate(retention, now = new Date()) {
  const d = new Date(now);
  if (retention.unit === 'years') d.setFullYear(d.getFullYear() - retention.amount);
  else if (retention.unit === 'months') {
    const day = d.getDate();
    d.setDate(1);
    d.setMonth(d.getMonth() - retention.amount);
    // 31/03 menos 1 mês: 28 ou 29/02 (último dia), e não 03/03.
    d.setDate(Math.min(day, new Date(d.getFullYear(), d.getMonth() + 1, 0).getDate()));
  } else d.setDate(d.getDate() - retention.amount);
  return d;
}

const valid = (ms) => (Number.isFinite(ms) && ms > 0 ? ms : null);

/** Data do critério de um arquivo (fs.Stats), em ms; null quando o sistema não informa. */
export function fileDate(st, criterion) {
  const modified = valid(st.mtimeMs);
  const accessed = valid(st.atimeMs);
  const created = valid(st.birthtimeMs);
  if (criterion === 'modified') return modified;
  if (criterion === 'accessed') return accessed;
  if (criterion === 'created') return created;
  const known = [modified, accessed, created].filter((v) => v !== null);
  return known.length ? Math.max(...known) : null;
}

/** Data do critério de um arquivo do OneDrive/SharePoint (driveItem); null quando não há. */
export function cloudDate(item, criterion) {
  const modified = valid(Date.parse(item.lastModifiedDateTime));
  const created = valid(Date.parse(item.createdDateTime));
  if (criterion === 'modified') return modified;
  if (criterion === 'created') return created;
  if (criterion === 'accessed') return null;
  const known = [modified, created].filter((v) => v !== null);
  return known.length ? Math.max(...known) : null;
}

export const ageDays = (ms, now = Date.now()) => Math.max(0, Math.floor((now - ms) / DAY));

/** Faixas de idade do relatório. */
export const AGE_BUCKETS = [
  { key: 'ate-1-ano', label: 'Até 1 ano', max: 365 },
  { key: '1-2-anos', label: '1 a 2 anos', max: 730 },
  { key: '2-5-anos', label: '2 a 5 anos', max: 1826 },
  { key: '5-10-anos', label: '5 a 10 anos', max: 3652 },
  { key: 'mais-de-10-anos', label: 'Mais de 10 anos', max: Infinity },
];

export function ageBucket(days) {
  return AGE_BUCKETS.find((b) => days < b.max) || AGE_BUCKETS.at(-1);
}

/** "5 anos", "1 mês", "30 dias". */
export function amountText(retention) {
  const u = UNITS[retention.unit] || UNITS.years;
  return `${retention.amount} ${retention.amount === 1 ? u.one : u.many}`;
}

/** Descrição em português, ex.: "Arquivos sem uso há mais de 5 anos (somente *.tmp, *.bak)". */
export function describeRetention(retention, kind) {
  if (!retention) return '';
  if (kind === 'mail') {
    const { includeTrash: trash, includeJunk: junk } = retention;
    const where =
      trash && junk
        ? 'inclusive a Lixeira e o Lixo Eletrônico'
        : trash
          ? 'inclusive a Lixeira; sem o Lixo Eletrônico'
          : junk
            ? 'inclusive o Lixo Eletrônico; sem a Lixeira'
            : 'sem a Lixeira e o Lixo Eletrônico';
    return `Mensagens recebidas há mais de ${amountText(retention)} (${where})`;
  }
  const c = FILE_CRITERIA[retention.criterion] || FILE_CRITERIA.used;
  const verb = retention.criterion === 'created' ? 'criados há mais de' : `${c.phrase} há mais de`;
  const only = retention.patterns?.length ? ` (somente ${retention.patterns.slice(0, 5).join(', ')}${retention.patterns.length > 5 ? '…' : ''})` : '';
  return `Arquivos ${verb} ${amountText(retention)}${only}`;
}

/** Nomes de arquivo aceitos pela política (curingas * e ?; sem padrões: todos). */
export function patternMatcher(patterns = []) {
  if (!patterns.length) return () => true;
  const regexes = patterns.map((p) => new RegExp(`^${p.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.')}$`, 'i'));
  return (name) => regexes.some((re) => re.test(name));
}
