// Textos das políticas de retenção usados nas telas (os mesmos do servidor: src/retention/policy.js).

export const FILE_CRITERIA = {
  used: {
    label: 'Sem uso (nenhuma data recente)',
    short: 'sem uso',
    date: 'Data mais recente (uso)',
    period: 'com o uso mais recente (modificação, acesso ou criação) entre',
    hint: 'O mais seguro: o arquivo só expira se a modificação, o último acesso e a criação forem todos mais antigos que o limite.',
  },
  'opened-or-modified': {
    label: 'Sem abertura nem modificação',
    short: 'sem abertura nem modificação',
    date: 'Última abertura ou modificação',
    period: 'com a última abertura ou modificação entre',
    hint: 'A data mais recente entre a última abertura (último acesso) e a última modificação; a criação não conta. O último acesso só é confiável se o Windows o registrar no servidor de arquivos. Não disponível no OneDrive e no SharePoint.',
  },
  modified: {
    label: 'Última modificação',
    short: 'sem modificação',
    date: 'Última modificação',
    period: 'modificados pela última vez entre',
    hint: 'Data em que o conteúdo foi alterado pela última vez (confiável em qualquer servidor).',
  },
  accessed: {
    label: 'Último acesso (abertura)',
    short: 'sem acesso',
    date: 'Último acesso',
    period: 'abertos (acessados) pela última vez entre',
    hint: 'Só é confiável se o Windows registrar o último acesso no servidor de arquivos (muitas vezes desligado). Não disponível no OneDrive e no SharePoint.',
  },
  created: { label: 'Criação', short: 'criados', date: 'Criação', period: 'criados entre', hint: 'Data em que o arquivo foi criado ou copiado para o repositório.' },
};

export const MAIL_CRITERIA = {
  received: { label: 'Data de recebimento', short: 'recebidas', date: 'Recebida em', hint: 'Data em que a mensagem chegou à caixa (nos Itens Enviados, a data do envio).' },
};

export const UNITS = {
  days: { one: 'dia', many: 'dias', max: 36500 },
  months: { one: 'mês', many: 'meses', max: 1200 },
  years: { one: 'ano', many: 'anos', max: 100 },
};

export const DEFAULT_MAX_DELETIONS = 1000;

/** Critérios que dependem do último acesso (sem esse dado no OneDrive e no SharePoint). */
export const ACCESS_CRITERIA = new Set(['accessed', 'opened-or-modified']);

/** "AAAA-MM-DD" → "DD/MM/AAAA" (sem passar por Date: sem desvio de fuso). */
export function dateBr(text) {
  const [y, m, d] = String(text).split('-');
  return `${d}/${m}/${y}`;
}

/** "5 anos", "1 mês", "30 dias" (na data fixa: "desde DD/MM/AAAA"; no período: "de DD/MM/AAAA a DD/MM/AAAA"). */
export function amountText(r) {
  if (r.from && r.to) return `de ${dateBr(r.from)} a ${dateBr(r.to)}`;
  if (r.since) return `desde ${dateBr(r.since)}`;
  const u = UNITS[r.unit] || UNITS.years;
  return `${r.amount} ${Number(r.amount) === 1 ? u.one : u.many}`;
}

/** Descrição da regra, ex.: "Arquivos sem uso há mais de 5 anos (somente *.tmp)". */
export function describeRetention(r, kind) {
  if (!r) return '';
  if (kind === 'mail') {
    const where =
      r.includeTrash && r.includeJunk
        ? 'inclusive a Lixeira e o Lixo Eletrônico'
        : r.includeTrash
          ? 'inclusive a Lixeira; sem o Lixo Eletrônico'
          : r.includeJunk
            ? 'inclusive o Lixo Eletrônico; sem a Lixeira'
            : 'sem a Lixeira e o Lixo Eletrônico';
    const when = r.since ? `antes de ${dateBr(r.since)}` : `há mais de ${amountText(r)}`;
    return `Mensagens recebidas ${when} (${where})`;
  }
  const c = FILE_CRITERIA[r.criterion] || FILE_CRITERIA.used;
  const created = r.criterion === 'created';
  const rule =
    r.from && r.to
      ? `${c.period} ${dateBr(r.from)} e ${dateBr(r.to)}`
      : r.since
        ? `${created ? 'criados antes de' : `${c.short} desde`} ${dateBr(r.since)}`
        : `${created ? 'criados há mais de' : `${c.short} há mais de`} ${amountText(r)}`;
  const patterns = r.patterns || [];
  const only = patterns.length ? ` (somente ${patterns.slice(0, 5).join(', ')}${patterns.length > 5 ? '…' : ''})` : '';
  return `Arquivos ${rule}${only}${foldersText(r.folders)}`;
}

/** " — somente nas pastas A, B e mais 2" (vazio sem pastas). */
export function foldersText(folders) {
  if (!folders?.length) return '';
  const names = folders.slice(0, 3).map((f) => f.path);
  const more = folders.length > 3 ? ` e mais ${folders.length - 3}` : '';
  return ` — somente ${folders.length === 1 ? 'na pasta' : 'nas pastas'} ${names.join(', ')}${more}`;
}

/**
 * Instante de uma data fixa "AAAA-MM-DD" (00:00 do dia na hora local) ou null se a data for
 * inválida, anterior a 02/01/1980 ou futura (o servidor recusa as mesmas).
 */
export function sinceInstant(text, now = new Date()) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(text ?? '').trim());
  if (!m) return null;
  const [y, mo, d] = [Number(m[1]), Number(m[2]), Number(m[3])];
  const date = new Date(y, mo - 1, d);
  if (date.getFullYear() !== y || date.getMonth() !== mo - 1 || date.getDate() !== d) return null;
  if (y < 1980 || (y === 1980 && mo === 1 && d < 2) || +date > +now) return null;
  return date;
}

/**
 * Período (prévia no formulário): { from, to } — 00:00 do dia inicial e do dia final — ou null se uma
 * das datas for inválida (como no servidor) ou a inicial for posterior à final.
 */
export function periodPreview(r, now = new Date()) {
  const from = sinceInstant(r.from, now);
  const to = sinceInstant(r.to, now);
  return from && to && +from <= +to ? { from, to } : null;
}

/**
 * Data de corte (prévia no formulário): a data fixa, o dia final do período ou agora menos a idade
 * máxima, pelo calendário. null quando a regra está incompleta ou é inválida.
 */
export function cutoffPreview(r, now = new Date()) {
  if (r.mode === 'period' || (!r.mode && (r.from || r.to))) return periodPreview(r, now)?.to || null;
  if (r.mode === 'date' || (r.mode !== 'age' && r.since)) return sinceInstant(r.since, now);
  const amount = Number(r.amount);
  if (!Number.isInteger(amount) || amount < 1 || !UNITS[r.unit] || amount > UNITS[r.unit].max) return null;
  const d = new Date(now);
  if (r.unit === 'years' || r.unit === 'months') {
    // Como no servidor: anos em meses, com o último dia do mês (29/02 menos 1 ano: 28/02).
    const months = r.unit === 'years' ? amount * 12 : amount;
    const day = d.getDate();
    d.setDate(1);
    d.setMonth(d.getMonth() - months);
    d.setDate(Math.min(day, new Date(d.getFullYear(), d.getMonth() + 1, 0).getDate()));
  } else d.setDate(d.getDate() - amount);
  return d;
}

const CLOUD_TYPES = new Set(['onedrive', 'sharepoint']);

/**
 * Forma da exclusão por extenso ("exclusão definitiva", "exclusão para a lixeira"...). Nas pastas
 * do Windows ela é sempre definitiva: "para a lixeira" vale só para o e-mail, o OneDrive e o
 * SharePoint. types: tipos dos repositórios da política (arquivos).
 */
export function deletionModeText(kind, retention, types = []) {
  if (retention?.deleteMode !== 'trash') return 'exclusão definitiva';
  if (kind === 'mail') return 'exclusão para a lixeira';
  const cloud = types.some((t) => CLOUD_TYPES.has(t));
  const local = types.some((t) => !CLOUD_TYPES.has(t));
  if (cloud && local) return 'exclusão para a lixeira no OneDrive e no SharePoint e definitiva nas pastas do Windows';
  return cloud ? 'exclusão para a lixeira' : 'exclusão definitiva (pastas do Windows)';
}

/** "1 exclusão", "1.000 exclusões". */
export const deletionsText = (n) => `${Number(n).toLocaleString('pt-BR')} ${Number(n) === 1 ? 'exclusão' : 'exclusões'}`;

/** Idade em texto: "12 anos", "8 meses", "20 dias". */
export function ageText(days) {
  const n = Math.max(0, Math.floor(Number(days) || 0));
  const years = Math.floor(n / 365);
  if (years >= 1) return `${years} ${years === 1 ? 'ano' : 'anos'}`;
  const months = Math.floor(n / 30.4);
  if (months >= 2) return `${months} meses`;
  return `${n} ${n === 1 ? 'dia' : 'dias'}`;
}
