// Textos das políticas de retenção usados nas telas (os mesmos do servidor: src/retention/policy.js).

export const FILE_CRITERIA = {
  used: {
    label: 'Sem uso (nenhuma data recente)',
    short: 'sem uso',
    date: 'Data mais recente (uso)',
    hint: 'O mais seguro: o arquivo só expira se a modificação, o último acesso e a criação forem todos mais antigos que o limite.',
  },
  modified: {
    label: 'Última modificação',
    short: 'sem modificação',
    date: 'Última modificação',
    hint: 'Data em que o conteúdo foi alterado pela última vez (confiável em qualquer servidor).',
  },
  accessed: {
    label: 'Último acesso (abertura)',
    short: 'sem acesso',
    date: 'Último acesso',
    hint: 'Só é confiável se o Windows registrar o último acesso no servidor de arquivos (muitas vezes desligado). Não disponível no OneDrive e no SharePoint.',
  },
  created: { label: 'Criação', short: 'criados', date: 'Criação', hint: 'Data em que o arquivo foi criado ou copiado para o repositório.' },
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

/** "5 anos", "1 mês", "30 dias". */
export function amountText(r) {
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
    return `Mensagens recebidas há mais de ${amountText(r)} (${where})`;
  }
  const c = FILE_CRITERIA[r.criterion] || FILE_CRITERIA.used;
  const verb = r.criterion === 'created' ? 'criados há mais de' : `${c.short} há mais de`;
  const patterns = r.patterns || [];
  const only = patterns.length ? ` (somente ${patterns.slice(0, 5).join(', ')}${patterns.length > 5 ? '…' : ''})` : '';
  return `Arquivos ${verb} ${amountText(r)}${only}`;
}

/** Data de corte de hoje (prévia no formulário): agora menos a idade máxima, pelo calendário. */
export function cutoffPreview(r, now = new Date()) {
  const amount = Number(r.amount);
  if (!Number.isInteger(amount) || amount < 1 || !UNITS[r.unit] || amount > UNITS[r.unit].max) return null;
  const d = new Date(now);
  if (r.unit === 'years') d.setFullYear(d.getFullYear() - amount);
  else if (r.unit === 'months') {
    const day = d.getDate();
    d.setDate(1);
    d.setMonth(d.getMonth() - amount);
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
