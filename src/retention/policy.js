import path from 'node:path';

// Políticas de retenção: arquivos e mensagens mais antigos que a idade máxima, pelo critério de data
// escolhido, são listados no relatório e — se a política excluir — eliminados.
//
// Retenção (depois de validada):
//   { criterion,
//     amount, unit: 'days' | 'months' | 'years'   // idade máxima (o corte anda a cada execução), ou
//     since: 'AAAA-MM-DD',                        // data fixa: sem uso desde esse dia (00:00, hora do servidor), ou
//     from, to: 'AAAA-MM-DD',                     // arquivos: período — a data do critério a partir de "from" e/ou
//                                                 // até "to" (dias inteiros; ao menos um dos dois)
//     maxDeletions,              // limite de exclusões por execução (0: sem limite)
//     deleteMode: 'permanent' | 'trash',
//     patterns: ['*.tmp', ...]   // arquivos: só estes nomes (vazio: todos)
//     folders: [{ repositoryId, path }]  // arquivos: só estas pastas (relativas à raiz; ausente: tudo)
//     includeTrash, includeJunk  // e-mail: inclui a Lixeira e o Lixo Eletrônico }
// As chaves novas (since, from/to, folders) só aparecem quando usadas: uma política salva antes delas continua
// com o mesmo JSON (o agendador compara a política da execução com a salva, scheduler.js).
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
    period: 'com o uso mais recente (modificação, acesso ou criação)',
    hint: 'O mais seguro: o arquivo só expira se a modificação, o último acesso e a criação forem todos mais antigos que o limite.',
  },
  'opened-or-modified': {
    label: 'Sem abertura nem modificação',
    phrase: 'sem abertura nem modificação',
    period: 'com a última abertura ou modificação',
    hint: 'A data mais recente entre a última abertura (último acesso) e a última modificação; a criação não conta. O último acesso só é confiável se o Windows o registrar no servidor de arquivos. Não disponível no OneDrive e no SharePoint.',
  },
  modified: { label: 'Última modificação', phrase: 'sem modificação', period: 'modificados pela última vez', hint: 'Data em que o conteúdo foi alterado pela última vez (confiável em qualquer servidor).' },
  accessed: {
    label: 'Último acesso (abertura)',
    phrase: 'sem acesso',
    period: 'abertos (acessados) pela última vez',
    hint: 'Só é confiável se o Windows registrar o último acesso no servidor de arquivos (muitas vezes desligado). Não disponível no OneDrive e no SharePoint.',
  },
  created: { label: 'Criação', phrase: 'criados', period: 'criados', hint: 'Data em que o arquivo foi criado ou copiado para o repositório.' },
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
const MAX_FOLDERS = 200;
export const DEFAULT_MAX_DELETIONS = 1000;
const DAY = 86400000;
// Critérios que dependem do último acesso (abertura), que o OneDrive e o SharePoint não informam.
const ACCESS_CRITERIA = new Set(['accessed', 'opened-or-modified']);

const toInt = (value) => (typeof value === 'number' ? value : typeof value === 'string' && /^\s*\d+\s*$/.test(value) ? Number(value) : NaN);

/**
 * Valida a retenção de uma política. kind: 'files' | 'mail'; cloud: algum repositório do OneDrive
 * ou do SharePoint (sem data de último acesso); targets: os repositórios da política (para validar as
 * pastas); now: o "agora" da data fixa (que não pode estar no futuro).
 */
export function sanitizeRetention(input, kind, { cloud = false, targets = null, now = new Date() } = {}) {
  const r = input && typeof input === 'object' && !Array.isArray(input) ? input : {};
  const mail = kind === 'mail';
  const criteria = mail ? MAIL_CRITERIA : FILE_CRITERIA;
  const criterion = r.criterion === undefined || r.criterion === null || r.criterion === '' ? (mail ? 'received' : 'used') : r.criterion;
  if (!Object.hasOwn(criteria, criterion)) throw new RetentionError('Escolha o critério de data da política.');
  if (ACCESS_CRITERIA.has(criterion) && cloud) {
    throw new RetentionError(
      criterion === 'accessed'
        ? 'O OneDrive e o SharePoint não informam o último acesso de cada arquivo: use "Sem uso", "Última modificação" ou "Criação".'
        : 'O OneDrive e o SharePoint não informam a última abertura (último acesso) de cada arquivo: use "Sem uso", "Última modificação" ou "Criação".',
    );
  }
  // Corte: idade máxima (relativa a cada execução), data fixa ("sem uso desde") ou período (a data do
  // critério entre dois dias). Sem "mode", vale a forma salva: com "from"/"to", período; com "since",
  // data fixa.
  const filled = (v) => typeof v === 'string' && v.trim() !== '';
  const mode = ['age', 'date', 'period'].includes(r.mode) ? r.mode : filled(r.from) || filled(r.to) ? 'period' : filled(r.since) ? 'date' : 'age';
  let cut;
  if (mode === 'period') {
    if (mail) throw new RetentionError('O período (a partir de / até) vale para arquivos; nas mensagens, use a idade máxima ou a data fixa.');
    // Uma das datas pode ficar em branco: só "a partir de" (até o início de cada execução) ou só "até".
    if (!filled(r.from) && !filled(r.to)) throw new RetentionError('Informe o período: a data "a partir de", a data "até" ou as duas.');
    const from = filled(r.from) ? dayField(r.from, now, 'inicial do período') : null;
    const to = filled(r.to) ? dayField(r.to, now, 'final do período') : null;
    if (from && to && from > to) throw new RetentionError(`A data inicial do período (${dateBr(from)}) é posterior à data final (${dateBr(to)}).`);
    cut = { ...(from ? { from } : {}), ...(to ? { to } : {}) };
  } else if (mode === 'date') {
    cut = { since: sinceField(r.since, now, mail) };
  } else {
    const unit = r.unit === undefined || r.unit === null || r.unit === '' ? 'years' : r.unit;
    if (!Object.hasOwn(UNITS, unit)) throw new RetentionError('Escolha a unidade da idade máxima: dias, meses ou anos.');
    const amount = toInt(r.amount);
    if (!Number.isInteger(amount) || amount < 1 || amount > UNITS[unit].max) {
      throw new RetentionError(`Informe a idade máxima: de 1 a ${UNITS[unit].max} ${UNITS[unit].many}.`);
    }
    cut = { amount, unit };
  }
  const maxDeletions = r.maxDeletions === undefined || r.maxDeletions === null || r.maxDeletions === '' ? DEFAULT_MAX_DELETIONS : toInt(r.maxDeletions);
  if (!Number.isInteger(maxDeletions) || maxDeletions < 0 || maxDeletions > 10_000_000) {
    throw new RetentionError('Informe o limite de exclusões por execução (0 = sem limite).');
  }
  const out = { criterion, ...cut, maxDeletions, deleteMode: r.deleteMode === 'trash' ? 'trash' : 'permanent' };
  // Pastas num formato inesperado (um objeto, um texto) são recusadas: ignorá-las faria a política
  // valer para o repositório inteiro.
  if (r.folders !== undefined && r.folders !== null && !Array.isArray(r.folders)) {
    throw new RetentionError('Informe as pastas da política como uma lista de { repositoryId, path }.');
  }
  if (mail) {
    if (Array.isArray(r.folders) && r.folders.length) throw new RetentionError('A escolha de pastas vale para arquivos (repositórios), e não para e-mails.');
    out.includeTrash = r.includeTrash !== false;
    out.includeJunk = r.includeJunk !== false;
  } else {
    const items = Array.isArray(r.patterns) ? r.patterns : String(r.patterns || '').split(/[\r\n]+/);
    const patterns = [...new Set(items.map((p) => String(p).trim()).filter(Boolean))];
    if (patterns.length > MAX_PATTERNS) throw new RetentionError(`Informe no máximo ${MAX_PATTERNS} padrões de nomes.`);
    if (patterns.some((p) => p.length > 200 || /[\\/]/.test(p))) throw new RetentionError('Os padrões de nomes valem para o nome do arquivo (sem pastas), com até 200 caracteres. Ex.: *.tmp');
    out.patterns = patterns;
    const folders = sanitizeFolders(r.folders, { targets });
    if (folders.length) out.folders = folders;
  }
  return out;
}

/**
 * Dia do período: "AAAA-MM-DD" válido (calendário real), a partir de 02/01/1980 e não no futuro (hora
 * do servidor). which: "inicial do período" ou "final do período".
 */
function dayField(value, now, which) {
  const text = String(value ?? '').trim();
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(text);
  if (!m) throw new RetentionError(`Informe a data ${which} (dia, mês e ano).`);
  const [y, mo, d] = [Number(m[1]), Number(m[2]), Number(m[3])];
  const date = new Date(y, mo - 1, d);
  if (date.getFullYear() !== y || date.getMonth() !== mo - 1 || date.getDate() !== d) throw new RetentionError(`Data ${which} inválida: ${text}.`);
  if (y < 1980 || (y === 1980 && mo === 1 && d < 2)) throw new RetentionError(`A data ${which} deve ser a partir de 02/01/1980 (datas anteriores são tratadas como desconhecidas).`);
  if (+date > +now) throw new RetentionError(`A data ${which} não pode estar no futuro.`);
  return text;
}

/** "AAAA-MM-DD" válido (calendário real), a partir de 02/01/1980 e não no futuro (hora do servidor). */
function sinceField(value, now, mail) {
  const text = String(value ?? '').trim();
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(text);
  if (!m) throw new RetentionError('Informe a data de corte (dia, mês e ano).');
  const [y, mo, d] = [Number(m[1]), Number(m[2]), Number(m[3])];
  const date = new Date(y, mo - 1, d);
  if (date.getFullYear() !== y || date.getMonth() !== mo - 1 || date.getDate() !== d) throw new RetentionError(`Data de corte inválida: ${text}.`);
  if (y < 1980 || (y === 1980 && mo === 1 && d < 2)) throw new RetentionError('A data de corte deve ser a partir de 02/01/1980 (datas anteriores são tratadas como desconhecidas).');
  if (+date > +now) {
    throw new RetentionError(`A data de corte não pode estar no futuro: ${mail ? 'todas as mensagens seriam consideradas antigas' : 'todos os arquivos seriam considerados sem uso'}.`);
  }
  return text;
}

/** Instante da data fixa: 00:00 do dia, na hora local do servidor (e não UTC). */
export function sinceDate(since) {
  const [y, m, d] = String(since).split('-').map(Number);
  return new Date(y, m - 1, d);
}

/** "AAAA-MM-DD" → "DD/MM/AAAA" (sem passar por Date: sem desvio de fuso). */
export function dateBr(text) {
  const [y, m, d] = String(text).split('-');
  return `${d}/${m}/${y}`;
}

// ---------- pastas da política ----------

const winStyle = (p) => /^[a-z]:[\\/]/i.test(p) || /^[\\/]{2}[^\\/]+[\\/]+[^\\/]+/.test(p);
const isCloudType = (type) => type === 'onedrive' || type === 'sharepoint';
// Nomes de pasta proibidos no Windows/SharePoint ou perigosos (curingas, fluxos alternativos ":").
const BAD_SEGMENT = /[*?"<>|:\u0000-\u001f]/;

/**
 * Caminho de pasta relativo à raiz do repositório, com "/" (vazio: a raiz, o repositório inteiro).
 * Aceita relativo (Financeiro\2019, Documentos/Projetos) ou, nas pastas do Windows, o caminho
 * completo dentro do repositório (D:\Dados\Financeiro\2019). Lança RetentionError.
 */
export function folderRelative(repo, value) {
  let text = String(value ?? '').trim();
  if (!text) throw new RetentionError(`Informe o caminho da pasta do repositório "${repo.name}".`);
  if (text.length > 1000) throw new RetentionError('O caminho de uma pasta deve ter no máximo 1000 caracteres.');
  if (!isCloudType(repo.type)) {
    const win = winStyle(repo.path);
    const p = win ? path.win32 : path.posix;
    // Caminho completo: C:\..., \\servidor\..., ou /... num repositório fora do Windows.
    const absolute = winStyle(text) || (!win && text.startsWith('/'));
    if (absolute) {
      const rel = p.relative(p.resolve(repo.path), p.resolve(text));
      if (rel === '') return '';
      if (rel === '..' || rel.startsWith(`..${p.sep}`) || p.isAbsolute(rel)) {
        throw new RetentionError(`A pasta "${text}" não fica dentro do repositório "${repo.name}" (${repo.path}). Informe uma pasta dentro dele.`);
      }
      text = rel;
    }
  } else if (/^https?:\/\//i.test(text)) {
    throw new RetentionError('Nas pastas do OneDrive e do SharePoint, informe o caminho dentro da biblioteca (ex.: Documentos/Projetos), e não o endereço.');
  }
  const segments = text.split(/[\\/]+/).filter(Boolean);
  if (!segments.length) return '';
  for (const [i, s] of segments.entries()) {
    if (s === '.' || s === '..') throw new RetentionError(`Caminho de pasta inválido: "${text}" (sem "." nem "..").`);
    // Espaço no início do primeiro nome também é recusado: a validação seguinte (a cada execução)
    // tiraria o espaço com o trim e a política passaria a valer para outra pasta. Nas pastas de
    // dentro ("Financeiro\ Arquivo") ele é mantido.
    if (BAD_SEGMENT.test(s) || /[. ]$/.test(s) || (i === 0 && /^\s/.test(s)) || s.length > 255) {
      throw new RetentionError(`Caminho de pasta inválido: "${text}". Informe a pasta sem curingas (* ?), sem os caracteres : " < > |, sem espaço no início e sem ponto ou espaço no fim dos nomes.`);
    }
  }
  return segments.join('/');
}

/** Chave de comparação de um caminho relativo de pasta ("/" como separador). */
export function folderKey(rel, insensitive) {
  const n = String(rel ?? '').replace(/[\\/]+/g, '/').replace(/^\/+|\/+$/g, '');
  return insensitive ? n.toLowerCase() : n;
}

/** Comparação sem diferenciar maiúsculas: pastas do Windows e da nuvem. */
export const foldersInsensitive = (repo) => isCloudType(repo?.type) || winStyle(repo?.path || '');

/**
 * Valida as pastas da política: [{ repositoryId, path }], com o caminho relativo à raiz do
 * repositório. targets: os repositórios da política (obrigatório quando há pastas: sem eles a
 * validação falha, em vez de valer para o repositório inteiro). Pastas repetidas ou dentro de outra
 * pasta escolhida saem; a raiz (o caminho do próprio repositório) vale o repositório inteiro.
 * O resultado é ordenado: validar de novo devolve o mesmo JSON.
 */
export function sanitizeFolders(input, { targets = null } = {}) {
  if (input !== undefined && input !== null && !Array.isArray(input)) throw new RetentionError('Informe as pastas da política como uma lista de { repositoryId, path }.');
  const items = input || [];
  if (!items.length) return [];
  if (!Array.isArray(targets)) throw new RetentionError('Não foi possível conferir as pastas da política: os repositórios não foram informados.');
  if (items.length > MAX_FOLDERS) throw new RetentionError(`Informe no máximo ${MAX_FOLDERS} pastas.`);
  const byId = new Map(targets.filter(Boolean).map((t) => [t.id, t]));
  const perRepo = new Map(); // repositoryId → [{ path, key }]
  const whole = new Set();
  for (const raw of items) {
    const item = raw && typeof raw === 'object' && !Array.isArray(raw) ? raw : {};
    const repo = byId.get(item.repositoryId);
    if (!repo) throw new RetentionError('Uma das pastas é de um repositório que não está marcado na política.');
    const rel = folderRelative(repo, item.path);
    if (rel === '') {
      whole.add(repo.id);
      continue;
    }
    const list = perRepo.get(repo.id) || perRepo.set(repo.id, []).get(repo.id);
    list.push({ path: rel, key: folderKey(rel, foldersInsensitive(repo)) });
  }
  const out = [];
  for (const id of [...perRepo.keys()].sort()) {
    if (whole.has(id)) continue; // a raiz escolhida: o repositório inteiro
    const sorted = perRepo.get(id).sort((a, b) => (a.key < b.key ? -1 : a.key > b.key ? 1 : 0));
    const kept = [];
    for (const f of sorted) {
      // Repetida ou dentro de outra pasta escolhida: já está coberta (não lista nem exclui duas vezes).
      if (kept.some((k) => f.key === k.key || f.key.startsWith(`${k.key}/`))) continue;
      kept.push(f);
    }
    for (const f of kept) out.push({ repositoryId: id, path: f.path });
  }
  return out;
}

/** Pastas da política num repositório (relativas, com "/"); null: o repositório inteiro. */
export function foldersOf(retention, repositoryId) {
  const list = (retention?.folders || []).filter((f) => f.repositoryId === repositoryId).map((f) => f.path);
  return list.length ? list : null;
}

/**
 * Escopo de pastas: allowDir(rel) — a pasta está dentro de uma pasta da política ou é uma das
 * pastas acima dela (é preciso passar por ela); allowFile(rel) — o arquivo está dentro de uma pasta
 * da política; folderOf(rel) — a pasta da política que contém o arquivo. rel: caminho relativo à
 * raiz (separador \ ou /). Sem pastas válidas, nada é permitido (nunca vale o repositório inteiro).
 */
export function folderScope(paths, { insensitive = false } = {}) {
  const list = (paths || []).map((p) => folderKey(p, insensitive)).filter(Boolean);
  const folderOf = (rel) => {
    const r = folderKey(rel, insensitive);
    return list.find((f) => r.startsWith(`${f}/`)) ?? null;
  };
  return {
    folders: list,
    allowDir: (rel) => {
      const r = folderKey(rel, insensitive);
      return list.some((f) => r === f || r.startsWith(`${f}/`) || f.startsWith(`${r}/`));
    },
    allowFile: (rel) => folderOf(rel) !== null,
    folderOf,
  };
}

/**
 * Pastas escolhidas numa biblioteca do OneDrive/SharePoint: cada uma vale pelo caminho dentro da
 * biblioteca ou com o nome de uma biblioteca na frente (Documentos/Projetos), que vale só nela; o
 * nome da própria biblioteca vale a biblioteca inteira (null). libraries: nomes das bibliotecas do
 * mesmo site ou conta (a pasta que começa pelo nome de outra delas é daquela biblioteca, e não um
 * caminho nesta); sem a lista, só o nome desta biblioteca é reconhecido.
 * Devolve { rels: [{ spec, rel }], scope, specOf(key) } (sem diferenciar maiúsculas).
 */
export function cloudFolderScope(specs, library, libraries = null) {
  const lib = folderKey(library, true);
  const names = new Set([lib, ...(libraries || []).map((n) => folderKey(n, true))]);
  const rels = [];
  for (const spec of specs) {
    const key = folderKey(spec, true);
    if (key === lib) return null;
    const first = key.split('/')[0];
    if (first === lib) rels.push({ spec, rel: key.slice(lib.length + 1) });
    else if (!names.has(first)) rels.push({ spec, rel: key });
    // Começa pelo nome de outra biblioteca: vale só nela.
  }
  const scope = folderScope(rels.map((r) => r.rel), { insensitive: true });
  // A pasta escolhida (como foi escrita) que corresponde à chave encontrada pelo escopo.
  const specOf = (key) => (key ? (rels.find((r) => r.rel === key)?.spec ?? key) : null);
  return { rels, scope, specOf };
}

/**
 * O arquivo (registro de um relatório de retenção) está dentro das pastas escolhidas na política?
 * Sem pastas para o repositório dele, sim. Conferência pelo caminho (a das pastas do Windows pelo
 * caminho real é feita em deleteFile, com `within`).
 */
export function recordInScope(retention, record) {
  const specs = foldersOf(retention, record.repositoryId);
  if (!specs) return true;
  if (record.cloud) {
    // A análise conhecia as bibliotecas do site e gravou a pasta escolhida que valeu para o arquivo:
    // a conferência usa só ela (sem a lista de bibliotecas, outra pasta poderia valer de outro jeito).
    const chosen = record.retention?.folder;
    const list = chosen && specs.some((s) => folderKey(s, true) === folderKey(chosen, true)) ? [chosen] : specs;
    const cloud = cloudFolderScope(list, record.cloud.library || '');
    return cloud === null || cloud.scope.allowFile(record.relativePath || '');
  }
  return folderScope(specs, { insensitive: process.platform === 'win32' }).allowFile(record.relativePath || '');
}

/**
 * Data de corte: a data fixa (00:00 do dia, hora local) ou agora menos a idade máxima (meses e anos
 * pelo calendário, na hora local).
 */
export function cutoffDate(retention, now = new Date()) {
  // Período: até o fim do dia final (expiram as datas anteriores às 00:00 do dia seguinte, montado
  // pelo calendário: num dia sem 00:00, no início do horário de verão, a hora não escorrega). Com o dia
  // final hoje, o fim é o momento em que a execução começa: um arquivo aberto ou modificado durante a
  // execução não entra no período (e a conferência na hora de excluir o mantém).
  if (retention.to) {
    const [y, m, d] = String(retention.to).split('-').map(Number);
    const end = new Date(y, m - 1, d + 1);
    return +end > +now ? new Date(now) : end;
  }
  // Só "a partir de": até o momento em que a execução começa.
  if (retention.from) return new Date(now);
  if (retention.since) return sinceDate(retention.since);
  const d = new Date(now);
  if (retention.unit === 'years' || retention.unit === 'months') {
    const months = retention.unit === 'years' ? retention.amount * 12 : retention.amount;
    const day = d.getDate();
    d.setDate(1);
    d.setMonth(d.getMonth() - months);
    // 31/03 menos 1 mês: 28 ou 29/02 (último dia), e não 03/03; 29/02 menos 1 ano: 28/02.
    d.setDate(Math.min(day, new Date(d.getFullYear(), d.getMonth() + 1, 0).getDate()));
  } else d.setDate(d.getDate() - retention.amount);
  return d;
}

/**
 * Datas anteriores a 02/01/1980 são tratadas como desconhecidas: são valores padrão de sistemas
 * antigos ou de datas perdidas (01/01/1970, 01/01/1980 do FAT e do ZIP), e não a idade real do item.
 */
export const MIN_VALID_DATE = Date.UTC(1980, 0, 2);
export const validDate = (ms) => (Number.isFinite(ms) && ms >= MIN_VALID_DATE ? ms : null);
const valid = validDate;

/** Data do critério de um arquivo (fs.Stats), em ms; null quando o sistema não informa. */
export function fileDate(st, criterion) {
  const modified = valid(st.mtimeMs);
  const accessed = valid(st.atimeMs);
  const created = valid(st.birthtimeMs);
  if (criterion === 'modified') return modified;
  if (criterion === 'accessed') return accessed;
  if (criterion === 'created') return created;
  // Abertura ou modificação: a mais recente das duas; a criação (ou a cópia) não conta.
  const known = (criterion === 'opened-or-modified' ? [modified, accessed] : [modified, accessed, created]).filter((v) => v !== null);
  return known.length ? Math.max(...known) : null;
}

/** Data do critério de um arquivo do OneDrive/SharePoint (driveItem); null quando não há. */
export function cloudDate(item, criterion) {
  const modified = valid(Date.parse(item.lastModifiedDateTime));
  const created = valid(Date.parse(item.createdDateTime));
  if (criterion === 'modified') return modified;
  if (criterion === 'created') return created;
  // Sem último acesso na nuvem: nunca expira por um critério que dependa dele.
  if (ACCESS_CRITERIA.has(criterion)) return null;
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

/**
 * Aviso do limite de exclusões da execução, pelos números finais (null: nada deixou de ser
 * excluído por causa do limite). kind: 'files' | 'mail'.
 */
export function limitWarning(
  { limit, deleted, failures, skipped },
  kind = 'files',
  { found = 'expirado', where = 'na política', owner = 'da política', later = 'Eles serão excluídos nas próximas execuções.' } = {},
) {
  if (!limit || !skipped) return null;
  const mail = kind === 'mail';
  const o = mail ? 'a' : 'o'; // "mensagens listadas", "arquivos listados"
  const items = mail ? `${skipped} mensagem(ns) ${found.replace(/o$/, 'a')}(s)` : `${skipped} arquivo(s) ${found}(s)`;
  if (failures >= limit) {
    return `${limit === 1 ? 'Uma falha' : `${limit} falhas`} de exclusão nesta execução (o limite ${owner}): as exclusões foram interrompidas e ${items} só foram listad${o}s. Confira a aba Erros (permissões, itens em uso, rótulos de retenção...).`;
  }
  if (deleted >= limit) {
    return `Limite de ${limit} ${limit === 1 ? 'exclusão' : 'exclusões'} desta execução atingido: ${items} só foram listad${o}s. Confira o relatório e, se estiver certo, aumente o limite ${where}.`;
  }
  const next = mail ? later.replace('Eles serão excluídos', 'Elas serão excluídas') : later;
  return `${items} não foram excluíd${o}s nesta execução: as vagas do limite ficaram reservadas para exclusões que falharam. ${next}`;
}

/**
 * Janela da expiração de uma análise: { cutoffMs, floorMs } — expiram as datas do critério a partir
 * de floorMs (início do período, 00:00 do dia inicial; null sem período) e anteriores a cutoffMs.
 */
export function expiryWindow(retention) {
  return { cutoffMs: Date.parse(retention.cutoff), floorMs: retention.from ? sinceDate(retention.from).getTime() : null };
}

/** A data do critério (ms) está na janela de expiração? */
export const inWindow = (when, w) => when !== null && when < w.cutoffMs && (w.floorMs === null || when >= w.floorMs);

/**
 * Período por extenso: "entre 01/01/2015 e 31/12/2018", "a partir de 01/01/2015" ou "até 31/12/2018"
 * (between: false → "de 01/01/2015 a 31/12/2018"). Vazio sem período.
 */
export function periodText(retention, { between = true } = {}) {
  const { from, to } = retention || {};
  if (from && to) return between ? `entre ${dateBr(from)} e ${dateBr(to)}` : `de ${dateBr(from)} a ${dateBr(to)}`;
  if (from) return `a partir de ${dateBr(from)}`;
  if (to) return `até ${dateBr(to)}`;
  return '';
}

/** "5 anos", "1 mês", "30 dias" (na data fixa: "desde DD/MM/AAAA"; no período: "de DD/MM/AAAA a DD/MM/AAAA"). */
export function amountText(retention) {
  if (retention.from || retention.to) return periodText(retention, { between: false });
  if (retention.since) return `desde ${dateBr(retention.since)}`;
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
    const when = retention.since ? `antes de ${dateBr(retention.since)}` : `há mais de ${amountText(retention)}`;
    return `Mensagens recebidas ${when} (${where})`;
  }
  const c = FILE_CRITERIA[retention.criterion] || FILE_CRITERIA.used;
  const created = retention.criterion === 'created';
  const rule =
    retention.from || retention.to
      ? `${c.period} ${periodText(retention)}`
      : retention.since
        ? `${created ? 'criados antes de' : `${c.phrase} desde`} ${dateBr(retention.since)}`
        : `${created ? 'criados há mais de' : `${c.phrase} há mais de`} ${amountText(retention)}`;
  const only = retention.patterns?.length ? ` (somente ${retention.patterns.slice(0, 5).join(', ')}${retention.patterns.length > 5 ? '…' : ''})` : '';
  return `Arquivos ${rule}${only}${foldersText(retention.folders)}`;
}

/** " — somente nas pastas A, B e mais 2" (vazio sem pastas). */
export function foldersText(folders) {
  if (!folders?.length) return '';
  const names = folders.slice(0, 3).map((f) => f.path);
  const more = folders.length > 3 ? ` e mais ${folders.length - 3}` : '';
  return ` — somente ${folders.length === 1 ? 'na pasta' : 'nas pastas'} ${names.join(', ')}${more}`;
}

/**
 * O nome combina com o padrão (curingas * e ?, sem diferenciar maiúsculas)? Comparação linear, sem
 * expressões regulares: um padrão com vários * não fica lento com nomes compridos.
 */
function globMatch(pattern, name) {
  let p = 0;
  let n = 0;
  let star = -1;
  let mark = 0;
  while (n < name.length) {
    // O curinga vem primeiro: um "*" no próprio nome (possível em compartilhamentos Samba) não o anula.
    if (p < pattern.length && pattern[p] === '*') {
      star = p++;
      mark = n;
    } else if (p < pattern.length && (pattern[p] === '?' || pattern[p] === name[n])) {
      p++;
      n++;
    } else if (star !== -1) {
      p = star + 1;
      n = ++mark;
    } else {
      return false;
    }
  }
  while (pattern[p] === '*') p++;
  return p === pattern.length;
}

/** Nomes de arquivo aceitos pela política (curingas * e ?; sem padrões: todos). */
export function patternMatcher(patterns = []) {
  if (!patterns.length) return () => true;
  const lower = patterns.map((p) => p.toLowerCase().replace(/\*+/g, '*'));
  return (name) => {
    const n = String(name).toLowerCase();
    return lower.some((p) => globMatch(p, n));
  };
}
