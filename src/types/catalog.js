// Busca por tipo de arquivo: categorias (pela extensão), extensões escolhidas pela pessoa, tamanho
// mínimo, conferência do tipo real pelo conteúdo e limite de exclusões automáticas por execução.
//
// Busca (depois de validada):
//   { categories: ['video', ...], extensions: ['.xyz', ...], checkContent: true | false,
//     minSizeMB: 0 (todos) | número, maxDeletions: 1000 (0: sem limite) }

export class FileTypesError extends Error {
  constructor(message) {
    super(message);
    this.status = 400;
  }
}

/**
 * Categorias, na ordem da tela. work: tipos de arquivo de trabalho (a tela avisa antes de excluí-los
 * automaticamente).
 */
export const CATEGORIES = {
  video: {
    label: 'Vídeos',
    extensions: ['.mp4', '.m4v', '.mov', '.avi', '.mkv', '.wmv', '.flv', '.f4v', '.webm', '.mpg', '.mpeg', '.m2v', '.3gp', '.3g2', '.vob', '.m2ts', '.mts', '.ogv', '.rm', '.rmvb', '.asf', '.divx', '.xvid'],
  },
  audio: {
    label: 'Músicas e áudio',
    extensions: ['.mp3', '.wav', '.wma', '.aac', '.m4a', '.m4b', '.flac', '.ogg', '.oga', '.opus', '.aif', '.aiff', '.mid', '.midi', '.amr', '.ape', '.ra', '.mka', '.alac'],
  },
  image: {
    label: 'Imagens e fotos',
    extensions: ['.jpg', '.jpeg', '.jfif', '.png', '.gif', '.bmp', '.tif', '.tiff', '.webp', '.heic', '.heif', '.cr2', '.cr3', '.nef', '.arw', '.dng', '.orf', '.rw2', '.psd', '.svg', '.ico'],
  },
  executable: {
    label: 'Executáveis e instaladores',
    extensions: ['.exe', '.msi', '.msix', '.msixbundle', '.appx', '.appxbundle', '.bat', '.cmd', '.com', '.scr', '.pif', '.ps1', '.vbs', '.vbe', '.wsf', '.jar', '.apk', '.dmg', '.pkg', '.deb', '.rpm'],
  },
  archive: {
    label: 'Compactados',
    extensions: ['.zip', '.rar', '.7z', '.tar', '.gz', '.tgz', '.bz2', '.tbz2', '.xz', '.txz', '.cab', '.arj', '.lz', '.lzh', '.zst', '.z'],
  },
  disk: {
    label: 'Imagens de disco e máquinas virtuais',
    extensions: ['.iso', '.img', '.vhd', '.vhdx', '.vmdk', '.vdi', '.qcow2', '.ova', '.ovf', '.wim', '.esd'],
  },
  temporary: {
    label: 'Temporários e backups',
    extensions: ['.tmp', '.temp', '.bak', '.old', '.orig', '.swp', '.dmp', '.chk', '.gid', '.crdownload', '.part'],
  },
  document: {
    label: 'Documentos (Word, PDF e texto)',
    extensions: ['.doc', '.docx', '.docm', '.dot', '.dotx', '.odt', '.rtf', '.pdf', '.txt', '.wpd', '.pages'],
    work: true,
  },
  spreadsheet: {
    label: 'Planilhas',
    extensions: ['.xls', '.xlsx', '.xlsm', '.xlsb', '.xlt', '.xltx', '.ods', '.csv', '.numbers'],
    work: true,
  },
  presentation: {
    label: 'Apresentações',
    extensions: ['.ppt', '.pptx', '.pptm', '.pps', '.ppsx', '.pot', '.potx', '.odp'],
    work: true,
  },
  email: {
    label: 'E-mails e arquivos de dados do Outlook',
    extensions: ['.pst', '.ost', '.msg', '.eml', '.mbox', '.emlx'],
    work: true,
  },
  database: {
    label: 'Bancos de dados',
    extensions: ['.mdb', '.accdb', '.sqlite', '.sqlite3', '.db', '.dbf', '.mdf', '.ldf', '.ndf', '.frm', '.ibd'],
    work: true,
  },
};

// Ficam de fora das categorias extensões que também são de outros tipos comuns: .ts (TypeScript),
// .raw (dados) e .key (chaves de criptografia). Elas podem ser informadas em "Outras extensões".

/** Extensões escolhidas pela pessoa (fora das categorias). */
export const CUSTOM = { label: 'Extensão escolhida' };

export const DEFAULT_MAX_DELETIONS = 1000;
const MAX_EXTENSIONS = 200;
const EXTENSION_RE = /^\.[a-z0-9][a-z0-9._~-]{0,30}$/;

const toNumber = (value) => (typeof value === 'number' ? value : typeof value === 'string' && value.trim() !== '' ? Number(value.replace(',', '.')) : NaN);

/** ".MP3", "mp3", "*.mp3" -> ".mp3" (ou null, se não for uma extensão válida). */
export function normalizeExtension(value) {
  let e = String(value || '').trim().toLowerCase();
  if (e.startsWith('*')) e = e.slice(1);
  if (!e) return null;
  if (!e.startsWith('.')) e = `.${e}`;
  return EXTENSION_RE.test(e) && !e.endsWith('.') ? e : null;
}

/** Valida a busca por tipo recebida da interface (ou da API). */
export function sanitizeFileTypes(input) {
  const t = input && typeof input === 'object' && !Array.isArray(input) ? input : {};
  const categories = [...new Set((Array.isArray(t.categories) ? t.categories : []).filter((c) => typeof c === 'string'))];
  const unknown = categories.filter((c) => !Object.hasOwn(CATEGORIES, c));
  if (unknown.length) throw new FileTypesError(`Tipo de arquivo desconhecido: ${unknown.join(', ')}.`);
  const raw = Array.isArray(t.extensions) ? t.extensions : String(t.extensions || '').split(/[\s,;]+/);
  const items = raw.map((e) => String(e).trim()).filter(Boolean);
  const invalid = items.filter((e) => !normalizeExtension(e));
  if (invalid.length) {
    throw new FileTypesError(`Extensão inválida: ${invalid.slice(0, 5).join(', ')}. Use letras e números, como .mp3 ou .tar.gz.`);
  }
  const extensions = [...new Set(items.map(normalizeExtension))];
  if (extensions.length > MAX_EXTENSIONS) throw new FileTypesError(`Informe no máximo ${MAX_EXTENSIONS} extensões.`);
  if (!categories.length && !extensions.length) throw new FileTypesError('Escolha ao menos um tipo de arquivo ou informe uma extensão.');
  const minSizeMB = t.minSizeMB === undefined || t.minSizeMB === null || t.minSizeMB === '' ? 0 : toNumber(t.minSizeMB);
  if (!Number.isFinite(minSizeMB) || minSizeMB < 0 || minSizeMB > 1_048_576) throw new FileTypesError('Informe o tamanho mínimo em MB (0 = todos os tamanhos).');
  const maxDeletions = t.maxDeletions === undefined || t.maxDeletions === null || t.maxDeletions === '' ? DEFAULT_MAX_DELETIONS : toNumber(t.maxDeletions);
  if (!Number.isInteger(maxDeletions) || maxDeletions < 0 || maxDeletions > 10_000_000) {
    throw new FileTypesError('Informe o limite de exclusões automáticas por execução (0 = sem limite).');
  }
  return {
    // Na ordem do catálogo (o resultado não depende da ordem dos cliques).
    categories: Object.keys(CATEGORIES).filter((c) => categories.includes(c)),
    extensions: extensions.sort(),
    checkContent: t.checkContent === true,
    minSizeMB: Math.round(minSizeMB * 100) / 100,
    maxDeletions,
  };
}

/**
 * Tipo do arquivo pela extensão (a mais longa que combina: ".tar.gz" antes de ".gz"), ou null.
 * Devolve { category, extension }; category 'custom' para as extensões escolhidas.
 */
export function typeMatcher(fileTypes) {
  const byExtension = new Map();
  for (const c of fileTypes.categories || []) for (const e of CATEGORIES[c].extensions) if (!byExtension.has(e)) byExtension.set(e, c);
  for (const e of fileTypes.extensions || []) if (!byExtension.has(e)) byExtension.set(e, 'custom');
  const compound = [...byExtension.keys()].filter((e) => e.indexOf('.', 1) !== -1).sort((a, b) => b.length - a.length);
  return (name) => {
    const lower = String(name).toLowerCase();
    for (const e of compound) if (lower.endsWith(e) && lower.length > e.length) return { category: byExtension.get(e), extension: e };
    const dot = lower.lastIndexOf('.');
    if (dot <= 0) return null; // sem extensão (ou arquivo oculto como ".profile")
    const extension = lower.slice(dot);
    return byExtension.has(extension) ? { category: byExtension.get(extension), extension } : null;
  };
}

export const categoryLabel = (key) => (key === 'custom' ? CUSTOM.label : CATEGORIES[key]?.label || key);

/** Descrição em português, ex.: "Vídeos, Músicas e áudio e .xyz (acima de 100 MB)". */
export function describeFileTypes(fileTypes) {
  if (!fileTypes) return '';
  const parts = [...(fileTypes.categories || []).map((c) => CATEGORIES[c]?.label || c), ...(fileTypes.extensions || [])];
  const list = parts.length > 1 ? `${parts.slice(0, -1).join(', ')} e ${parts.at(-1)}` : parts[0] || '';
  const extras = [fileTypes.minSizeMB ? `acima de ${String(fileTypes.minSizeMB).replace('.', ',')} MB` : '', fileTypes.checkContent ? 'conferindo o tipo real' : ''].filter(Boolean);
  return `${list}${extras.length ? ` (${extras.join('; ')})` : ''}`;
}
