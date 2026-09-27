// Busca por tipo de arquivo na interface: nomes das categorias (do catálogo do servidor, lido de
// /api/info) e a descrição da busca, como no servidor.
let catalog = { categories: [], customLabel: 'Extensão escolhida', defaultMaxDeletions: 1000 };

/** Guarda o catálogo recebido do servidor (ao carregar as informações do servidor). */
export function setFileTypesCatalog(value) {
  if (value?.categories) catalog = value;
}

export const fileTypesCatalog = () => catalog;

/** Nome de uma categoria ("custom": as extensões informadas pela pessoa). */
export const categoryLabel = (key) => (key === 'custom' ? catalog.customLabel : catalog.categories.find((c) => c.key === key)?.label || key || '—');

/** "Vídeos, Músicas e áudio e .xyz (a partir de 100 MB; conferindo o tipo real)". */
export function describeFileTypes(fileTypes) {
  if (!fileTypes) return '';
  const parts = [...(fileTypes.categories || []).map(categoryLabel), ...(fileTypes.extensions || [])];
  const list = parts.length > 1 ? `${parts.slice(0, -1).join(', ')} e ${parts.at(-1)}` : parts[0] || '';
  const extras = [fileTypes.minSizeMB ? `a partir de ${String(fileTypes.minSizeMB).replace('.', ',')} MB` : '', fileTypes.checkContent ? 'conferindo o tipo real' : ''].filter(Boolean);
  return `${list}${extras.length ? ` (${extras.join('; ')})` : ''}`;
}
