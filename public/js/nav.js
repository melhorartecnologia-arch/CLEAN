// Parâmetros da rota atual (fragmento da URL: #/caminho?chave=valor).

export function currentPath() {
  return (location.hash.replace(/^#/, '') || '/').split('?')[0];
}

/** Atualiza os parâmetros da URL sem recarregar a tela (ex.: filtros do relatório). */
export function replaceQuery(params) {
  const entries = Object.entries(params).filter(([, v]) => v !== '' && v !== null && v !== undefined);
  const query = new URLSearchParams(entries).toString();
  history.replaceState(null, '', `#${currentPath()}${query ? `?${query}` : ''}`);
}

export function go(path) {
  location.hash = `#${path}`;
}
