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

/** Destaca o item do menu (data-nav) correspondente à tela atual. */
export function setActiveNav(key) {
  let active = null;
  document.querySelectorAll('.nav a').forEach((a) => {
    if (key && a.dataset.nav === key) {
      a.setAttribute('aria-current', 'page');
      active = a;
    } else {
      a.removeAttribute('aria-current');
    }
  });
  // Em telas estreitas o menu fica numa faixa horizontal com rolagem: centraliza o item ativo.
  const bar = active?.closest('.sidebar');
  if (bar && bar.scrollWidth > bar.clientWidth) {
    const offset = active.getBoundingClientRect().left - bar.getBoundingClientRect().left;
    bar.scrollLeft += offset - (bar.clientWidth - active.offsetWidth) / 2;
  }
}
