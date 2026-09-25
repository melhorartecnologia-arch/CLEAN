// Painel inicial: números gerais, primeiros passos e análises recentes.
import { get } from '../api.js';
import { html, render as paint, icon, fmtNum, fmtCompact, fmtDateTime, statusBadge, plural } from '../ui.js';

export async function render(root) {
  const [repos, lists, scans] = await Promise.all([get('/api/repositories'), get('/api/lists'), get('/api/scans')]);
  const terms = lists.reduce((sum, l) => sum + l.termCount, 0);
  const lastDone = scans.find((s) => s.status === 'completed');
  const running = scans.filter((s) => s.status === 'running' || s.status === 'queued');
  const ready = repos.length > 0 && lists.length > 0 && terms > 0;

  paint(
    root,
    html`<div class="page-head">
        <div>
          <h1>Painel</h1>
          <div class="sub">Procure termos da lista de referência no nome e no conteúdo dos arquivos e descubra quem mexeu neles por último.</div>
        </div>
        <div class="actions">
          <a class="btn primary" href="#/analises/nova">${icon('play')} Nova análise</a>
        </div>
      </div>

      <section class="tiles" aria-label="Resumo">
        <div class="tile"><div class="label">Repositórios</div><div class="value">${fmtNum(repos.length)}</div><div class="detail">pastas e compartilhamentos</div></div>
        <div class="tile"><div class="label">Listas de referência</div><div class="value">${fmtNum(lists.length)}</div><div class="detail">${plural(terms, 'termo', 'termos')}</div></div>
        <div class="tile"><div class="label">Análises</div><div class="value">${fmtNum(scans.length)}</div><div class="detail">${running.length ? `${running.length} em andamento` : 'nenhuma em andamento'}</div></div>
        <div class="tile">
          <div class="label">Arquivos com ocorrências</div>
          <div class="value">${lastDone ? fmtCompact(lastDone.stats?.filesMatched) : '—'}</div>
          <div class="detail">${lastDone ? `na última análise concluída (${fmtDateTime(lastDone.finishedAt)})` : 'nenhuma análise concluída'}</div>
        </div>
      </section>

      ${ready
        ? ''
        : html`<section class="card">
            <h2>Primeiros passos</h2>
            <ol class="steps">
              <li class="${repos.length ? 'done' : ''}">
                <b>Cadastre os repositórios</b>
                <p class="muted small">Pastas locais (D:\\Dados) ou compartilhamentos de rede (\\\\servidor\\pasta).</p>
                <a href="#/repositorios">Ir para Repositórios</a>
              </li>
              <li class="${terms ? 'done' : ''}">
                <b>Monte a lista de referência</b>
                <p class="muted small">Nomes, palavras, códigos ou modelos prontos como CPF e CNPJ. Dá para importar de .txt ou .csv.</p>
                <a href="#/listas">Ir para Listas de referência</a>
              </li>
              <li>
                <b>Rode a análise</b>
                <p class="muted small">O relatório mostra o que foi encontrado, onde e o último usuário de cada arquivo.</p>
                <a href="#/analises/nova">Nova análise</a>
              </li>
            </ol>
          </section>`}

      <section class="card">
        <div class="card-head">
          <h2>Análises recentes</h2>
          <a href="#/analises">Ver todas</a>
        </div>
        ${scans.length === 0
          ? html`<div class="empty">Nenhuma análise realizada ainda.</div>`
          : html`<div class="table-wrap">
              <table class="data">
                <thead>
                  <tr><th>Análise</th><th>Situação</th><th>Início</th><th class="num">Arquivos verificados</th><th class="num">Com ocorrências</th></tr>
                </thead>
                <tbody>
                  ${scans.slice(0, 6).map(
                    (s) => html`<tr>
                      <td><a href="#/analises/${s.id}">${s.name}</a></td>
                      <td>${statusBadge(s.status)}</td>
                      <td class="nowrap">${fmtDateTime(s.startedAt || s.createdAt)}</td>
                      <td class="num">${fmtNum(s.stats?.filesSeen)}</td>
                      <td class="num">${fmtNum(s.stats?.filesMatched)}</td>
                    </tr>`,
                  )}
                </tbody>
              </table>
            </div>`}
      </section>`,
  );

  if (running.length === 0) return null;
  let stopped = false;
  const timer = setInterval(async () => {
    try {
      const latest = await get('/api/scans');
      if (!stopped && !latest.some((s) => s.status === 'running' || s.status === 'queued')) {
        clearInterval(timer);
        render(root);
      }
    } catch {
      // tenta de novo no próximo ciclo
    }
  }, 3000);
  return () => {
    stopped = true;
    clearInterval(timer);
  };
}
