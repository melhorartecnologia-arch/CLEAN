// Painel inicial: números gerais, primeiros passos e análises recentes (arquivos e e-mail).
import { get } from '../api.js';
import { html, render as paint, icon, fmtNum, fmtCompact, fmtDateTime, fmtServerDateTime, statusBadge, plural } from '../ui.js';
import { zoneNote } from '../schedule-form.js';

const isMail = (s) => s.kind === 'mail';
const running = (s) => s.status === 'running' || s.status === 'queued';

export async function render(root, { ctx }) {
  const [repos, lists, sources, scans, schedules, policies] = await Promise.all([
    get('/api/repositories'),
    get('/api/lists'),
    get('/api/mail-sources'),
    get('/api/scans'),
    get('/api/schedules'),
    get('/api/schedules?purpose=retention'),
  ]);
  // Agendamentos de análise e políticas de retenção agendadas.
  const planned = [...schedules, ...policies];
  const upcoming = planned.filter((s) => s.state === 'active').sort((a, b) => String(a.nextRunAt).localeCompare(String(b.nextRunAt)));
  const attention = planned.filter((s) => s.problems.length || s.lastRun?.status === 'failed');
  const isPolicy = (s) => s.purpose === 'retention';
  const editLink = (s) => `${isPolicy(s) ? '#/retencao' : '#/agendamentos'}/${s.id}`;
  const terms = lists.reduce((sum, l) => sum + l.termCount, 0);
  // Os números de ocorrências vêm das análises por termos (as execuções da retenção não têm termos).
  const lastFiles = scans.find((s) => !isMail(s) && !s.retention && s.status === 'completed');
  const lastMail = scans.find((s) => isMail(s) && !s.retention && s.status === 'completed');
  const active = scans.filter(running);
  const ready = (repos.length > 0 || sources.length > 0) && terms > 0;
  const link = (s) => `${isMail(s) ? '#/email/analises' : '#/analises'}/${s.id}`;

  paint(
    root,
    html`<div class="page-head">
        <div>
          <h1>Painel</h1>
          <div class="sub">Procure os termos das listas de referência em arquivos e caixas de e-mail e veja quem interagiu com cada item encontrado.</div>
        </div>
        <div class="actions">
          <a class="btn primary" href="#/analises/nova">${icon('play')} Analisar arquivos</a>
          <a class="btn primary" href="#/email/analises/nova">${icon('play')} Analisar e-mails</a>
        </div>
      </div>

      <section class="tiles" aria-label="Resumo">
        <div class="tile"><div class="label">Repositórios</div><div class="value">${fmtNum(repos.length)}</div><div class="detail">pastas e compartilhamentos</div></div>
        <div class="tile"><div class="label">Caixas de e-mail</div><div class="value">${fmtNum(sources.length)}</div><div class="detail">${plural(sources.length, 'conexão', 'conexões')} cadastrada${sources.length === 1 ? '' : 's'}</div></div>
        <div class="tile"><div class="label">Listas de referência</div><div class="value">${fmtNum(lists.length)}</div><div class="detail">${plural(terms, 'termo', 'termos')}</div></div>
        <div class="tile">
          <div class="label">Arquivos com ocorrências</div>
          <div class="value">${lastFiles ? fmtCompact(lastFiles.stats?.filesMatched) : '—'}</div>
          <div class="detail">${lastFiles ? `na última análise de arquivos (${fmtDateTime(lastFiles.finishedAt)})` : 'nenhuma análise de arquivos concluída'}</div>
        </div>
        <div class="tile">
          <div class="label">Mensagens com ocorrências</div>
          <div class="value">${lastMail ? fmtCompact(lastMail.stats?.messagesMatched) : '—'}</div>
          <div class="detail">${lastMail ? `na última análise de e-mail (${fmtDateTime(lastMail.finishedAt)})` : 'nenhuma análise de e-mail concluída'}</div>
        </div>
      </section>

      ${ready
        ? ''
        : html`<section class="card">
            <h2>Primeiros passos</h2>
            <ol class="steps">
              <li class="${repos.length || sources.length ? 'done' : ''}">
                <b>Cadastre onde procurar</b>
                <p class="muted small">Pastas locais (D:\\Dados) ou compartilhamentos de rede (\\\\servidor\\pasta), e caixas de e-mail do Microsoft 365, do Google Workspace ou de um servidor IMAP.</p>
                <a href="#/repositorios">Repositórios</a> · <a href="#/email/caixas">Caixas de e-mail</a>
              </li>
              <li class="${terms ? 'done' : ''}">
                <b>Monte a lista de referência</b>
                <p class="muted small">Nomes, palavras, códigos ou modelos prontos como CPF e CNPJ. Dá para importar de .txt ou .csv.</p>
                <a href="#/listas">Ir para Listas de referência</a>
              </li>
              <li>
                <b>Rode a análise</b>
                <p class="muted small">O relatório mostra o que foi encontrado, onde e quem interagiu: o último usuário de cada arquivo ou o remetente e a caixa de cada mensagem.</p>
                <a href="#/analises/nova">Analisar arquivos</a> · <a href="#/email/analises/nova">Analisar e-mails</a>
              </li>
            </ol>
          </section>`}

      ${planned.length
        ? html`<section class="card">
            <div class="card-head">
              <h2>Próximas execuções agendadas</h2>
              <span class="small">${schedules.length ? html`<a href="#/agendamentos">Agendamentos</a>` : ''}${schedules.length && policies.length ? ' · ' : ''}${policies.length ? html`<a href="#/retencao">Retenção</a>` : ''}</span>
            </div>
            <p class="muted small">${zoneNote(ctx.info)}</p>
            ${attention.length
              ? html`<div class="alert">${icon('alert')}<div><b>Precisa${attention.length > 1 ? 'm' : ''} de atenção:</b> ${attention.map((s, i) => html`${i ? ', ' : ''}<a href="${editLink(s)}">${s.name}</a>${isPolicy(s) ? ' (retenção)' : ''}`)}.</div></div>`
              : ''}
            ${upcoming.length
              ? html`<ul class="upcoming">
                  ${upcoming.slice(0, 5).map(
                    (s) => html`<li>
                      <span class="nowrap"><b>${fmtServerDateTime(s.nextRunAt)}</b></span>
                      <a href="${editLink(s)}">${s.name}</a>
                      <span class="kind-badge">${s.kind === 'mail' ? 'E-mail' : 'Arquivos'}</span>
                      ${isPolicy(s) ? html`<span class="chip">retenção</span>` : ''}
                      ${s.action === 'delete' ? html`<span class="chip danger">${isPolicy(s) ? 'exclui os expirados' : 'exclusão automática'}</span>` : ''}
                    </li>`,
                  )}
                </ul>`
              : html`<p class="muted">Nenhuma execução agendada.</p>`}
          </section>`
        : ''}

      <section class="card">
        <div class="card-head">
          <h2>Análises recentes</h2>
          <span class="small"><a href="#/analises">Arquivos</a> · <a href="#/email/analises">E-mail</a></span>
        </div>
        ${scans.length === 0
          ? html`<div class="empty">Nenhuma análise realizada ainda.</div>`
          : html`<div class="table-wrap">
              <table class="data">
                <thead>
                  <tr><th>Análise</th><th>Tipo</th><th>Situação</th><th>Início</th><th class="num">Itens verificados</th><th class="num">Com ocorrências</th></tr>
                </thead>
                <tbody>
                  ${scans.slice(0, 8).map(
                    (s) => html`<tr>
                      <td><a href="${link(s)}">${s.name}</a></td>
                      <td><span class="kind-badge">${isMail(s) ? 'E-mail' : 'Arquivos'}</span>${s.retention ? html` <span class="chip">retenção</span>` : ''}</td>
                      <td>${statusBadge(s.status)}</td>
                      <td class="nowrap">${fmtDateTime(s.startedAt || s.createdAt)}</td>
                      <td class="num">${fmtNum(isMail(s) ? s.stats?.messagesSeen : s.stats?.filesSeen)}</td>
                      <td class="num">${fmtNum(isMail(s) ? s.stats?.messagesMatched : s.stats?.filesMatched)}${s.retention ? html`<div class="muted small">expirad${isMail(s) ? 'as' : 'os'}</div>` : ''}</td>
                    </tr>`,
                  )}
                </tbody>
              </table>
            </div>`}
      </section>`,
  );

  if (active.length === 0) return null;
  let stopped = false;
  const timer = setInterval(async () => {
    try {
      const latest = await get('/api/scans');
      if (!stopped && !latest.some(running)) {
        clearInterval(timer);
        render(root, { ctx });
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
