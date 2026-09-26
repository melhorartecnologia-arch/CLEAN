// Microsoft 365 / Exchange Online pela API Microsoft Graph, com permissões de aplicativo (sem usuário
// conectado): um registro de aplicativo no Microsoft Entra ID com Mail.Read e User.Read.All e um
// segredo do cliente. Cada mensagem é baixada no formato MIME original (com os anexos).
import { pool, ApiError } from './http.js';
import { SkipMailboxError, folderMatcher, addressMatcher, deletionItems } from './common.js';
import { GraphClient, GRAPH_ENDPOINTS, commonGraphError, detailedGraphError, enc } from '../cloud/graph-client.js';

// O Exchange Online aceita até 4 requisições simultâneas por caixa para cada aplicativo.
const MAX_CONCURRENCY = 4;

// Usuário sem caixa no Exchange Online (sem licença, caixa desativada ou hospedada localmente).
const NO_MAILBOX = new Set(['MailboxNotEnabledForRESTAPI', 'MailboxNotHostedInExchangeOnline', 'ErrorNonExistentMailbox', 'ErrorInvalidUser']);

// Identificadores imutáveis: continuam válidos se a mensagem for movida de pasta depois da análise
// (necessário para a exclusão manual pelo relatório).
const IMMUTABLE_IDS = { Prefer: 'IdType="ImmutableId"' };

export { GRAPH_ENDPOINTS };

/** Traduz erros do Graph/Entra ID para mensagens com a providência a tomar. */
export function graphError(err) {
  if (!(err instanceof ApiError)) return err;
  const common = commonGraphError(err);
  if (common) return common;
  if (err.status === 403 && /AccessDenied/i.test(err.code)) {
    return new ApiError(
      'Acesso negado à caixa de correio: conceda ao aplicativo a permissão Mail.Read (tipo Aplicativo) com consentimento do administrador e confira se uma política de acesso do Exchange Online (RBAC para aplicativos ou Application Access Policy) libera esta caixa.',
      err,
    );
  }
  return detailedGraphError(err);
}

export class GraphConnector extends GraphClient {
  translate(err) {
    return graphError(err);
  }

  /** Caixas a analisar: todas as do locatário (usuários com e-mail) ou as da lista. */
  async mailboxes() {
    const excluded = addressMatcher(this.source.excludeMailboxes);
    if (this.source.scope !== 'all') {
      return (this.source.mailboxes || []).map((m) => ({ address: m.address, name: '' })).filter((m) => !excluded(m.address));
    }
    const out = [];
    let url = '/users?$select=id,displayName,mail,userPrincipalName&$top=999';
    while (url) {
      const page = await this.api(url);
      for (const u of page?.value || []) if (u.mail && !excluded(u.mail)) out.push({ address: u.mail, name: u.displayName || '', id: u.id });
      url = this.next(page);
    }
    return out.sort((a, b) => a.address.localeCompare(b.address));
  }

  /** Identificador do usuário dono da caixa (o endereço pode ser diferente do nome de logon). */
  resolveUser(mailbox, { signal } = {}) {
    return super.resolveUser(mailbox, { signal, notFound: `A caixa ${mailbox.address} não foi encontrada no Microsoft 365.` });
  }

  async wellKnownFolder(userId, name) {
    try {
      return (await this.api(`/users/${enc(userId)}/mailFolders/${name}?$select=id`))?.id || null;
    } catch (err) {
      if (err.status === 404 && !NO_MAILBOX.has(err.code)) return null;
      throw err;
    }
  }

  /** Pastas da caixa (com subpastas), sem as excluídas pelas opções e pelos padrões da conexão. */
  async folders(userId, { includeTrash = true, includeJunk = false } = {}) {
    const skip = new Set();
    let trash = null;
    try {
      trash = await this.wellKnownFolder(userId, 'deleteditems');
      const junk = await this.wellKnownFolder(userId, 'junkemail');
      const sync = await this.wellKnownFolder(userId, 'syncissues');
      if (!includeTrash && trash) skip.add(trash);
      if (!includeJunk && junk) skip.add(junk);
      if (sync) skip.add(sync); // registros de sincronização do Outlook
    } catch (err) {
      if (NO_MAILBOX.has(err.code) || err.status === 404) throw new SkipMailboxError('usuário sem caixa de correio no Exchange Online (sem licença, desativada ou local).');
      // Com "todas as caixas" e o acesso limitado pelo RBAC para aplicativos, as caixas fora do
      // escopo respondem 403: são ignoradas (com aviso), não contadas como erro.
      if (err.status === 403 && this.source.scope === 'all') throw new SkipMailboxError('acesso à caixa não liberado para o aplicativo (permissão Mail.Read ou escopo do RBAC para aplicativos).');
      throw err;
    }
    const excluded = folderMatcher(this.source.excludeFolders);
    const fields = '$select=id,displayName,childFolderCount,totalItemCount&$top=250';
    const out = [];
    // inTrash: a pasta Itens Excluídos e as subpastas dela.
    const walk = async (url, parent, parentInTrash) => {
      for (let next = url; next; ) {
        const page = await this.api(next);
        for (const f of page?.value || []) {
          const path = parent ? `${parent}/${f.displayName}` : f.displayName;
          if (skip.has(f.id) || excluded(path)) continue;
          const inTrash = parentInTrash || (trash !== null && f.id === trash);
          out.push({ id: f.id, path, total: Number(f.totalItemCount) || 0, inTrash });
          if (f.childFolderCount > 0) await walk(`/users/${enc(userId)}/mailFolders/${enc(f.id)}/childFolders?${fields}`, path, inTrash);
        }
        next = this.next(page);
      }
    };
    await walk(`/users/${enc(userId)}/mailFolders?${fields}`, '', false);
    return out;
  }

  /**
   * Mensagens da caixa, já baixadas (MIME), com até `concurrency` downloads simultâneos.
   * Entrega { folder, id, raw, truncated, size, receivedAt, webLink } ou { folder, id, error }.
   * before: só as recebidas antes desta data; headersOnly: sem baixar a mensagem (retenção) —
   * entrega também { subject, from, internetMessageId, headersOnly: true }.
   */
  async *messages(mailbox, { since = null, before = null, headersOnly = false, includeTrash = true, includeJunk = false, maxBytes = 50 * 1048576, concurrency = 4, onFolder } = {}) {
    const user = await this.resolveUser(mailbox);
    mailbox.name ||= user.name;
    const folders = await this.folders(user.id, { includeTrash, includeJunk });
    const userPath = `/users/${enc(user.id)}`;
    const conditions = [since && `receivedDateTime ge ${since.toISOString()}`, before && `receivedDateTime lt ${before.toISOString()}`].filter(Boolean);
    const filter = conditions.length ? `&$filter=${enc(conditions.join(' and '))}` : '';
    const fields = headersOnly ? 'id,receivedDateTime,webLink,subject,from,internetMessageId' : 'id,receivedDateTime,webLink';
    const size = `&$expand=${enc("singleValueExtendedProperties($filter=id eq 'Integer 0x0E08')")}`;
    const self = this;
    async function* list() {
      for (const folder of folders) {
        onFolder?.(folder.path);
        if (!folder.total) continue;
        let url = `${userPath}/mailFolders/${enc(folder.id)}/messages?$select=${fields}&$top=100${size}${filter}`;
        try {
          while (url) {
            const page = await self.api(url, { headers: IMMUTABLE_IDS });
            for (const m of page?.value || []) {
              const from = m.from?.emailAddress;
              yield {
                folder: folder.path,
                id: m.id,
                receivedAt: m.receivedDateTime || null,
                webLink: m.webLink || null,
                size: Number(m.singleValueExtendedProperties?.[0]?.value) || 0,
                ...(headersOnly
                  ? { subject: m.subject || '', from: from ? { name: from.name || '', address: from.address || '' } : null, internetMessageId: m.internetMessageId || null, inTrash: folder.inTrash, headersOnly: true }
                  : {}),
              };
            }
            url = self.next(page);
          }
        } catch (err) {
          // Uma pasta que não pôde ser lida (removida durante a análise, erro do serviço...) vira
          // erro dela: as demais pastas da caixa continuam.
          if (self.signal?.aborted) throw err;
          yield { folder: folder.path, id: null, error: err };
        }
      }
    }
    if (headersOnly) {
      yield* list();
      return;
    }
    yield* pool(list(), Math.max(1, Math.min(concurrency, MAX_CONCURRENCY)), async (item) => {
      if (item.error) return item;
      try {
        const res = await this.api(`${userPath}/messages/${enc(item.id)}/$value`, { type: 'buffer', maxBytes, retries: 4, headers: { Accept: '*/*', ...IMMUTABLE_IDS } });
        return { ...item, raw: res.data, truncated: res.truncated, size: item.size || res.size };
      } catch (err) {
        if (this.signal?.aborted) throw err;
        return { ...item, raw: null, error: err };
      }
    });
  }

  /** Teste da conexão: autenticação, listagem de usuários (se for o caso) e acesso a até 3 caixas. */
  async test() {
    const details = [];
    await this.accessToken();
    details.push('Autenticação no Microsoft Entra ID: OK.');
    let boxes;
    if (this.source.scope === 'all') {
      const page = await this.api('/users?$select=id,displayName,mail&$top=50');
      boxes = (page?.value || []).filter((u) => u.mail).map((u) => ({ address: u.mail, id: u.id, name: u.displayName }));
      details.push(`Listagem de usuários (User.Read.All): OK — ${boxes.length}${page?.['@odata.nextLink'] ? '+' : ''} com e-mail.`);
      if (boxes.length === 0) return { ok: false, message: 'Nenhum usuário com e-mail encontrado no locatário.', details };
    } else {
      boxes = (this.source.mailboxes || []).map((m) => ({ address: m.address }));
      if (boxes.length === 0) return { ok: false, message: 'Informe ao menos uma caixa de e-mail.', details };
    }
    let checked = 0;
    let skipped = null;
    for (const box of boxes) {
      if (checked >= 3) break;
      try {
        const user = await this.resolveUser(box);
        const inbox = await this.api(`/users/${enc(user.id)}/mailFolders/inbox?$select=displayName,totalItemCount`);
        details.push(`${box.address}: ${Number(inbox?.totalItemCount) || 0} mensagem(ns) em "${inbox?.displayName || 'Caixa de Entrada'}".`);
        checked++;
      } catch (err) {
        // Em "todas as caixas", usuários sem caixa ou fora do escopo do RBAC são pulados.
        if (this.source.scope === 'all' && (NO_MAILBOX.has(err.code) || err.status === 404 || err.status === 403)) {
          skipped = err;
          continue;
        }
        return { ok: false, message: `${box.address}: ${err.message}`, details };
      }
    }
    if (checked === 0) {
      const reason = skipped ? ` Último erro: ${skipped.message}` : '';
      return { ok: false, message: `Nenhuma das caixas testadas está disponível para o aplicativo.${reason}`, details };
    }
    return { ok: true, message: 'Conexão com o Microsoft 365 funcionando.', details };
  }

  /**
   * Exclui mensagens da caixa: 'permanent' = exclusão definitiva (permanentDelete: a mensagem vai
   * para a área de expurgo e some para o usuário; retenções e bloqueios de litígio continuam
   * valendo), 'trash' = move para Itens Excluídos. Exige a permissão Mail.ReadWrite.
   * items: ids ou { id }. options: signal (padrão: o da conexão; a análise usa null para que as
   * exclusões em andamento terminem mesmo ao cancelar), onResult(id, resultado) a cada mensagem e
   * shouldStop() (não começa outras). Retorna Map(id → { ok, missing?, error? }).
   */
  async deleteMessages(mailbox, items, mode = 'permanent', { signal = this.signal, onResult, shouldStop } = {}) {
    const results = new Map();
    const list = deletionItems(items);
    if (list.length === 0) return results;
    const user = await this.resolveUser(mailbox, { signal });
    const userPath = `/users/${enc(user.id)}`;
    const run = pool(list, MAX_CONCURRENCY, async ({ id }) => {
      if (shouldStop?.()) return undefined;
      const result = await this.deleteOne(userPath, id, mode, signal);
      results.set(id, result);
      onResult?.(id, result);
      return undefined;
    });
    for await (const _ of run); // eslint-disable-line no-unused-vars
    return results;
  }

  async deleteOne(userPath, id, mode, signal) {
    // Uma tentativa interrompida (falha de rede, tempo esgotado, erro 5xx) pode ter sido feita pelo
    // servidor: nesse caso, "não encontrada" na repetição quer dizer que a exclusão funcionou.
    let uncertain = false;
    const onRetry = ({ error }) => {
      if (error.status !== 429) uncertain = true;
    };
    const options = { method: 'POST', headers: IMMUTABLE_IDS, retries: 4, signal, onRetry };
    try {
      if (mode === 'trash') await this.api(`${userPath}/messages/${enc(id)}/move`, { ...options, json: { destinationId: 'deleteditems' } });
      else await this.api(`${userPath}/messages/${enc(id)}/permanentDelete`, options);
      return { ok: true };
    } catch (err) {
      if (err.status === 404 && /ErrorItemNotFound/i.test(err.code)) {
        return uncertain ? { ok: true } : { ok: false, missing: true, error: 'Mensagem não encontrada (já excluída ou movida).' };
      }
      if (err.status === 403) {
        return { ok: false, error: 'Sem permissão para excluir: conceda ao aplicativo a permissão Mail.ReadWrite (tipo Aplicativo) ou a função "Application Mail.ReadWrite" do RBAC para aplicativos.' };
      }
      return { ok: false, error: err.message };
    }
  }

  async close() {}
}
