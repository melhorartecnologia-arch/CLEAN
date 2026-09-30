// Microsoft 365 / Exchange Online (e Outlook.com) pela API Microsoft Graph. Dois modos de acesso:
// - aplicativo (sem usuário conectado): um registro de aplicativo no Microsoft Entra ID com Mail.Read e
//   User.Read.All e um segredo do cliente ou um certificado — alcança todas as caixas do locatário;
// - conta conectada (OAuth delegado): a caixa da conta que entrou e as caixas compartilhadas com ela
//   (Mail.Read e Mail.Read.Shared).
// Cada mensagem é baixada no formato MIME original (com os anexos).
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

/** Destinatário do Graph ({ emailAddress: { name, address } }) no formato { name, address }. */
const mapRecipient = (r) => ({ name: r?.emailAddress?.name || '', address: r?.emailAddress?.address || '' });

/**
 * Um usuário do Microsoft Entra ID no formato do catálogo de contas: endereço principal, apelidos
 * (proxyAddresses smtp:), situação, tipo, licença, data de criação e dados do cadastro.
 */
function graphAccount(u) {
  const proxies = (u.proxyAddresses || []).filter((p) => /^smtp:/i.test(p));
  const primary = proxies.find((p) => p.startsWith('SMTP:'))?.slice(5) || u.mail || '';
  const aliases = proxies
    .filter((p) => p.startsWith('smtp:'))
    .map((p) => p.slice(5))
    .filter((a) => a && a.toLowerCase() !== primary.toLowerCase());
  const phones = [u.mobilePhone, ...(u.businessPhones || [])].filter(Boolean);
  return {
    address: u.mail || u.userPrincipalName || '',
    name: u.displayName || '',
    login: u.userPrincipalName || '',
    aliases,
    enabled: u.accountEnabled == null ? null : Boolean(u.accountEnabled),
    type: u.userType || '',
    licensed: (u.assignedLicenses || []).length > 0,
    created: u.createdDateTime || null,
    lastActivity: null, // exige AuditLog.Read.All (signInActivity, beta): fora do catálogo
    department: u.department || '',
    title: u.jobTitle || '',
    location: u.officeLocation || '',
    phone: phones.join(', '),
    orgUnit: '',
    admin: null,
  };
}

/** Traduz erros do Graph/Entra ID para mensagens com a providência a tomar. delegated: conta conectada. */
export function graphError(err, { delegated = false } = {}) {
  if (!(err instanceof ApiError)) return err;
  const common = commonGraphError(err);
  if (common) return common;
  if (err.status === 403 && /AccessDenied/i.test(err.code)) {
    if (delegated) {
      return new ApiError(
        'Acesso negado à caixa de correio: a conta conectada não tem acesso a ela. Numa caixa compartilhada ou de outra pessoa, a conta precisa da permissão de Acesso Total (Exchange Online) — e a entrada precisa ter autorizado as caixas compartilhadas (Mail.Read.Shared, pedida quando o locatário é informado).',
        err,
      );
    }
    return new ApiError(
      'Acesso negado à caixa de correio: conceda ao aplicativo a permissão Mail.Read (tipo Aplicativo) com consentimento do administrador e confira se uma política de acesso do Exchange Online (RBAC para aplicativos ou Application Access Policy) libera esta caixa.',
      err,
    );
  }
  return detailedGraphError(err);
}

export class GraphConnector extends GraphClient {
  translate(err) {
    return graphError(err, { delegated: this.delegated });
  }

  /**
   * Conta conectada: a caixa informada não existe para o Microsoft Graph (endereço errado ou
   * diferente do nome de logon) ou não tem caixa no Exchange Online. É um erro (e não uma caixa
   * ignorada): quem cadastrou a lista pediu para analisá-la.
   */
  missingMailbox(address, err) {
    const own = this.ownMailbox(address);
    const message = own
      ? `A conta conectada (${address}) não tem caixa de correio no Exchange Online (sem licença, desativada ou local).`
      : `A caixa ${address} não foi encontrada para a conta conectada: confira o endereço (use o endereço principal da caixa ou o nome de logon dela no Microsoft 365) e se a conta tem acesso a ela.`;
    return new ApiError(message, err);
  }

  /** A caixa é a da própria conta conectada (acessada por /me). */
  ownMailbox(address) {
    const account = this.source.graph?.account || {};
    const key = String(address || '').trim().toLowerCase();
    return Boolean(key) && [account.address, account.username].some((a) => String(a || '').toLowerCase() === key);
  }

  /**
   * Caixas a analisar: todas as do locatário (usuários com e-mail) ou as da lista. Com a conta
   * conectada, sempre as da lista (a da conta e as compartilhadas com ela).
   */
  async mailboxes() {
    const excluded = addressMatcher(this.source.excludeMailboxes);
    if (this.source.scope !== 'all' || this.delegated) {
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

  /**
   * Catálogo de contas do domínio (listagem): todos os usuários do locatário com os dados do cadastro
   * (nome, endereço principal, apelidos, situação, tipo, licença, data de criação, departamento etc.).
   * Exige a permissão User.Read.All (aplicativo ou, com a conta conectada, delegada).
   */
  async *directory() {
    const select =
      'id,displayName,userPrincipalName,mail,proxyAddresses,accountEnabled,userType,createdDateTime,department,jobTitle,officeLocation,mobilePhone,businessPhones,assignedLicenses';
    let url = `/users?$select=${select}&$top=999`;
    while (url) {
      const page = await this.api(url);
      for (const u of page?.value || []) {
        const acct = graphAccount(u);
        if (acct.address) yield acct;
      }
      url = this.next(page);
    }
  }

  /**
   * Dono da caixa: { id, name, path } — path é o início dos endereços da caixa no Graph. Com um
   * aplicativo, o usuário é localizado pelo e-mail (pode ser diferente do nome de logon); com a conta
   * conectada, a caixa dela é "/me" e as compartilhadas são acessadas pelo endereço.
   */
  async resolveUser(mailbox, { signal } = {}) {
    if (this.delegated) {
      if (this.ownMailbox(mailbox.address)) return { id: 'me', name: this.source.graph?.account?.name || '', path: '/me' };
      return { id: mailbox.address, name: mailbox.name || '', path: `/users/${enc(mailbox.address)}` };
    }
    const user = await super.resolveUser(mailbox, { signal, notFound: `A caixa ${mailbox.address} não foi encontrada no Microsoft 365.` });
    return { ...user, path: `/users/${enc(user.id)}` };
  }

  async wellKnownFolder(base, name) {
    try {
      return (await this.api(`${base}/mailFolders/${name}?$select=id`))?.id || null;
    } catch (err) {
      // Pasta que não existe na caixa (ex.: sem Lixo Eletrônico). Com a conta conectada, o endereço da
      // caixa não foi conferido antes: "recurso não encontrado" é a própria caixa.
      if (err.status === 404 && !NO_MAILBOX.has(err.code) && !(this.delegated && /ResourceNotFound/i.test(err.code))) return null;
      throw err;
    }
  }

  /**
   * Pastas da caixa (com subpastas), sem as excluídas pelas opções e pelos padrões da conexão.
   * base: início dos endereços da caixa (resolveUser().path).
   */
  async folders(base, { includeTrash = true, includeJunk = false, address = '' } = {}) {
    const skip = new Set();
    let trash = null;
    try {
      trash = await this.wellKnownFolder(base, 'deleteditems');
      const junk = await this.wellKnownFolder(base, 'junkemail');
      const sync = await this.wellKnownFolder(base, 'syncissues');
      if (!includeTrash && trash) skip.add(trash);
      if (!includeJunk && junk) skip.add(junk);
      if (sync) skip.add(sync); // registros de sincronização do Outlook
    } catch (err) {
      if (this.delegated && (NO_MAILBOX.has(err.code) || err.status === 404)) throw this.missingMailbox(address, err);
      if (NO_MAILBOX.has(err.code) || err.status === 404) throw new SkipMailboxError('usuário sem caixa de correio no Exchange Online (sem licença, desativada ou local).');
      // Com "todas as caixas" e o acesso limitado pelo RBAC para aplicativos, as caixas fora do
      // escopo respondem 403: são ignoradas (com aviso), não contadas como erro.
      if (err.status === 403 && this.source.scope === 'all' && !this.delegated) throw new SkipMailboxError('acesso à caixa não liberado para o aplicativo (permissão Mail.Read ou escopo do RBAC para aplicativos).');
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
          if (f.childFolderCount > 0) await walk(`${base}/mailFolders/${enc(f.id)}/childFolders?${fields}`, path, inTrash);
        }
        next = this.next(page);
      }
    };
    await walk(`${base}/mailFolders?${fields}`, '', false);
    return out;
  }

  /**
   * Mensagens da caixa, já baixadas (MIME), com até `concurrency` downloads simultâneos.
   * Entrega { folder, id, raw, truncated, size, receivedAt, webLink } ou { folder, id, error }.
   * before: só as recebidas antes desta data; headersOnly: sem baixar a mensagem (retenção) —
   * entrega também { subject, from, internetMessageId, headersOnly: true }.
   */
  async *messages(mailbox, { since = null, before = null, headersOnly = false, fullHeaders = false, includeTrash = true, includeJunk = false, maxBytes = 50 * 1048576, concurrency = 4, onFolder } = {}) {
    const user = await this.resolveUser(mailbox);
    mailbox.name ||= user.name;
    const folders = await this.folders(user.path, { includeTrash, includeJunk, address: mailbox.address });
    const userPath = user.path;
    const conditions = [since && `receivedDateTime ge ${since.toISOString()}`, before && `receivedDateTime lt ${before.toISOString()}`].filter(Boolean);
    const filter = conditions.length ? `&$filter=${enc(conditions.join(' and '))}` : '';
    // A listagem de mensagens (fullHeaders) traz também destinatários, data de envio e se há anexos.
    const fields = headersOnly
      ? fullHeaders
        ? 'id,receivedDateTime,sentDateTime,webLink,subject,from,toRecipients,ccRecipients,internetMessageId,hasAttachments'
        : 'id,receivedDateTime,webLink,subject,from,internetMessageId'
      : 'id,receivedDateTime,webLink';
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
                  ? {
                      subject: m.subject || '',
                      from: from ? { name: from.name || '', address: from.address || '' } : null,
                      internetMessageId: m.internetMessageId || null,
                      inTrash: folder.inTrash,
                      headersOnly: true,
                      ...(fullHeaders
                        ? {
                            sent: m.sentDateTime || null,
                            to: (m.toRecipients || []).map(mapRecipient),
                            cc: (m.ccRecipients || []).map(mapRecipient),
                            hasAttachments: Boolean(m.hasAttachments),
                          }
                        : {}),
                    }
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
    if (this.delegated) {
      const me = await this.api('/me?$select=displayName,mail,userPrincipalName');
      details.push(`Conta conectada: ${me?.displayName ? `${me.displayName} — ` : ''}${me?.mail || me?.userPrincipalName || '—'} (token renovado: OK).`);
    } else {
      details.push(`Autenticação no Microsoft Entra ID (${this.auth.mode === 'certificate' ? 'certificado' : 'segredo do cliente'}): OK.`);
    }
    let boxes;
    if (this.source.scope === 'all' && !this.delegated) {
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
        const inbox = await this.api(`${user.path}/mailFolders/inbox?$select=displayName,totalItemCount`);
        details.push(`${box.address}: ${Number(inbox?.totalItemCount) || 0} mensagem(ns) em "${inbox?.displayName || 'Caixa de Entrada'}".`);
        checked++;
      } catch (err) {
        // Em "todas as caixas", usuários sem caixa ou fora do escopo do RBAC são pulados.
        if (this.delegated && (NO_MAILBOX.has(err.code) || err.status === 404)) return { ok: false, message: this.missingMailbox(box.address, err).message, details };
        if (this.source.scope === 'all' && !this.delegated && (NO_MAILBOX.has(err.code) || err.status === 404 || err.status === 403)) {
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
    const userPath = user.path;
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
        if (this.delegated) {
          return {
            ok: false,
            error: 'Sem permissão para excluir: conecte a conta de novo com "Permitir excluir" marcado (permissão Mail.ReadWrite; nas caixas compartilhadas, Mail.ReadWrite.Shared e Acesso Total à caixa).',
          };
        }
        return { ok: false, error: 'Sem permissão para excluir: conceda ao aplicativo a permissão Mail.ReadWrite (tipo Aplicativo) ou a função "Application Mail.ReadWrite" do RBAC para aplicativos.' };
      }
      return { ok: false, error: err.message };
    }
  }

  async close() {}
}
