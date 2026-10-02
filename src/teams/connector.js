// Microsoft Teams pela API Microsoft Graph. Usa o mesmo registro de aplicativo do Microsoft 365
// (as credenciais de uma conexão "Microsoft 365"), com as permissões do Teams consentidas pelo
// administrador: ChannelMessage.Read.All (mensagens de canais), Chat.Read.All (conversas),
// Team.ReadBasic.All e User.Read.All (listagens) — e ChannelMessage.ReadWrite.All para excluir.
//
// Percorre as mensagens dos canais das equipes (e as respostas) e as mensagens dos chats (1:1, em
// grupo e de reunião). O corpo (HTML) é convertido em texto; os arquivos anexados ficam no
// SharePoint/OneDrive e são baixados sob demanda para a análise do conteúdo.
import { GraphClient, commonGraphError, detailedGraphError, enc } from '../cloud/graph-client.js';
import { htmlToText } from '../scan/extractors/xml.js';
import { ApiError, pool } from '../mail/http.js';
import { folderMatcher, addressMatcher } from '../mail/common.js';

const MAX_CONCURRENCY = 4;
const PAGE = 50;

/** Identificador de compartilhamento do Graph para uma URL (GET /shares/{id}/driveItem). */
function shareId(url) {
  const b64 = Buffer.from(String(url), 'utf8').toString('base64').replace(/=+$/, '').replace(/\//g, '_').replace(/\+/g, '-');
  return `u!${b64}`;
}

/** Nome de exibição de um chat a partir do tipo, do tópico e dos participantes. */
function chatLabel(chat) {
  if (chat.topic) return chat.topic;
  const names = (chat.members || []).map((m) => m.displayName || m.email).filter(Boolean);
  if (chat.chatType === 'oneOnOne') return names.join(', ') || 'Conversa';
  if (chat.chatType === 'meeting') return `Reunião${names.length ? `: ${names.join(', ')}` : ''}`;
  return names.length ? `Grupo: ${names.join(', ')}` : 'Conversa em grupo';
}

export class TeamsConnector extends GraphClient {
  /**
   * source: conexão Microsoft 365 (graph) com os segredos decifrados, mais `teams` com o escopo da
   * análise: { scope ('all'|'list'), scanChannels, scanChats, teamIds, userEmails, excludeTeams,
   * excludeChannels, includeReplies }.
   */
  constructor(source, options = {}) {
    super(source, options);
    this.teams = source.teams || {};
  }

  translate(err) {
    if (err instanceof ApiError && err.status === 403) {
      return new ApiError(
        'Acesso negado pelo Microsoft Graph ao Teams: o aplicativo precisa das permissões do Teams com consentimento do administrador — ChannelMessage.Read.All (mensagens de canais) e Chat.Read.All (conversas), além de Team.ReadBasic.All e User.Read.All para as listagens (e ChannelMessage.ReadWrite.All para excluir).',
        err,
      );
    }
    return commonGraphError(err) || detailedGraphError(err);
  }

  // ----- listagens -----

  /** Equipes a analisar: todas do locatário (escopo "all") ou as da lista (por id). */
  async *teamList() {
    const excluded = new Set((this.teams.excludeTeams || []).map((t) => String(t).toLowerCase()));
    const keep = (t) => t.id && !excluded.has(t.id.toLowerCase()) && !excluded.has(String(t.displayName || '').toLowerCase());
    if (this.teams.scope === 'list') {
      for (const id of this.teams.teamIds || []) {
        try {
          const t = await this.api(`/teams/${enc(id)}?$select=id,displayName`);
          if (t?.id && keep(t)) yield { id: t.id, name: t.displayName || '' };
        } catch (err) {
          if (err.status === 404) continue; // equipe removida
          throw err;
        }
      }
      return;
    }
    let url = '/teams?$select=id,displayName&$top=100';
    while (url) {
      const page = await this.api(url);
      for (const t of page?.value || []) if (keep(t)) yield { id: t.id, name: t.displayName || '' };
      url = this.next(page);
    }
  }

  /** Canais de uma equipe (padrão, privados e compartilhados), sem os excluídos pelas opções. */
  async *channels(team) {
    const excluded = folderMatcher(this.teams.excludeChannels);
    let url = `/teams/${enc(team.id)}/channels?$select=id,displayName,membershipType&$top=${PAGE}`;
    while (url) {
      const page = await this.api(url);
      for (const c of page?.value || []) {
        if (excluded(c.displayName || '')) continue;
        yield { id: c.id, name: c.displayName || '', membershipType: c.membershipType || 'standard' };
      }
      url = this.next(page);
    }
  }

  /** Usuários do locatário (para enumerar os chats) — todos ou os da lista. */
  async *users() {
    if (this.teams.scope === 'list') {
      for (const address of this.teams.userEmails || []) {
        try {
          const u = await this.api(`/users/${enc(address)}?$select=id,displayName,mail,userPrincipalName`);
          if (u?.id) yield { id: u.id, address: u.mail || u.userPrincipalName || address, name: u.displayName || '' };
        } catch (err) {
          if (err.status === 404) continue;
          throw err;
        }
      }
      return;
    }
    const excluded = addressMatcher(this.teams.excludeUsers);
    let url = '/users?$select=id,displayName,mail,userPrincipalName&$top=999';
    while (url) {
      const page = await this.api(url);
      for (const u of page?.value || []) {
        const address = u.mail || u.userPrincipalName || '';
        if (address && !excluded(address)) yield { id: u.id, address, name: u.displayName || '' };
      }
      url = this.next(page);
    }
  }

  /**
   * Conversas a analisar, uma a uma: canais das equipes (kind 'channel') e chats (kind 'chat'). Os
   * chats são enumerados por usuário e não se repetem (o mesmo chat aparece para cada participante).
   * Um erro ao listar uma equipe/usuário vira uma conversa de erro (as demais continuam).
   */
  async *conversations() {
    if (this.teams.scanChannels !== false) {
      let teams;
      try {
        teams = this.teamList();
        for await (const team of teams) {
          let channels;
          try {
            channels = [];
            for await (const c of this.channels(team)) channels.push(c);
          } catch (err) {
            yield { error: this.translate(err), path: `Equipe "${team.name}"` };
            continue;
          }
          for (const channel of channels) {
            yield {
              id: `${team.id}|${channel.id}`,
              kind: 'channel',
              teamId: team.id,
              teamName: team.name,
              channelId: channel.id,
              channelName: channel.name,
              membershipType: channel.membershipType,
              path: `${team.name} › ${channel.name}`,
            };
          }
        }
      } catch (err) {
        yield { error: this.translate(err), path: 'equipes' };
      }
    }
    if (this.teams.scanChats === false) return;
    const seen = new Set();
    try {
      for await (const user of this.users()) {
        let chats;
        try {
          chats = [];
          let url = `/users/${enc(user.id)}/chats?$select=id,topic,chatType&$top=${PAGE}&$expand=members`;
          while (url) {
            const page = await this.api(url);
            for (const c of page?.value || []) chats.push(c);
            url = this.next(page);
          }
        } catch (err) {
          yield { error: this.translate(err), path: `chats de ${user.address}` };
          continue;
        }
        for (const chat of chats) {
          if (seen.has(chat.id)) continue;
          seen.add(chat.id);
          const members = (chat.members || []).map((m) => ({ displayName: m.displayName || '', email: m.email || '' }));
          yield {
            id: chat.id,
            kind: 'chat',
            chatId: chat.id,
            chatType: chat.chatType || 'group',
            topic: chat.topic || '',
            members,
            path: `Chat: ${chatLabel({ ...chat, members })}`,
          };
        }
      }
    } catch (err) {
      yield { error: this.translate(err), path: 'usuários' };
    }
  }

  // ----- mensagens -----

  /** Monta o registro de uma mensagem do Graph (chatMessage) para a análise. */
  message(conv, m, { replyTo = null } = {}) {
    const from = m.from?.user?.displayName || (m.from?.application?.displayName ? `${m.from.application.displayName} (aplicativo)` : '') || '';
    const html = (m.body?.contentType || '').toLowerCase() === 'html';
    const raw = m.body?.content || '';
    const text = html ? htmlToText(raw) : raw;
    // Anexos de arquivo (referência a um item do SharePoint/OneDrive) e outros (cartões etc.).
    const attachments = (m.attachments || []).map((a) => ({
      id: a.id,
      name: a.name || a.id,
      contentType: a.contentType || '',
      contentUrl: a.contentUrl || '',
      reference: (a.contentType || '').toLowerCase() === 'reference' && Boolean(a.contentUrl),
    }));
    return {
      id: m.id,
      replyTo,
      conv,
      from,
      fromId: m.from?.user?.id || '',
      subject: m.subject || '',
      text,
      date: m.createdDateTime || null,
      edited: m.lastEditedDateTime || m.lastModifiedDateTime || null,
      webUrl: m.webUrl || null,
      attachments,
      size: Buffer.byteLength(raw, 'utf8'),
    };
  }

  /**
   * Mensagens de uma conversa (canal ou chat), já convertidas. Nos canais, inclui as respostas de
   * cada mensagem raiz (se includeReplies). Entrega só mensagens de verdade (messageType 'message')
   * que não foram apagadas. since: só as criadas a partir desta data.
   */
  async *messages(conv, { since = null } = {}) {
    const real = (m) => m && m.messageType === 'message' && !m.deletedDateTime;
    const after = (m) => !since || (m.createdDateTime && new Date(m.createdDateTime) >= since);
    if (conv.kind === 'channel') {
      const base = `/teams/${enc(conv.teamId)}/channels/${enc(conv.channelId)}/messages`;
      let url = `${base}?$top=${PAGE}`;
      while (url) {
        const page = await this.api(url);
        for (const m of page?.value || []) {
          if (real(m) && after(m)) yield this.message(conv, m);
          if (this.teams.includeReplies !== false && Number(m.replies?.['@odata.count'] ?? 1) !== 0) {
            // Respostas da mensagem raiz (uma página costuma bastar; segue a paginação se houver).
            let rurl = `${base}/${enc(m.id)}/replies?$top=${PAGE}`;
            while (rurl) {
              const rp = await this.api(rurl);
              for (const r of rp?.value || []) if (real(r) && after(r)) yield this.message(conv, r, { replyTo: m.id });
              rurl = this.next(rp);
            }
          }
        }
        url = this.next(page);
      }
      return;
    }
    // Chat
    let url = `/chats/${enc(conv.chatId)}/messages?$top=${PAGE}`;
    while (url) {
      const page = await this.api(url);
      for (const m of page?.value || []) if (real(m) && after(m)) yield this.message(conv, m);
      url = this.next(page);
    }
  }

  // ----- anexos -----

  /**
   * Baixa o conteúdo de um anexo de arquivo (referência a um item do SharePoint/OneDrive), até
   * maxBytes. Resolve a URL de compartilhamento para o item e baixa o conteúdo.
   * Devolve { data, truncated, size } ou lança.
   */
  async downloadAttachment(att, { maxBytes = 50 * 1048576, signal = this.signal } = {}) {
    const item = await this.api(`/shares/${enc(shareId(att.contentUrl))}/driveItem?$select=id,name,size,parentReference`, { signal });
    const driveId = item?.parentReference?.driveId;
    if (!driveId || !item?.id) throw new ApiError('Anexo não encontrado no SharePoint/OneDrive.');
    return this.api(`/drives/${enc(driveId)}/items/${enc(item.id)}/content`, { type: 'buffer', maxBytes, retries: 4, signal, headers: { Accept: '*/*' } });
  }

  // ----- exclusão -----

  /**
   * Exclui (softDelete, recuperável) as mensagens encontradas em uma conversa. Vale apenas para
   * mensagens de canais; nos chats, a exclusão por aplicativo não é oferecida pelo Graph.
   * items: [{ id (recordId), messageId, replyTo, conv }]. Retorna Map(recordId → { ok, missing?, error? }).
   */
  async deleteMessages(items, { signal = this.signal, onResult, shouldStop } = {}) {
    const results = new Map();
    const done = (id, r) => {
      if (results.has(id)) return;
      results.set(id, r);
      onResult?.(id, r);
    };
    const run = pool(items, MAX_CONCURRENCY, async (item) => {
      if (shouldStop?.()) return undefined;
      done(item.recordId, await this.deleteOne(item, signal));
      return undefined;
    });
    for await (const _ of run); // eslint-disable-line no-unused-vars
    return results;
  }

  async deleteOne(item, signal) {
    const conv = item.conv;
    if (conv.kind !== 'channel') {
      return { ok: false, error: 'A exclusão de mensagens de chat não é oferecida pelo Microsoft Graph com permissões de aplicativo: exclua pelo próprio Teams.' };
    }
    const base = `/teams/${enc(conv.teamId)}/channels/${enc(conv.channelId)}/messages`;
    const path = item.replyTo ? `${base}/${enc(item.replyTo)}/replies/${enc(item.messageId)}/softDelete` : `${base}/${enc(item.messageId)}/softDelete`;
    let uncertain = false;
    const onRetry = ({ error }) => {
      if (error.status !== 429) uncertain = true;
    };
    try {
      await this.api(path, { method: 'POST', retries: 4, signal, onRetry });
      return { ok: true, note: 'excluída no Teams (recuperável por um administrador)' };
    } catch (err) {
      if (err.status === 404) return uncertain ? { ok: true } : { ok: false, missing: true, error: 'Mensagem não encontrada (já excluída).' };
      if (err.status === 403) {
        return { ok: false, error: 'Sem permissão para excluir: conceda ao aplicativo a permissão ChannelMessage.ReadWrite.All (tipo Aplicativo) com consentimento do administrador.' };
      }
      return { ok: false, error: this.translate(err).message };
    }
  }

  /** Teste da conexão: autenticação e uma amostra de equipes/usuários conforme o escopo. */
  async test() {
    const details = [];
    await this.accessToken();
    details.push(`Autenticação no Microsoft Entra ID (${this.auth.mode === 'certificate' ? 'certificado' : 'segredo do cliente'}): OK.`);
    if (this.teams.scanChannels !== false) {
      let count = 0;
      for await (const team of this.teamList()) {
        if (++count > 3) break;
        details.push(`Equipe "${team.name}": OK.`);
      }
      details.push(count ? `Equipes acessíveis: ${count}${count > 3 ? '+' : ''}.` : 'Nenhuma equipe encontrada no escopo.');
    }
    if (this.teams.scanChats !== false) {
      let count = 0;
      for await (const user of this.users()) {
        if (++count > 3) break;
      }
      details.push(count ? `Usuários para chats: ${count}${count > 3 ? '+' : ''}.` : 'Nenhum usuário encontrado para os chats.');
    }
    return { ok: true, message: 'Conexão com o Microsoft Teams funcionando.', details };
  }

  async close() {}
}
