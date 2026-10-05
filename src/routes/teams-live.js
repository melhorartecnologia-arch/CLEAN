// Rotas /api/teams-live: leitura ao vivo (somente leitura) das conversas do Microsoft Teams de um
// usuário — chats e canais das equipes de que ele participa. Cada requisição fala com o Microsoft
// Graph na hora, com um conector criado com os segredos decifrados da conexão Microsoft 365 e um
// tempo limite curto. Nada é gravado; os segredos nunca voltam para o navegador.
import { Router } from 'express';
import { HttpError } from './validate.js';
import { TeamsConnector } from '../teams/connector.js';
import { friendlyError } from '../scan/errors.js';

export function teamsLiveRouter({ store, endpoints = {} }) {
  const router = Router();

  const graphSource = (id) => {
    const s = store.getMailSource(id);
    if (!s) throw new HttpError(404, 'Conexão de e-mail não encontrada.');
    if (s.type !== 'graph') throw new HttpError(400, 'O visualizador do Teams usa conexões Microsoft 365 (Graph).');
    return s;
  };

  /** Cria o conector (segredos decifrados), roda a leitura (com tempo limite) e fecha. */
  async function live(source, fn) {
    let secrets;
    try {
      secrets = store.openMailSecrets(source);
    } catch (err) {
      throw new HttpError(409, err.message);
    }
    const grantId = source.graph?.account?.grantId;
    const onRefreshToken = grantId ? (token) => store.saveRefreshToken(source.id, grantId, token) : undefined;
    const connector = new TeamsConnector({ ...source, secrets, teams: {} }, { signal: AbortSignal.timeout(60000), endpoints, onRefreshToken });
    try {
      return await fn(connector);
    } catch (err) {
      if (err instanceof HttpError) throw err;
      const status = [400, 403, 404].includes(err?.status) ? err.status : 502;
      throw new HttpError(status, friendlyError(err));
    } finally {
      await connector.close?.();
    }
  }

  // Resolve o usuário pelo e-mail (ou nome de logon).
  router.get('/:id/user', async (req, res) => {
    const source = graphSource(req.params.id);
    const address = String(req.query.address || '').trim();
    if (!address) throw new HttpError(400, 'Informe o e-mail do usuário.');
    const user = await live(source, (c) => c.findUser(address));
    res.json({ user });
  });

  // Conversas do usuário: uma página de chats (com token de continuação) e as equipes/canais.
  router.get('/:id/conversations', async (req, res) => {
    const source = graphSource(req.params.id);
    const userId = String(req.query.userId || '').trim();
    if (!userId) throw new HttpError(400, 'Informe o usuário.');
    const data = await live(source, async (c) => {
      const chats = await c.userChatsPage(userId, null);
      const teams = await c.userTeams(userId);
      return { chats: chats.items, chatsNext: chats.next, teams };
    });
    res.json(data);
  });

  // Mais chats (paginação): continua do token devolvido antes.
  router.get('/:id/chats', async (req, res) => {
    const source = graphSource(req.params.id);
    const userId = String(req.query.userId || '').trim();
    if (!userId) throw new HttpError(400, 'Informe o usuário.');
    const page = await live(source, (c) => c.userChatsPage(userId, nextToken(req)));
    res.json(page);
  });

  // Mensagens de uma conversa (chat ou canal), uma página, as mais recentes primeiro.
  router.get('/:id/messages', async (req, res) => {
    const source = graphSource(req.params.id);
    const kind = req.query.kind;
    const next = nextToken(req);
    const page = await live(source, (c) => {
      if (kind === 'chat') {
        const chatId = String(req.query.chatId || '').trim();
        if (!chatId) throw new HttpError(400, 'Informe o chat.');
        return c.chatMessagesPage(chatId, next);
      }
      if (kind === 'channel') {
        const teamId = String(req.query.teamId || '').trim();
        const channelId = String(req.query.channelId || '').trim();
        if (!teamId || !channelId) throw new HttpError(400, 'Informe a equipe e o canal.');
        return c.channelMessagesPage(teamId, channelId, next);
      }
      throw new HttpError(400, 'Tipo de conversa inválido.');
    });
    res.json(page);
  });

  // Respostas de uma mensagem de canal, uma página.
  router.get('/:id/replies', async (req, res) => {
    const source = graphSource(req.params.id);
    const teamId = String(req.query.teamId || '').trim();
    const channelId = String(req.query.channelId || '').trim();
    const messageId = String(req.query.messageId || '').trim();
    if (!teamId || !channelId || !messageId) throw new HttpError(400, 'Informe a equipe, o canal e a mensagem.');
    const page = await live(source, (c) => c.channelRepliesPage(teamId, channelId, messageId, nextToken(req)));
    res.json(page);
  });

  // Busca geral de um termo nas mensagens recentes dos chats e canais do usuário (varredura limitada).
  router.get('/:id/search', async (req, res) => {
    const source = graphSource(req.params.id);
    const userId = String(req.query.userId || '').trim();
    const q = String(req.query.q || '').trim();
    if (!userId) throw new HttpError(400, 'Informe o usuário.');
    if (q.length < 2) throw new HttpError(400, 'Digite ao menos 2 caracteres para a busca.');
    const data = await live(source, (c) => c.searchMessages(userId, q));
    res.json(data);
  });

  // Conteúdo de uma imagem embutida (hosted content) de uma mensagem. Faz proxy com o token no
  // servidor (o navegador nunca recebe o token) e só serve conteúdo de imagem.
  router.get('/:id/image', async (req, res) => {
    const source = graphSource(req.params.id);
    const messageId = String(req.query.messageId || '').trim();
    const hostedId = String(req.query.hostedId || '').trim();
    if (!messageId || !hostedId) throw new HttpError(400, 'Imagem inválida.');
    const spec = { kind: req.query.kind, messageId, hostedId };
    if (spec.kind === 'chat') {
      spec.chatId = String(req.query.chatId || '').trim();
      if (!spec.chatId) throw new HttpError(400, 'Informe o chat.');
    } else if (spec.kind === 'channel') {
      spec.teamId = String(req.query.teamId || '').trim();
      spec.channelId = String(req.query.channelId || '').trim();
      spec.replyTo = String(req.query.replyTo || '').trim() || null;
      if (!spec.teamId || !spec.channelId) throw new HttpError(400, 'Informe a equipe e o canal.');
    } else {
      throw new HttpError(400, 'Tipo de conversa inválido.');
    }
    const img = await live(source, (c) => c.hostedContent(spec));
    // Só exibe imagem de verdade (imagem rasterizada); SVG e qualquer outro tipo viram download
    // genérico, para não renderizar HTML/script embutido.
    const safe = /^image\/[\w.+-]+$/i.test(img.contentType || '') && !/svg/i.test(img.contentType);
    res.setHeader('Content-Type', safe ? img.contentType : 'application/octet-stream');
    res.setHeader('Cache-Control', 'private, max-age=120');
    res.send(img.data);
  });

  return router;
}

/** Token de continuação (um nextLink do Graph), validado pelo conector antes de usar. */
function nextToken(req) {
  const next = req.query.next;
  return typeof next === 'string' && next ? next : null;
}
