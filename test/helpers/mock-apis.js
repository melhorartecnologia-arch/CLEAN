// Servidor HTTP que imita as partes do Microsoft Graph e das APIs do Google usadas pelo CLEAN.
import http from 'node:http';
import crypto from 'node:crypto';

const json = (res, status, body, headers = {}) => {
  res.writeHead(status, { 'Content-Type': 'application/json', ...headers });
  res.end(JSON.stringify(body));
};

function readBody(req) {
  return new Promise((resolve) => {
    const chunks = [];
    req.on('data', (c) => chunks.push(c));
    req.on('end', () => resolve(Buffer.concat(chunks).toString('utf8')));
  });
}

/**
 * graph: { tenant, clientId, secret, users: [{ id, mail, displayName, noMailbox?, folders: [{ id, displayName,
 *          parent?, wellKnown? }], messages: { [folderId]: [{ id, raw, received }] } }], throttleOnce?: Set<messageId> }
 * google: { publicKey, admin, users: [{ mail, name, disabled?, labels: [{ id, name, type }],
 *          messages: [{ id, raw, labelIds, internalDate }] }] }
 */
export function startMockApis({ graph = null, google = null } = {}) {
  const calls = [];
  const throttled = new Set();
  const server = http.createServer(async (req, res) => {
    const url = new URL(req.url, 'http://x');
    const base = `http://127.0.0.1:${server.address().port}`;
    calls.push(`${req.method} ${url.pathname}`);
    const body = req.method === 'POST' ? await readBody(req) : '';
    try {
      // ---------------- Microsoft Entra ID / Graph ----------------
      const login = /^\/graph-login\/([^/]+)\/oauth2\/v2\.0\/token$/.exec(url.pathname);
      if (login) {
        const form = new URLSearchParams(body);
        if (decodeURIComponent(login[1]) !== graph.tenant) return json(res, 400, { error: 'invalid_request', error_description: 'AADSTS90002: Tenant not found.' });
        if (form.get('client_id') !== graph.clientId) return json(res, 400, { error: 'unauthorized_client', error_description: 'AADSTS700016: Application not found.' });
        if (form.get('client_secret') !== graph.secret) return json(res, 401, { error: 'invalid_client', error_description: 'AADSTS7000215: Invalid client secret provided.' });
        return json(res, 200, { access_token: 'graph-token', expires_in: 3600, token_type: 'Bearer' });
      }
      if (url.pathname.startsWith('/graph/v1.0/')) {
        if (req.headers.authorization !== 'Bearer graph-token') return json(res, 401, { error: { code: 'InvalidAuthenticationToken', message: 'Token inválido' } });
        const path = decodeURIComponent(url.pathname.slice('/graph/v1.0'.length));
        const users = graph.users;
        const find = (key) => users.find((u) => u.id === key || u.mail.toLowerCase() === String(key).toLowerCase() && u.upnIsMail !== false);
        if (path === '/users') {
          const filter = url.searchParams.get('$filter');
          if (filter) {
            const mail = /mail eq '([^']+)'/.exec(filter)?.[1];
            return json(res, 200, { value: users.filter((u) => u.mail === mail).map((u) => ({ id: u.id, mail: u.mail, displayName: u.displayName })) });
          }
          const skip = Number(url.searchParams.get('$skiptoken') || 0);
          const page = users.slice(skip, skip + 2).map((u) => ({ id: u.id, mail: u.mail, displayName: u.displayName, userPrincipalName: u.mail }));
          const next = skip + 2 < users.length ? { '@odata.nextLink': `${base}/graph/v1.0/users?$skiptoken=${skip + 2}` } : {};
          return json(res, 200, { value: page, ...next });
        }
        let m = /^\/users\/([^/]+)$/.exec(path);
        if (m) {
          const u = users.find((x) => x.id === m[1] || (x.upnIsMail !== false && x.mail === m[1]));
          return u ? json(res, 200, { id: u.id, mail: u.mail, displayName: u.displayName }) : json(res, 404, { error: { code: 'Request_ResourceNotFound', message: 'Não encontrado' } });
        }
        m = /^\/users\/([^/]+)\/(.*)$/.exec(path);
        const user = m && find(m[1]);
        if (!user) return json(res, 404, { error: { code: 'Request_ResourceNotFound', message: 'Usuário não encontrado' } });
        const rest = m[2];
        if (user.noMailbox) return json(res, 404, { error: { code: 'MailboxNotEnabledForRESTAPI', message: 'The mailbox is either inactive, soft-deleted, or is hosted on-premise.' } });
        if (user.accessDenied) return json(res, 403, { error: { code: 'ErrorAccessDenied', message: 'Access is denied. Check credentials and try again.' } });
        const folderJson = (f) => ({
          id: f.id,
          displayName: f.displayName,
          childFolderCount: user.folders.filter((c) => c.parent === f.id).length,
          totalItemCount: (user.messages[f.id] || []).length,
        });
        m = /^mailFolders\/(deleteditems|junkemail|syncissues|inbox)$/.exec(rest);
        if (m) {
          const f = user.folders.find((x) => x.wellKnown === m[1]);
          return f ? json(res, 200, folderJson(f)) : json(res, 404, { error: { code: 'ErrorItemNotFound', message: 'Pasta não encontrada' } });
        }
        if (rest === 'mailFolders') return json(res, 200, { value: user.folders.filter((f) => !f.parent).map(folderJson) });
        m = /^mailFolders\/([^/]+)\/childFolders$/.exec(rest);
        if (m) return json(res, 200, { value: user.folders.filter((f) => f.parent === m[1]).map(folderJson) });
        m = /^mailFolders\/([^/]+)\/messages$/.exec(rest);
        if (m) {
          if (graph.failFolders?.has(m[1])) return json(res, 404, { error: { code: 'ErrorItemNotFound', message: 'The specified object was not found in the store.' } });
          const list = user.messages[m[1]] || [];
          const filter = url.searchParams.get('$filter');
          const since = filter ? new Date(/receivedDateTime ge (\S+)/.exec(filter)[1]) : null;
          const all = list.filter((x) => !since || new Date(x.received) >= since);
          const skip = Number(url.searchParams.get('$skip') || 0);
          const page = all.slice(skip, skip + 2);
          const next = skip + 2 < all.length ? { '@odata.nextLink': `${base}/graph/v1.0/users/${user.id}/mailFolders/${m[1]}/messages?$skip=${skip + 2}${filter ? `&$filter=${encodeURIComponent(filter)}` : ''}` } : {};
          return json(res, 200, {
            value: page.map((x) => ({
              id: x.id,
              receivedDateTime: x.received,
              webLink: `https://outlook.office365.com/owa/?ItemID=${x.id}`,
              singleValueExtendedProperties: [{ id: 'Integer 0x0e08', value: String(x.raw.length) }],
            })),
            ...next,
          });
        }
        m = /^messages\/([^/]+)\/(permanentDelete|move)$/.exec(rest);
        if (m && req.method === 'POST') {
          if (user.readOnly) return json(res, 403, { error: { code: 'ErrorAccessDenied', message: 'Access is denied. Check credentials and try again.' } });
          const folderId = Object.keys(user.messages).find((f) => user.messages[f].some((x) => x.id === m[1]));
          if (!folderId) return json(res, 404, { error: { code: 'ErrorItemNotFound', message: 'The specified object was not found in the store.' } });
          const index = user.messages[folderId].findIndex((x) => x.id === m[1]);
          const [msg] = user.messages[folderId].splice(index, 1);
          graph.deleted = [...(graph.deleted || []), { user: user.id, id: m[1], how: m[2], prefer: req.headers.prefer || '' }];
          if (m[2] === 'move') {
            const target = user.folders.find((f) => f.wellKnown === JSON.parse(body).destinationId)?.id;
            (user.messages[target] ||= []).push(msg);
            return json(res, 201, { id: msg.id });
          }
          res.writeHead(204);
          return res.end();
        }
        m = /^messages\/([^/]+)\/\$value$/.exec(rest);
        if (m) {
          const msg = Object.values(user.messages).flat().find((x) => x.id === m[1]);
          if (!msg) return json(res, 404, { error: { code: 'ErrorItemNotFound', message: 'Mensagem não encontrada' } });
          if (graph.throttleOnce?.has(msg.id) && !throttled.has(msg.id)) {
            throttled.add(msg.id);
            return json(res, 429, { error: { code: 'ApplicationThrottled', message: 'Muitas requisições' } }, { 'Retry-After': '1' });
          }
          res.writeHead(200, { 'Content-Type': 'text/plain', 'Content-Length': msg.raw.length });
          return res.end(msg.raw);
        }
        return json(res, 404, { error: { code: 'NotImplemented', message: `Rota do simulador não implementada: ${path}` } });
      }

      // ---------------- Google ----------------
      if (url.pathname === '/google/token') {
        const form = new URLSearchParams(body);
        const [head, payload, signature] = String(form.get('assertion')).split('.');
        const valid = crypto.verify('RSA-SHA256', Buffer.from(`${head}.${payload}`), google.publicKey, Buffer.from(signature, 'base64url'));
        if (!valid) return json(res, 400, { error: 'invalid_grant', error_description: 'Invalid JWT Signature.' });
        const claims = JSON.parse(Buffer.from(payload, 'base64url').toString('utf8'));
        const known = claims.sub === google.admin || google.users.some((u) => u.mail === claims.sub);
        if (!known) return json(res, 400, { error: 'invalid_grant', error_description: 'Invalid email or User ID' });
        if (claims.scope.includes('directory') && claims.sub !== google.admin) return json(res, 401, { error: 'unauthorized_client', error_description: 'Client is unauthorized' });
        const kind = claims.scope.includes('directory') ? 'dir' : claims.scope === 'https://mail.google.com/' ? 'full' : claims.scope.endsWith('gmail.modify') ? 'modify' : 'gmail';
        if (google.deniedScopes?.has(claims.scope)) return json(res, 401, { error: 'unauthorized_client', error_description: 'Client is unauthorized to retrieve access tokens using this method, or client not authorized for any of the scopes requested.' });
        return json(res, 200, { access_token: `g|${claims.sub}|${kind}`, expires_in: 3600 });
      }
      if (url.pathname === '/directory/v1/users') {
        if (req.headers.authorization !== `Bearer g|${google.admin}|dir`) return json(res, 403, { error: { code: 403, message: 'Not Authorized to access this resource/api' } });
        return json(res, 200, { users: google.users.map((u) => ({ primaryEmail: u.mail, name: { fullName: u.name } })) });
      }
      const g = /^\/gmail\/v1\/users\/([^/]+)\/(.*)$/.exec(url.pathname);
      if (g) {
        const mail = decodeURIComponent(g[1]);
        const tokenKind = String(req.headers.authorization || '').startsWith(`Bearer g|${mail}|`) ? req.headers.authorization.split('|')[2] : null;
        if (!tokenKind || tokenKind === 'dir') return json(res, 401, { error: { code: 401, message: 'Invalid Credentials', status: 'UNAUTHENTICATED' } });
        const user = google.users.find((u) => u.mail === mail);
        if (user.disabled) return json(res, 400, { error: { code: 400, message: 'Mail service not enabled', status: 'FAILED_PRECONDITION' } });
        if (g[2] === 'labels') return json(res, 200, { labels: user.labels });
        if (g[2] === 'profile') return json(res, 200, { emailAddress: mail, messagesTotal: user.messages.length });
        if (g[2] === 'messages') {
          const spamTrash = url.searchParams.get('includeSpamTrash') === 'true';
          const q = url.searchParams.get('q') || '';
          const after = Number(/after:(\d+)/.exec(q)?.[1] || 0) * 1000;
          const smaller = Number(/smaller:(\d+)/.exec(q)?.[1] || Infinity);
          const larger = Number(/larger:(\d+)/.exec(q)?.[1] || -1);
          const list = user.messages.filter((m) => {
            const labels = m.labelIds || [];
            if (!spamTrash && (labels.includes('SPAM') || labels.includes('TRASH'))) return false;
            if (q.includes('-in:spam') && labels.includes('SPAM')) return false;
            if (q.includes('-in:trash') && labels.includes('TRASH')) return false;
            const size = m.size ?? m.raw.length;
            if (!(size < smaller) || !(size > larger)) return false;
            return Number(m.internalDate) >= after;
          });
          const start = Number(url.searchParams.get('pageToken') || 0);
          const page = list.slice(start, start + 2);
          return json(res, 200, { messages: page.map((m) => ({ id: m.id, threadId: m.id })), ...(start + 2 < list.length ? { nextPageToken: String(start + 2) } : {}) });
        }
        const act = /^messages\/([^/]+)(\/trash)?$/.exec(g[2]);
        if (act && (req.method === 'DELETE' || act[2])) {
          const needed = act[2] ? ['modify', 'full'] : ['full'];
          if (!needed.includes(tokenKind)) return json(res, 403, { error: { code: 403, message: 'Request had insufficient authentication scopes.', status: 'PERMISSION_DENIED' } });
          const index = user.messages.findIndex((x) => x.id === act[1]);
          if (index === -1) return json(res, 404, { error: { code: 404, message: 'Requested entity was not found.', status: 'NOT_FOUND' } });
          google.deleted = [...(google.deleted || []), { user: mail, id: act[1], how: act[2] ? 'trash' : 'delete' }];
          if (act[2]) {
            user.messages[index].labelIds = ['TRASH'];
            return json(res, 200, { id: act[1], labelIds: ['TRASH'] });
          }
          user.messages.splice(index, 1);
          res.writeHead(204);
          return res.end();
        }
        const msg = /^messages\/([^/]+)$/.exec(g[2]);
        if (msg) {
          const m = user.messages.find((x) => x.id === msg[1]);
          if (google.rateLimitOnce?.has(m.id) && !throttled.has(m.id)) {
            throttled.add(m.id);
            return json(res, 403, { error: { code: 403, message: 'User Rate Limit Exceeded', errors: [{ reason: 'userRateLimitExceeded', domain: 'usageLimits' }], status: 'PERMISSION_DENIED' } });
          }
          const common = { id: m.id, threadId: m.id, labelIds: m.labelIds, sizeEstimate: m.size ?? m.raw.length, internalDate: String(m.internalDate) };
          if (url.searchParams.get('format') === 'full') return json(res, 200, { ...common, payload: m.payload });
          return json(res, 200, { ...common, raw: m.raw.toString('base64url') });
        }
      }
      return json(res, 404, { error: { code: 'NotFound', message: `Rota do simulador não implementada: ${url.pathname}` } });
    } catch (err) {
      return json(res, 500, { error: { code: 'MockError', message: err.stack } });
    }
  });
  return new Promise((resolve) => {
    server.listen(0, '127.0.0.1', () => {
      const base = `http://127.0.0.1:${server.address().port}`;
      resolve({
        server,
        calls,
        endpoints: {
          graphLogin: `${base}/graph-login`,
          graph: `${base}/graph/v1.0`,
          googleToken: `${base}/google/token`,
          gmail: `${base}/gmail/v1`,
          googleDirectory: `${base}/directory/v1`,
        },
        close: () => new Promise((r) => server.close(r)),
      });
    });
  });
}
