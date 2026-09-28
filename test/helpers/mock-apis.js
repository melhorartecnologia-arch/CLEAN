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

// ---------------- OneDrive e SharePoint (Graph: drives, sites, itens) ----------------

/** Item do simulador no formato do Graph (trail: nomes das pastas acima dele; parentId: pasta). */
function itemJson(item, drive, trail = [], parentId = `${drive.id}-root`) {
  const who = (p) => (p ? { user: { displayName: p.name, email: p.email } } : undefined);
  return {
    id: item.id,
    name: item.name,
    size: item.content ? item.content.length : 0,
    ...(item.children ? { folder: { childCount: item.children.length } } : { file: { mimeType: 'application/octet-stream' } }),
    ...(item.remote ? { remoteItem: { id: 'remoto', parentReference: { driveId: 'outra' } } } : {}),
    createdDateTime: item.created || '2026-09-01T12:00:00Z',
    lastModifiedDateTime: item.modified || '2026-09-10T12:00:00Z',
    createdBy: who(item.createdBy),
    lastModifiedBy: who(item.lastModifiedBy),
    webUrl: `${drive.webUrl}/${encodeURI([...trail, item.name].join('/'))}`,
    eTag: `"{${item.id}},${item.version || 1}${item.renamed ? `,${item.name}` : ''}"`,
    cTag: `"c:{${item.id}},${item.version || 1}"`,
    parentReference: { driveId: drive.id, id: parentId },
  };
}

function findItem(list, id, trail = [], parentId = null) {
  for (const item of list) {
    if (item.id === id) return { item, parent: list, trail, parentId };
    if (item.children) {
      const found = findItem(item.children, id, [...trail, item.name], item.id);
      if (found) return found;
    }
  }
  return null;
}

/** Atende as rotas de arquivos do Graph; retorna false se a rota não for desta parte. */
function drivesApi({ req, res, path, url, base, graph, find, json }) {
  const drives = graph.drives || {};
  const sites = graph.sites || [];
  const page = (list, nextBase) => {
    const skip = Number(url.searchParams.get('$skiptoken') || 0);
    const size = graph.pageSize || 2;
    const next = skip + size < list.length ? { '@odata.nextLink': `${nextBase}${nextBase.includes('?') ? '&' : '?'}$skiptoken=${skip + size}` } : {};
    return { value: list.slice(skip, skip + size), ...next };
  };
  const siteJson = (s) => ({ id: s.id, name: s.name, displayName: s.displayName, webUrl: s.webUrl, isPersonalSite: Boolean(s.personal) });
  let m = /^\/users\/([^/]+)\/drive$/.exec(path);
  if (m) {
    const user = find(m[1]);
    if (!user) return json(res, 404, { error: { code: 'Request_ResourceNotFound', message: 'Usuário não encontrado' } }), true;
    const drive = user.driveId && drives[user.driveId];
    if (!drive) return json(res, 404, { error: { code: 'ResourceNotFound', message: "User's mysite not found." } }), true;
    return json(res, 200, { id: drive.id, name: drive.name, driveType: 'business', webUrl: drive.webUrl, owner: { user: { displayName: user.displayName, email: user.mail } } }), true;
  }
  if (path === '/sites/getAllSites') {
    if (graph.noGetAllSites) return json(res, 400, { error: { code: 'invalidRequest', message: 'Unsupported request' } }), true;
    return json(res, 200, page(sites.filter((s) => !s.parent).map(siteJson), `${base}/graph/v1.0/sites/getAllSites`)), true;
  }
  if (path === '/sites') {
    if (url.searchParams.get('search') !== '*') return json(res, 400, { error: { code: 'invalidRequest', message: 'search obrigatório' } }), true;
    return json(res, 200, page(sites.filter((s) => !s.parent && !s.personal).map(siteJson), `${base}/graph/v1.0/sites?search=*`)), true;
  }
  m = /^\/sites\/([^/:]+)(?::(\/.*))?$/.exec(path);
  if (m && !m[1].includes(',')) {
    const wanted = `https://${m[1]}${m[2] || ''}`.toLowerCase();
    const site = sites.find((s) => s.webUrl.toLowerCase() === wanted);
    return site ? json(res, 200, siteJson(site)) : json(res, 404, { error: { code: 'itemNotFound', message: 'Requested site could not be found' } }), true;
  }
  m = /^\/sites\/([^/]+)\/(sites|drives)$/.exec(path);
  if (m) {
    const site = sites.find((s) => s.id === m[1]);
    if (!site) return json(res, 404, { error: { code: 'itemNotFound', message: 'Site não encontrado' } }), true;
    if (m[2] === 'sites') return json(res, 200, { value: sites.filter((s) => s.parent === site.id).map(siteJson) }), true;
    if (site.denied) return json(res, 403, { error: { code: 'accessDenied', message: 'Access denied' } }), true;
    return json(res, 200, { value: (site.driveIds || []).map((id) => ({ id, name: drives[id].name, driveType: drives[id].driveType || 'documentLibrary', webUrl: drives[id].webUrl })) }), true;
  }
  m = /^\/drives\/([^/]+)\/(?:root|items\/([^/]+))(\/children|\/content|\/permanentDelete)?$/.exec(path);
  if (!m) return false;
  const drive = drives[m[1]];
  if (!drive) return json(res, 404, { error: { code: 'itemNotFound', message: 'Biblioteca não encontrada' } }), true;
  const found = m[2] ? findItem(drive.items, m[2]) : { item: { id: `${drive.id}-root`, children: drive.items }, parent: null, trail: [] };
  if (!found) return json(res, 404, { error: { code: 'itemNotFound', message: 'The resource could not be found.' } }), true;
  const { item, parent } = found;
  const action = m[3] || '';
  if (action === '/children') {
    if (graph.failFolders?.has(item.id)) return json(res, 403, { error: { code: 'accessDenied', message: 'Access denied' } }), true;
    // Falha só na 2ª página em diante da pasta (a 1ª já foi entregue).
    if (graph.failPageOf === (m[2] || 'root') && Number(url.searchParams.get('$skiptoken') || 0) > 0) return json(res, 403, { error: { code: 'accessDenied', message: 'Access denied' } }), true;
    const nextBase = `${base}/graph/v1.0/drives/${drive.id}/${m[2] ? `items/${item.id}` : 'root'}/children`;
    const trail = m[2] ? [...found.trail, item.name] : [];
    return json(res, 200, page(item.children.map((c) => itemJson(c, drive, trail, item.id)), nextBase)), true;
  }
  if (action === '/content') {
    res.writeHead(302, { Location: `${base}/download/${drive.id}/${item.id}` });
    res.end();
    return true;
  }
  const record = (how) => {
    graph.driveDeleted = [...(graph.driveDeleted || []), { drive: drive.id, id: item.id, how, ifMatch: req.headers['if-match'] || '' }];
    parent.splice(parent.indexOf(item), 1);
  };
  const current = itemJson(item, drive, found.trail, found.parentId || `${drive.id}-root`);
  const tagMismatch = () => {
    const tag = req.headers['if-match'];
    return tag && tag !== current.eTag && tag !== current.cTag;
  };
  if (action === '/permanentDelete' && req.method === 'POST') {
    if (graph.readOnlyDrives?.has(drive.id)) return json(res, 403, { error: { code: 'accessDenied', message: 'Access denied' } }), true;
    if (graph.lockedItems?.has(item.id)) return json(res, 423, { error: { code: 'resourceLocked', message: 'The resource you are attempting to access is locked' } }), true;
    if (tagMismatch()) return json(res, 412, { error: { code: 'resourceModified', message: 'ETag does not match current item\'s value' } }), true;
    record('permanent');
    res.writeHead(204);
    res.end();
    return true;
  }
  if (req.method === 'DELETE') {
    if (graph.readOnlyDrives?.has(drive.id)) return json(res, 403, { error: { code: 'accessDenied', message: 'Access denied' } }), true;
    if (graph.lockedItems?.has(item.id)) return json(res, 423, { error: { code: 'resourceLocked', message: 'The resource you are attempting to access is locked' } }), true;
    if (tagMismatch()) return json(res, 412, { error: { code: 'resourceModified', message: 'ETag does not match current item\'s value' } }), true;
    record('trash');
    (graph.recycle ||= []).push({ drive: drive.id, item });
    res.writeHead(204);
    res.end();
    return true;
  }
  if (req.method === 'GET' && !action) return json(res, 200, current), true;
  return false;
}

/**
 * graph: { tenant, clientId, secret, users: [{ id, mail, displayName, noMailbox?, sharedWith?: [e-mails com
 *          Acesso Total à caixa], folders: [{ id, displayName, parent?, wellKnown? }], messages: { [folderId]:
 *          [{ id, raw, received }] } }], throttleOnce?: Set<messageId>,
 *          deleteError?: { status, code } (resposta das exclusões), flakyDelete? (a 1ª exclusão de cada
 *          mensagem é feita, mas a resposta é um erro 503, como uma resposta perdida),
 *          certificatePem? (certificado do aplicativo: a asserção do cliente é conferida com ele),
 *          delegated?: { signInAs (e-mail da conta que entra pelo código), pendingPolls (consultas
 *          "authorization_pending" antes de concluir), publicClientDisabled?, declined?, revoked?,
 *          refreshTokens: { [token]: { user, scopes } } (tokens de atualização válidos; cada renovação
 *          devolve outro e o antigo continua valendo, como no Entra ID), issued: [tokens emitidos] } }
 * google: { publicKey, admin, users: [{ mail, name, disabled?, labels: [{ id, name, type }],
 *          messages: [{ id, raw, labelIds, internalDate }] }], flakyDelete? }
 */
/** Cabeçalho de uma mensagem MIME do simulador (primeira ocorrência). */
function mimeHeader(raw, name) {
  const head = raw.toString('utf8').split(/\r?\n\r?\n/)[0];
  return (new RegExp(`^${name}:[ \t]*(.*)$`, 'im').exec(head)?.[1] || '').trim();
}

const b64json = (value) => Buffer.from(JSON.stringify(value)).toString('base64url');
const unb64json = (text) => JSON.parse(Buffer.from(String(text), 'base64url').toString('utf8'));
const IMAP_SCOPE = 'https://outlook.office.com/IMAP.AccessAsUser.All';
const OIDC_SCOPES = new Set(['openid', 'profile', 'offline_access', 'email']);

/** Token de acesso delegado do simulador: "d.<json>" para o Graph e "i.<json>" para o IMAP. */
export const delegatedToken = (user, scopes) => `${scopes.includes(IMAP_SCOPE) ? 'i' : 'd'}.${b64json({ user, scopes })}`;
/** Conta do token de acesso do IMAP (ou null): usado pelo servidor IMAP de teste. */
export function imapTokenUser(token) {
  if (token === 'imap-app') return '*';
  if (!String(token).startsWith('i.')) return null;
  try {
    return unb64json(token.slice(2)).user;
  } catch {
    return null;
  }
}

/** Confere a asserção do cliente (certificado): assinatura PS256, impressão digital e prazos. */
function checkAssertion(graph, form, audience) {
  if (form.get('client_assertion_type') !== 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer') return 'AADSTS50027: Invalid client_assertion_type.';
  const [h, p, sig] = String(form.get('client_assertion')).split('.');
  let header;
  let claims;
  try {
    header = unb64json(h);
    claims = unb64json(p);
  } catch {
    return 'AADSTS50027: JWT token is invalid or malformed.';
  }
  if (!graph.certificatePem) return 'AADSTS700027: Client assertion failed signature validation.';
  const cert = new crypto.X509Certificate(graph.certificatePem);
  const thumbprint = Buffer.from(cert.fingerprint256.replace(/:/g, ''), 'hex').toString('base64url');
  if (header.alg !== 'PS256' || header['x5t#S256'] !== thumbprint) return 'AADSTS700027: Client assertion failed signature validation.';
  const valid = crypto.verify('sha256', Buffer.from(`${h}.${p}`), { key: cert.publicKey, padding: crypto.constants.RSA_PKCS1_PSS_PADDING, saltLength: crypto.constants.RSA_PSS_SALTLEN_DIGEST }, Buffer.from(sig, 'base64url'));
  if (!valid) return 'AADSTS700027: Client assertion failed signature validation.';
  const now = Date.now() / 1000;
  if (claims.aud !== audience || claims.iss !== graph.clientId || claims.sub !== graph.clientId || !claims.jti) return 'AADSTS700021: Client assertion audience or issuer is invalid.';
  if (!(claims.exp > now) || !(claims.nbf <= now + 300) || claims.exp - claims.nbf > 600) return 'AADSTS700024: Client assertion is not within its valid time range.';
  return null;
}

export function startMockApis({ graph = null, google = null } = {}) {
  const calls = [];
  const throttled = new Set();
  const flaked = new Set();
  const lostResponse = (res, id) => {
    flaked.add(id);
    return json(res, 503, { error: { code: 'ServiceUnavailable', message: 'Serviço indisponível' } });
  };
  const server = http.createServer(async (req, res) => {
    const url = new URL(req.url, 'http://x');
    const base = `http://127.0.0.1:${server.address().port}`;
    calls.push(`${req.method} ${url.pathname}`);
    const body = req.method === 'POST' ? await readBody(req) : '';
    try {
      // ---------------- Microsoft Entra ID / Graph ----------------
      const login = /^\/graph-login\/([^/]+)\/oauth2\/v2\.0\/(token|devicecode)$/.exec(url.pathname);
      if (login) {
        const form = new URLSearchParams(body);
        const tenant = decodeURIComponent(login[1]);
        const d = graph.delegated || {};
        const grant = form.get('grant_type');
        const delegatedGrant = login[2] === 'devicecode' || grant !== 'client_credentials';
        if (tenant !== graph.tenant && !(delegatedGrant && ['organizations', 'common', 'consumers'].includes(tenant))) return json(res, 400, { error: 'invalid_request', error_description: 'AADSTS90002: Tenant not found.' });
        if (form.get('client_id') !== graph.clientId) return json(res, 400, { error: 'unauthorized_client', error_description: 'AADSTS700016: Application not found.' });
        (graph.tokenRequests ||= []).push({ endpoint: login[2], grant, scope: form.get('scope') || '' });
        if (login[2] === 'devicecode') {
          const code = `device-${crypto.randomUUID()}`;
          (d.devices ||= {})[code] = { user: d.signInAs, pendingPolls: d.pendingPolls ?? 1, scopes: String(form.get('scope') || '').split(' ') };
          return json(res, 200, { device_code: code, user_code: 'ABCD-EFGH', verification_uri: 'https://microsoft.com/devicelogin', expires_in: 900, interval: 1, message: 'To sign in, use a web browser...' });
        }
        // Conta conectada: resposta com o token de acesso ao recurso pedido e um novo token de atualização.
        const delegatedTokens = (user, scopes) => {
          const refresh = `rt-${crypto.randomUUID()}`;
          (d.refreshTokens ||= {})[refresh] = { user, scopes };
          (d.issued ||= []).push(refresh);
          const granted = scopes.filter((x) => !OIDC_SCOPES.has(x));
          const u = graph.users.find((x) => x.mail === user) || {};
          const idToken = `${b64json({ alg: 'none' })}.${b64json({ oid: u.id || user, tid: 'tid-contoso', preferred_username: u.upn || user, name: u.displayName || '' })}.`;
          return { token_type: 'Bearer', scope: granted.join(' '), expires_in: 3600, access_token: delegatedToken(user, granted), refresh_token: refresh, id_token: idToken };
        };
        if (grant === 'urn:ietf:params:oauth:grant-type:device_code') {
          const device = d.devices?.[form.get('device_code')];
          if (!device) return json(res, 400, { error: 'bad_verification_code', error_description: 'AADSTS70000: Unknown device code.' });
          if (d.publicClientDisabled) return json(res, 401, { error: 'invalid_client', error_description: "AADSTS7000218: The request body must contain the following parameter: 'client_assertion' or 'client_secret'." });
          if (d.declined) return json(res, 400, { error: 'authorization_declined', error_description: 'AADSTS70019: The user declined.' });
          if (device.pendingPolls > 0) {
            device.pendingPolls--;
            return json(res, 400, { error: 'authorization_pending', error_description: 'AADSTS70016: OAuth 2.0 device flow error. Authorization is pending. Continue polling.' });
          }
          delete d.devices[form.get('device_code')];
          return json(res, 200, delegatedTokens(device.user, device.scopes));
        }
        if (grant === 'refresh_token') {
          const entry = d.refreshTokens?.[form.get('refresh_token')];
          if (d.revoked) return json(res, 400, { error: 'invalid_grant', error_description: 'AADSTS50173: The provided grant has expired due to it being revoked.' });
          if (!entry) return json(res, 400, { error: 'invalid_grant', error_description: 'AADSTS700082: The refresh token has expired due to inactivity.' });
          const asked = String(form.get('scope') || '').split(' ').filter((x) => x && !OIDC_SCOPES.has(x));
          const missing = asked.filter((x) => !entry.scopes.includes(x));
          if (missing.length) return json(res, 400, { error: 'invalid_grant', error_description: `AADSTS65001: The user or administrator has not consented to use the application (${missing.join(' ')}).` });
          return json(res, 200, delegatedTokens(entry.user, entry.scopes));
        }
        // Aplicativo: segredo do cliente ou certificado (asserção assinada).
        const imap = form.get('scope') === 'https://outlook.office365.com/.default';
        if (form.get('client_assertion')) {
          const problem = checkAssertion(graph, form, `${base}/graph-login/${login[1]}/oauth2/v2.0/token`);
          if (problem) return json(res, 401, { error: 'invalid_client', error_description: problem });
          graph.assertions = (graph.assertions || 0) + 1;
        } else if (form.get('client_secret') !== graph.secret) {
          return json(res, 401, { error: 'invalid_client', error_description: 'AADSTS7000215: Invalid client secret provided.' });
        }
        return json(res, 200, { access_token: imap ? 'imap-app' : 'graph-token', expires_in: 3600, token_type: 'Bearer' });
      }
      // Endereço de download pré-autenticado (como o do SharePoint): sem o cabeçalho de autorização.
      const download = /^\/download\/([^/]+)\/([^/]+)$/.exec(url.pathname);
      if (download && graph?.drives) {
        const found = findItem(graph.drives[download[1]]?.items || [], download[2]);
        if (!found) return json(res, 404, { error: { code: 'itemNotFound', message: 'Não encontrado' } });
        graph.downloads = [...(graph.downloads || []), { id: download[2], auth: req.headers.authorization || '' }];
        res.writeHead(200, { 'Content-Type': 'application/octet-stream', 'Content-Length': found.item.content.length });
        if (graph.slowDownloads?.has(download[2])) {
          // Download lento: metade agora, o resto depois; registra se o cliente desistiu antes.
          const half = Math.floor(found.item.content.length / 2);
          res.write(found.item.content.subarray(0, half));
          let finished = false;
          res.on('close', () => {
            if (!finished) graph.abortedDownloads = [...(graph.abortedDownloads || []), download[2]];
          });
          setTimeout(() => {
            finished = true;
            if (!res.destroyed) res.end(found.item.content.subarray(half));
          }, graph.slowMs || 3000);
          return undefined;
        }
        return res.end(found.item.content);
      }
      if (url.pathname.startsWith('/graph/v1.0/')) {
        // Aplicativo ("graph-token") ou conta conectada ("d.<json>": conta e permissões delegadas).
        const bearer = String(req.headers.authorization || '');
        let principal = null;
        if (bearer === 'Bearer graph-token') principal = { app: true };
        else if (bearer.startsWith('Bearer d.')) principal = unb64json(bearer.slice('Bearer d.'.length));
        if (!principal) return json(res, 401, { error: { code: 'InvalidAuthenticationToken', message: 'Token inválido' } });
        let path = decodeURIComponent(url.pathname.slice('/graph/v1.0'.length));
        const users = graph.users;
        const find = (key) =>
          users.find((u) => u.id === key || (u.mail.toLowerCase() === String(key).toLowerCase() && u.upnIsMail !== false) || (u.upn && u.upn.toLowerCase() === String(key).toLowerCase()));
        if (!principal.app) {
          (graph.delegatedCalls ||= []).push(`${req.method} ${path}`);
          const me = users.find((u) => u.mail === principal.user);
          const has = (scope) => principal.scopes.some((x) => x.endsWith(`/${scope}`) || x === scope);
          if (path === '/me') return json(res, 200, { id: me.id, displayName: me.displayName, mail: me.mail, userPrincipalName: me.upn || me.mail });
          if (path.startsWith('/me/')) path = `/users/${me.id}${path.slice(3)}`;
          const target = /^\/users\/([^/]+)\/(.+)$/.exec(path);
          if (!target) return json(res, 403, { error: { code: 'Authorization_RequestDenied', message: 'Insufficient privileges to complete the operation.' } });
          const owner = find(target[1]);
          const own = owner === me;
          const write = /\/(permanentDelete|move)$/.test(path) && req.method === 'POST';
          const needed = own ? (write ? ['Mail.ReadWrite'] : ['Mail.Read', 'Mail.ReadWrite']) : write ? ['Mail.ReadWrite.Shared'] : ['Mail.Read.Shared', 'Mail.ReadWrite.Shared'];
          const allowed = owner && (own || (owner.sharedWith || []).includes(principal.user)) && needed.some(has);
          if (!allowed) return json(res, 403, { error: { code: 'ErrorAccessDenied', message: 'Access is denied. Check credentials and try again.' } });
        }
        if (principal.app && path.startsWith('/me')) return json(res, 400, { error: { code: 'BadRequest', message: '/me request is only valid with delegated authentication flow.' } });
        if (drivesApi({ req, res, path, url, base, graph, find, json })) return;
        if (path === '/users') {
          const filter = url.searchParams.get('$filter');
          if (filter) {
            const mail = /mail eq '([^']+)'/.exec(filter)?.[1];
            return json(res, 200, { value: users.filter((u) => u.mail === mail).map((u) => ({ id: u.id, mail: u.mail, displayName: u.displayName })) });
          }
          const skip = Number(url.searchParams.get('$skiptoken') || 0);
          const page = users.slice(skip, skip + 2).map((u) => ({ id: u.id, mail: u.mail, displayName: u.displayName, userPrincipalName: u.upn || u.mail }));
          const next = skip + 2 < users.length ? { '@odata.nextLink': `${base}/graph/v1.0/users?$skiptoken=${skip + 2}` } : {};
          return json(res, 200, { value: page, ...next });
        }
        let m = /^\/users\/([^/]+)$/.exec(path);
        if (m) {
          const u = users.find((x) => x.id === m[1] || (x.upnIsMail !== false && x.mail === m[1]) || (x.upn && x.upn.toLowerCase() === m[1].toLowerCase()));
          return u
            ? json(res, 200, { id: u.id, mail: u.mail, displayName: u.displayName, userPrincipalName: u.upn || u.mail })
            : json(res, 404, { error: { code: 'Request_ResourceNotFound', message: 'Não encontrado' } });
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
          const ge = filter && /receivedDateTime ge (\S+)/.exec(filter);
          const lt = filter && /receivedDateTime lt (\S+)/.exec(filter);
          const since = ge ? new Date(ge[1]) : null;
          const before = lt ? new Date(lt[1]) : null;
          const all = list.filter((x) => (!since || new Date(x.received) >= since) && (!before || new Date(x.received) < before));
          const headers = (url.searchParams.get('$select') || '').includes('subject');
          const skip = Number(url.searchParams.get('$skip') || 0);
          const page = all.slice(skip, skip + 2);
          const next = skip + 2 < all.length ? { '@odata.nextLink': `${base}/graph/v1.0/users/${user.id}/mailFolders/${m[1]}/messages?$skip=${skip + 2}${filter ? `&$filter=${encodeURIComponent(filter)}` : ''}` } : {};
          return json(res, 200, {
            value: page.map((x) => ({
              id: x.id,
              receivedDateTime: x.received,
              webLink: `https://outlook.office365.com/owa/?ItemID=${x.id}`,
              singleValueExtendedProperties: [{ id: 'Integer 0x0e08', value: String(x.raw.length) }],
              ...(headers
                ? {
                    subject: mimeHeader(x.raw, 'Subject'),
                    internetMessageId: mimeHeader(x.raw, 'Message-ID') || null,
                    from: { emailAddress: { name: /^(.*?)\s*</.exec(mimeHeader(x.raw, 'From'))?.[1]?.replace(/"/g, '') || '', address: /<([^>]+)>/.exec(mimeHeader(x.raw, 'From'))?.[1] || mimeHeader(x.raw, 'From') } },
                  }
                : {}),
            })),
            ...next,
          });
        }
        m = /^messages\/([^/]+)\/(permanentDelete|move)$/.exec(rest);
        if (m && req.method === 'POST') {
          if (user.readOnly) return json(res, 403, { error: { code: 'ErrorAccessDenied', message: 'Access is denied. Check credentials and try again.' } });
          if (graph.deleteError) return json(res, graph.deleteError.status, { error: { code: graph.deleteError.code, message: graph.deleteError.message || 'Erro simulado' } });
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
          if (graph.flakyDelete && !flaked.has(m[1])) return lostResponse(res, m[1]);
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
          const before = Number(/before:(\d+)/.exec(q)?.[1] || Infinity) * 1000;
          const smaller = Number(/smaller:(\d+)/.exec(q)?.[1] || Infinity);
          const larger = Number(/larger:(\d+)/.exec(q)?.[1] || -1);
          const list = user.messages.filter((m) => {
            const labels = m.labelIds || [];
            if (!spamTrash && (labels.includes('SPAM') || labels.includes('TRASH'))) return false;
            if (q.includes('-in:spam') && labels.includes('SPAM')) return false;
            if (q.includes('-in:trash') && labels.includes('TRASH')) return false;
            const size = m.size ?? m.raw.length;
            if (!(size < smaller) || !(size > larger)) return false;
            return Number(m.internalDate) >= after && Number(m.internalDate) < before;
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
          if (google.flakyDelete && !flaked.has(act[1])) return lostResponse(res, act[1]);
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
          if (url.searchParams.get('format') === 'metadata') {
            const wanted = url.searchParams.getAll('metadataHeaders');
            return json(res, 200, { ...common, payload: { headers: wanted.map((name) => ({ name, value: mimeHeader(m.raw, name) })).filter((h) => h.value) } });
          }
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
