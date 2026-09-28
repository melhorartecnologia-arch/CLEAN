// Rotas /api/mail-sources: conexões com caixas de e-mail (Microsoft 365, Google Workspace, IMAP).
// Os segredos (segredo do cliente, chave do certificado, token da conta Microsoft conectada, chave da
// conta de serviço, senhas) são gravados cifrados e nunca voltam para o navegador: a API informa
// apenas se cada um está salvo. A entrada da conta Microsoft (código de dispositivo) e os
// certificados gerados ficam na memória do servidor até a conexão ser salva.
import crypto from 'node:crypto';
import { Router } from 'express';
import { HttpError, bad, text, lines, email, emailList, microsoftApp, microsoftCredentials, linkedCredentials, assertUnused } from './validate.js';
import { createConnector, MAIL_TYPES } from '../mail/connectors.js';
import { imapOAuthProblem } from '../mail/imap.js';
import { request } from '../mail/http.js';
import { GRAPH_ENDPOINTS } from '../cloud/graph-client.js';
import { startDeviceCode, pollDeviceCode, delegatedScopes, accountFromTokens, grantedScopes, canWriteMail } from '../cloud/microsoft-auth.js';
import { createCertificate, importCertificate } from '../cloud/certificate.js';
import { PendingCredentials } from '../cloud/pending-credentials.js';
import { friendlyError } from '../scan/errors.js';
import { checkDeleteSchedules } from '../schedule/scheduler.js';

/** Mensagem de um erro ao falar com a Microsoft (tempo esgotado em português). */
const microsoftError = (err) =>
  err?.name === 'TimeoutError' || err?.name === 'AbortError'
    ? 'Tempo esgotado ao falar com a Microsoft (login.microsoftonline.com): confira a conexão deste servidor com a internet e tente de novo.'
    : friendlyError(err);

const HOST_RE = /^(?=.{1,253}$)[a-z0-9]([a-z0-9-]*[a-z0-9])?(\.[a-z0-9]([a-z0-9-]*[a-z0-9])?)*$|^\[?[0-9a-f:.]+\]?$/i;
const MAX_MAILBOXES = 5000;
// Entrada concluída ou certificado gerado: disponíveis para salvar a conexão por este tempo.
const SIGN_IN_TTL_MS = 2 * 3600 * 1000;
const CERTIFICATE_TTL_MS = 24 * 3600 * 1000;

const key = (address) => String(address || '').trim().toLowerCase();

/** Lista de endereços de caixas (texto com um por linha ou lista), sem repetições. */
const addressList = (value, field) => emailList(value, field, { max: MAX_MAILBOXES, noun: 'caixas' });

/** Mesmo servidor, porta e segurança, sem relaxar a verificação do certificado. */
export function sameImapEndpoint(before, after) {
  if (!before || !after) return false;
  return (
    before.host === after.host &&
    Number(before.port) === Number(after.port) &&
    before.security === after.security &&
    (Boolean(before.allowSelfSigned) || !after.allowSelfSigned)
  );
}

/** Lê o JSON da chave da conta de serviço do Google (baixado no Google Cloud). */
function serviceAccount(value) {
  let data;
  try {
    data = JSON.parse(value);
  } catch {
    throw bad('O arquivo da chave da conta de serviço deve ser o JSON baixado do Google Cloud.');
  }
  if (data?.type !== 'service_account' || !data.client_email || !data.private_key) {
    throw bad('O JSON informado não é a chave de uma conta de serviço (faltam client_email ou private_key).');
  }
  try {
    crypto.createPrivateKey(data.private_key);
  } catch {
    throw bad('A chave privada da conta de serviço é inválida.');
  }
  return { clientEmail: String(data.client_email), clientId: String(data.client_id || ''), privateKey: String(data.private_key) };
}

/**
 * Valida os dados da conexão e combina os segredos novos com os já salvos (campo vazio mantém o
 * segredo salvo). Retorna o objeto a gravar, com os segredos cifrados.
 * options: forTest (teste da conexão: o nome não é obrigatório), pending (credenciais da Microsoft
 * ainda não salvas: { signIn(id), certificate(id) }), endpoints (endereços trocados nos testes).
 */
export function parseMailSource(body = {}, existing = null, box, { forTest = false, pending = {}, endpoints = {} } = {}) {
  const type = body.type;
  if (!MAIL_TYPES[type]) throw bad('Escolha o tipo de conexão: Microsoft 365, Google Workspace ou IMAP.');
  const sameType = existing?.type === type;
  const prev = sameType ? existing.secrets || {} : {};
  const seal = (value) => box.seal(value);
  const name = text(body.name, 'o nome da conexão', { required: !forTest, max: 200 });
  // Microsoft 365, ou IMAP da Microsoft com login OAuth: credenciais do Microsoft Entra ID.
  const imapOAuth = type === 'imap' && body.imap?.auth === 'oauth';
  const microsoft = type === 'graph' || imapOAuth;
  // As credenciais do aplicativo (segredo ou certificado) valem também ao trocar entre Microsoft 365 e
  // IMAP com OAuth (o mesmo aplicativo, no mesmo locatário); a conta conectada, só para o mesmo uso.
  const hadMicrosoft = Boolean(existing?.graph) && (existing.type === 'graph' || existing.imap?.auth === 'oauth');
  const creds = microsoft
    ? microsoftCredentials(body.graph || {}, {
        previous: hadMicrosoft ? { graph: existing.graph, secrets: existing.secrets || {} } : null,
        box,
        pending,
        purpose: type === 'graph' ? 'graph' : 'imap',
        sourceId: existing?.id || null,
      })
    : null;
  const msAccount = creds?.graph.account || null; // conta Microsoft conectada
  const data = {
    name,
    type,
    description: text(body.description, 'a descrição', { max: 1000 }),
    // Com a conta conectada, as caixas são sempre as da lista (a da conta e as compartilhadas com ela).
    scope: type !== 'imap' && body.scope === 'all' && !msAccount ? 'all' : 'list',
    mailboxes: [],
    excludeMailboxes: lines(body.excludeMailboxes, 500),
    excludeFolders: lines(body.excludeFolders, 200),
    // Exclusão das mensagens encontradas: definitiva ou movendo para a lixeira.
    allowDelete: body.allowDelete === true,
    deleteMode: body.deleteMode === 'trash' ? 'trash' : 'permanent',
    // Os campos dos outros tipos ficam nulos (ao trocar o tipo, os dados antigos são descartados).
    graph: null,
    gmail: null,
    imap: null,
  };
  const secrets = {};

  if (creds) {
    data.graph = creds.graph;
    Object.assign(secrets, creds.secrets);
    // A exclusão pelo Microsoft Graph com a conta conectada precisa da permissão Mail.ReadWrite na entrada.
    if (type === 'graph' && msAccount && data.allowDelete && !forTest && !canWriteMail(msAccount.scopes)) {
      throw bad('Para permitir a exclusão com a conta conectada, conecte a conta de novo com "Permitir excluir" marcado: a entrada atual autorizou somente a leitura das mensagens (Mail.Read).');
    }
  }

  if (type === 'gmail') {
    const g = body.gmail || {};
    const json = text(g.serviceAccountJson, 'a chave da conta de serviço', { max: 20000 });
    const account = json ? serviceAccount(json) : null;
    const clientEmail = account?.clientEmail || (sameType ? existing.gmail?.clientEmail : '');
    if (!clientEmail) throw bad('Envie o arquivo JSON da chave da conta de serviço.');
    data.gmail = {
      clientEmail,
      clientId: account ? account.clientId : sameType ? existing.gmail?.clientId || '' : '',
      adminEmail: email(g.adminEmail, 'o e-mail do administrador'),
    };
    secrets.privateKey = account ? seal(account.privateKey) : prev.privateKey;
    if (!secrets.privateKey) throw bad('Envie o arquivo JSON da chave da conta de serviço.');
    if (data.scope === 'all' && !data.gmail.adminEmail) throw bad('Para analisar todas as caixas do domínio, informe o e-mail de um administrador (usado para listar os usuários).');
  }

  if (type === 'imap') {
    const i = body.imap || {};
    const host = text(i.host, 'o servidor IMAP', { required: true, max: 253 }).toLowerCase();
    if (!HOST_RE.test(host)) throw bad('Servidor IMAP inválido: informe apenas o nome ou o IP (ex.: imap.empresa.com.br).');
    const security = ['tls', 'starttls', 'none'].includes(i.security) ? i.security : 'tls';
    const port = Number(i.port) || (security === 'tls' ? 993 : 143);
    if (!Number.isInteger(port) || port < 1 || port > 65535) throw bad('Porta inválida.');
    data.imap = { host, port, security, allowSelfSigned: Boolean(i.allowSelfSigned), auth: imapOAuth ? 'oauth' : 'password' };
    // O token da Microsoft só vai para os servidores IMAP da Microsoft, com TLS verificado.
    const problem = imapOAuth ? imapOAuthProblem(data.imap, endpoints) : null;
    if (problem) throw bad(problem);
    // Senhas salvas só valem para o mesmo destino e a mesma proteção: trocar o servidor, a porta
    // ou a segurança (ou passar a aceitar certificado não confiável) exige informá-las de novo, para
    // que uma senha salva não seja enviada a outro endereço nem sem a proteção com que foi cadastrada.
    const kept = !imapOAuth && sameType && existing.imap?.auth !== 'oauth' && sameImapEndpoint(existing.imap, data.imap) ? prev : {};
    const defaultPassword = imapOAuth ? '' : text(i.defaultPassword, 'a senha padrão', { max: 1000 });
    if (!imapOAuth) {
      secrets.defaultPassword = defaultPassword ? seal(defaultPassword) : kept.defaultPassword;
      secrets.passwords = {};
    }
    const seen = new Set();
    for (const [index, m] of (Array.isArray(body.mailboxes) ? body.mailboxes : []).entries()) {
      const address = email(m?.address, `o e-mail da caixa ${index + 1}`);
      if (!address || seen.has(key(address))) continue;
      seen.add(key(address));
      // Com OAuth, o login (XOAUTH2) é o próprio e-mail da caixa: um login de senha não é guardado.
      const login = imapOAuth ? '' : text(m?.login, `o login da caixa ${address}`, { max: 320 });
      data.mailboxes.push(login && login !== address ? { address, login } : { address });
      if (imapOAuth) continue; // login pelo token da Microsoft, sem senha
      const password = text(m?.password, `a senha da caixa ${address}`, { max: 1000 });
      const saved = password ? seal(password) : kept.passwords?.[key(address)];
      if (saved) secrets.passwords[key(address)] = saved;
      else if (!secrets.defaultPassword) throw bad(`Informe a senha da caixa ${address} (ou uma senha padrão da conexão).`);
    }
    // Sem caixas informadas, a da própria conta conectada.
    if (data.mailboxes.length === 0 && msAccount?.address) data.mailboxes.push({ address: msAccount.address });
    if (data.mailboxes.length > MAX_MAILBOXES) throw bad(`Informe no máximo ${MAX_MAILBOXES} caixas.`);
  } else if (data.scope === 'list') {
    data.mailboxes = addressList(body.mailboxes, 'o e-mail da caixa').map((address) => ({ address }));
    if (data.mailboxes.length === 0 && msAccount?.address) data.mailboxes = [{ address: msAccount.address }];
  }

  if (data.scope === 'list' && data.mailboxes.length === 0) throw bad('Informe ao menos uma caixa de e-mail.');
  data.secrets = Object.fromEntries(Object.entries(secrets).filter(([, v]) => v && (typeof v === 'string' || Object.keys(v).length)));
  return data;
}

/** Conta Microsoft conectada, como mostrada na interface (sem o identificador da entrada). */
function publicAccount(account, purpose) {
  if (!account) return null;
  const { id, tenantId, username, name, address, connectedAt } = account;
  // No IMAP, a permissão do token (IMAP.AccessAsUser.All) já inclui alterar as mensagens.
  return { id, tenantId, username, name, address, connectedAt, canDelete: purpose === 'imap' || canWriteMail(account.scopes) };
}

/** Conexão sem os segredos (apenas indica quais estão salvos). */
export function publicMailSource(source) {
  const { secrets = {}, ...rest } = source;
  const out = {
    ...rest,
    typeLabel: MAIL_TYPES[source.type] || source.type,
    mailboxes: (source.mailboxes || []).map((m) => ({ ...m, hasPassword: Boolean(secrets.passwords?.[key(m.address)]) })),
  };
  if (source.graph) {
    out.graph = {
      ...source.graph,
      auth: source.graph.auth || 'secret',
      account: publicAccount(source.graph.account, source.type === 'imap' ? 'imap' : 'graph'),
      hasClientSecret: Boolean(secrets.clientSecret),
      hasCertificateKey: Boolean(secrets.certificateKey),
      hasRefreshToken: Boolean(secrets.refreshToken),
    };
  }
  if (source.gmail) out.gmail = { ...source.gmail, hasPrivateKey: Boolean(secrets.privateKey) };
  if (source.imap) out.imap = { ...source.imap, hasDefaultPassword: Boolean(secrets.defaultPassword) };
  return out;
}

export function mailSourcesRouter({ store, manager = null, scheduler = null, endpoints = {} }) {
  const router = Router();
  const pending = new PendingCredentials();
  const login = endpoints.graphLogin || GRAPH_ENDPOINTS.login;
  // Credenciais ainda não salvas, para parseMailSource: a entrada da conta (com a situação dela) e o
  // certificado gerado ou importado.
  const lookups = {
    signIn: (id) => pending.get(id, 'device'),
    certificate: (id) => pending.get(id, 'certificate'),
  };
  const parse = (body, existing, options = {}) => parseMailSource(body, existing, store.secrets, { ...options, pending: lookups, endpoints });
  // Depois de salvar, a entrada e o certificado usados saem da memória.
  const forgetPending = (body) => {
    const g = body?.graph || {};
    if (pending.get(g.signIn, 'device')) pending.delete(g.signIn);
    if (pending.get(g.certificateId, 'certificate')) pending.delete(g.certificateId);
  };

  const find = (id) => {
    const source = store.getMailSource(id);
    if (!source) throw new HttpError(404, 'Conexão de e-mail não encontrada.');
    return source;
  };

  // ---------- Conta Microsoft conectada: entrada pelo código de dispositivo ----------
  //
  // O CLEAN pede um código à Microsoft; a pessoa abre o endereço mostrado (em qualquer navegador),
  // digita o código e entra com a conta. Enquanto isso, a tela consulta o resultado (POST, com o
  // identificador no corpo: não fica nos registros de acesso de um proxy); ao terminar, o token de
  // atualização fica na memória do servidor até a conexão ser salva (gravado cifrado). A entrada fica
  // presa à conexão em que foi começada (ou a uma conexão nova).

  router.post('/oauth/device', async (req, res) => {
    const body = req.body || {};
    const purpose = body.type === 'imap' ? 'imap' : 'graph';
    const sourceId = typeof body.id === 'string' && body.id ? find(body.id).id : null;
    const { tenantId, clientId } = microsoftApp(body.graph || {}, { delegated: true });
    const write = body.allowDelete === true;
    const scopes = delegatedScopes(purpose, { write, tenantId });
    let start;
    try {
      start = await startDeviceCode({ login, tenantId, clientId, scopes });
    } catch (err) {
      throw bad(microsoftError(err));
    }
    const now = Date.now();
    const flow = pending.add(
      'device',
      {
        sourceId,
        tenantId,
        clientId,
        purpose,
        scopes,
        write,
        // Identificador da autorização: a gravação dos tokens renovados confere se ainda é a da conexão.
        grantId: crypto.randomUUID(),
        deviceCode: start.deviceCode,
        interval: start.interval,
        nextPoll: now + start.interval * 1000,
        codeExpires: now + start.expiresIn * 1000,
        status: 'pending',
      },
      start.expiresIn * 1000 + SIGN_IN_TTL_MS,
    );
    res.status(201).json({ flowId: flow.id, userCode: start.userCode, verificationUri: start.verificationUri, expiresAt: new Date(flow.codeExpires).toISOString(), interval: start.interval, allowDelete: write });
  });

  /** A entrada terminou: guarda o token e os dados da conta (e, no Microsoft 365, o e-mail da caixa). */
  async function completeSignIn(flow, tokens) {
    const info = accountFromTokens(tokens);
    let address = info.username;
    if (flow.purpose === 'graph') {
      try {
        const me = await request(`${endpoints.graph || GRAPH_ENDPOINTS.graph}/me?$select=displayName,mail,userPrincipalName`, {
          headers: { Authorization: `Bearer ${tokens.access_token}` },
          retries: 2,
          timeoutMs: 30000,
        });
        address = me?.mail || me?.userPrincipalName || address;
        info.name ||= me?.displayName || '';
        info.username ||= me?.userPrincipalName || '';
      } catch {
        // sem os dados do perfil: fica o nome de logon informado na entrada
      }
    }
    // Permissões concedidas (informadas pela Microsoft; sem elas, as pedidas na entrada).
    const granted = grantedScopes(tokens);
    const asked = flow.scopes.filter((scope) => !['openid', 'profile', 'offline_access', 'email'].includes(scope));
    flow.account = { ...info, address, scopes: granted.length ? granted : asked, grantId: flow.grantId, purpose: flow.purpose, connectedAt: new Date().toISOString() };
    flow.refreshToken = tokens.refresh_token;
    flow.status = 'connected';
    flow.deviceCode = null;
    flow.notice = null;
    flow.expiresAt = Date.now() + SIGN_IN_TTL_MS;
  }

  /**
   * Consulta a Microsoft (respeitando o intervalo pedido por ela) e atualiza a situação da entrada.
   * Uma falha passageira (rede, tempo esgotado, erro temporário da Microsoft) não encerra a entrada.
   */
  async function advanceSignIn(flow) {
    if (flow.status !== 'pending') return;
    if (Date.now() >= flow.codeExpires) {
      Object.assign(flow, { status: 'failed', error: 'O código expirou antes de a entrada terminar: clique em "Conectar conta" de novo.', deviceCode: null });
      return;
    }
    if (flow.polling) return flow.polling;
    if (Date.now() < flow.nextPoll) return;
    flow.polling = (async () => {
      try {
        const result = await pollDeviceCode({ login, tenantId: flow.tenantId, clientId: flow.clientId, deviceCode: flow.deviceCode });
        flow.notice = null;
        if (result.status === 'slow_down') flow.interval = Math.min(flow.interval + 5, 60);
        if (result.status === 'connected') await completeSignIn(flow, result.tokens);
      } catch (err) {
        if (err?.retryable || err?.name === 'TimeoutError') flow.notice = `Falha momentânea ao consultar a Microsoft (${microsoftError(err).replace(/\.$/, '')}). Tentando de novo…`;
        else Object.assign(flow, { status: 'failed', error: microsoftError(err), deviceCode: null });
      } finally {
        flow.nextPoll = Date.now() + flow.interval * 1000;
        flow.polling = null;
      }
    })();
    return flow.polling;
  }

  router.post('/oauth/device/status', async (req, res) => {
    const flow = pending.get(req.body?.flowId, 'device');
    if (!flow) return res.json({ status: 'failed', error: 'A entrada expirou ou foi cancelada: clique em "Conectar conta" de novo.' });
    await advanceSignIn(flow);
    res.json({
      status: flow.status,
      ...(flow.status === 'connected' ? { account: publicAccount(flow.account, flow.purpose) } : {}),
      ...(flow.status === 'failed' ? { error: flow.error } : {}),
      ...(flow.status === 'pending' && flow.notice ? { notice: flow.notice } : {}),
    });
  });

  router.post('/oauth/device/cancel', (req, res) => {
    const id = req.body?.flowId;
    if (pending.get(id, 'device')) pending.delete(id);
    res.status(204).end();
  });

  // ---------- Certificado do aplicativo: gerado pelo CLEAN ou importado (PEM) ----------

  router.post('/certificate', async (req, res) => {
    const body = req.body || {};
    let result;
    if (typeof body.pem === 'string' && body.pem.trim()) {
      if (body.pem.length > 200000) throw bad('O arquivo é grande demais para um certificado.');
      try {
        result = importCertificate(body.pem);
      } catch (err) {
        throw bad(err.message);
      }
    } else {
      // O nome da conexão só dá nome ao certificado (CN, até 64 caracteres com o "CLEAN - "): um nome
      // longo é cortado.
      const name = String(body.name ?? '').replace(/[^\p{L}\p{N} ._-]/gu, '').trim().slice(0, 56).trim();
      result = await createCertificate({ commonName: name ? `CLEAN - ${name}` : 'CLEAN' });
    }
    // O certificado ainda não salvo que este substitui sai da memória.
    if (pending.get(body.replaces, 'certificate')) pending.delete(body.replaces);
    const item = pending.add('certificate', result, CERTIFICATE_TTL_MS);
    res.status(201).json({ certificateId: item.id, certificate: result.certificate });
  });

  router.post('/certificate/discard', (req, res) => {
    const id = req.body?.certificateId;
    if (pending.get(id, 'certificate')) pending.delete(id);
    res.status(204).end();
  });

  router.get('/', (req, res) => {
    const list = store
      .listMailSources()
      .map(publicMailSource)
      .sort((a, b) => a.name.localeCompare(b.name, 'pt-BR'));
    res.json(list);
  });

  router.post('/', (req, res) => {
    const source = store.createMailSource(parse(req.body, null));
    forgetPending(req.body);
    res.status(201).json(publicMailSource(source));
  });

  // Testa a conexão com os dados do formulário (campos de senha vazios usam os segredos salvos).
  router.post('/test', async (req, res) => {
    const id = typeof req.body?.id === 'string' ? req.body.id : '';
    const existing = id ? find(id) : null;
    const data = parse(req.body, existing, { forTest: true });
    let secrets;
    try {
      secrets = store.openMailSecrets(data);
    } catch (err) {
      return res.json({ ok: false, message: err.message, details: [] });
    }
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(new Error('timeout')), 60000);
    // Conta conectada: o token renovado no teste substitui o da entrada (ainda não salva) ou o salvo.
    const grantId = data.graph?.account?.grantId;
    const signIn = req.body?.graph?.signIn;
    const onRefreshToken = grantId
      ? (token) => {
          const flow = pending.get(signIn, 'device');
          if (flow && flow.grantId === grantId) flow.refreshToken = token;
          else if (existing) store.saveRefreshToken(existing.id, grantId, token);
        }
      : undefined;
    try {
      const connector = createConnector({ ...data, id: existing?.id || 'teste', secrets }, { signal: controller.signal, endpoints, onRefreshToken });
      const result = await connector.test();
      // A leitura funciona, mas salvar com "Permitir excluir" seria recusado: o teste já avisa.
      const account = data.type === 'graph' ? data.graph?.account : null;
      if (result.ok && account && data.allowDelete && !canWriteMail(account.scopes)) {
        return res.json({
          ok: false,
          message: 'A leitura das caixas funciona, mas a exclusão está permitida e a conta conectada autorizou só a leitura das mensagens: clique em "Conectar conta" de novo com "Permitir excluir" marcado (sem isso, a conexão não pode ser salva).',
          details: result.details,
        });
      }
      res.json(result);
    } catch (err) {
      const message = controller.signal.aborted ? 'Tempo esgotado ao testar a conexão (60 s).' : friendlyError(err);
      res.json({ ok: false, message, details: [] });
    } finally {
      clearTimeout(timer);
    }
  });

  router.get('/:id', (req, res) => {
    res.json(publicMailSource(find(req.params.id)));
  });

  /**
   * Repositórios do OneDrive/SharePoint que usam as credenciais desta conexão: recebem as novas
   * credenciais (ex.: um segredo renovado ou um certificado novo) ou deixam de estar ligados
   * (conexão removida, de outro tipo ou com a conta conectada, cujas permissões são só de e-mail).
   */
  const syncLinkedRepositories = (source) => {
    for (const repo of store.listRepositories()) {
      if (!repo.credentialsFrom || repo.credentialsFrom !== source.id) continue;
      const linked = source.deleted ? null : linkedCredentials(source);
      if (!linked) {
        store.updateRepository(repo.id, { credentialsFrom: null });
        continue;
      }
      const tenantChanged = String(repo.graph?.tenantId || '').toLowerCase() !== String(source.graph.tenantId).toLowerCase();
      store.updateRepository(repo.id, linked);
      if (tenantChanged && repo.allowDelete) manager?.revokeDeletion('repository', repo.id, 'o locatário das credenciais foi alterado');
    }
  };

  router.put('/:id', (req, res) => {
    const existing = find(req.params.id);
    const before = { allowDelete: existing.allowDelete, deleteMode: existing.deleteMode };
    const data = parse(req.body, existing);
    // Agendamentos com exclusão automática afetados (inclusive por repositórios ligados a estas
    // credenciais) ficam suspensos, com aviso.
    const { result: updated, warning } = checkDeleteSchedules(store, () => {
      const saved = store.updateMailSource(existing.id, data);
      syncLinkedRepositories(saved);
      return saved;
    }, scheduler);
    forgetPending(req.body);
    // Análises em andamento deixam de excluir se a exclusão foi desligada ou mudou de forma.
    if (before.allowDelete && (!updated.allowDelete || updated.deleteMode !== before.deleteMode)) {
      manager?.revokeDeletion('mail', existing.id, updated.allowDelete ? 'a forma de exclusão da conexão foi alterada' : 'a opção "Permitir exclusão" foi desligada');
    }
    res.json({ ...publicMailSource(updated), ...(warning ? { scheduleWarning: warning } : {}) });
  });

  router.delete('/:id', (req, res) => {
    const existing = find(req.params.id);
    assertUnused(store, 'mail', existing.id, 'A conexão de e-mail');
    store.deleteMailSource(existing.id);
    syncLinkedRepositories({ ...existing, deleted: true });
    if (existing.allowDelete) manager?.revokeDeletion('mail', existing.id, 'a conexão foi removida do cadastro');
    res.status(204).end();
  });

  return router;
}
