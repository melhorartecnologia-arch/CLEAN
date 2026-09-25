// Rotas /api/mail-sources: conexões com caixas de e-mail (Microsoft 365, Google Workspace, IMAP).
// Os segredos (segredo do cliente, chave da conta de serviço, senhas) são gravados cifrados e nunca
// voltam para o navegador: a API informa apenas se cada um está salvo.
import crypto from 'node:crypto';
import { Router } from 'express';
import { HttpError, bad, text, lines } from './validate.js';
import { createConnector, MAIL_TYPES } from '../mail/connectors.js';
import { friendlyError } from '../scan/errors.js';

const EMAIL_RE = /^[^\s@<>()",;:]+@[^\s@<>()",;:]+$/;
const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const DOMAIN_RE = /^(?=.{3,253}$)[a-z0-9-]+(\.[a-z0-9-]+)+$/i;
const HOST_RE = /^(?=.{1,253}$)[a-z0-9]([a-z0-9-]*[a-z0-9])?(\.[a-z0-9]([a-z0-9-]*[a-z0-9])?)*$|^\[?[0-9a-f:.]+\]?$/i;
const MAX_MAILBOXES = 5000;

const key = (address) => String(address || '').trim().toLowerCase();

function email(value, field) {
  const v = text(value, field, { max: 320 });
  if (v && !EMAIL_RE.test(v)) throw bad(`${field[0].toUpperCase()}${field.slice(1)} inválido: "${v.slice(0, 80)}".`);
  return v;
}

/** Lista de endereços (texto com um por linha ou lista), sem repetições. */
function addressList(value, field) {
  const items = Array.isArray(value) ? value.map((v) => (typeof v === 'object' && v ? v.address : v)) : String(value || '').split(/[\r\n,;]+/);
  const seen = new Set();
  const out = [];
  for (const raw of items) {
    const address = email(raw, field);
    if (!address || seen.has(key(address))) continue;
    seen.add(key(address));
    out.push(address);
  }
  if (out.length > MAX_MAILBOXES) throw bad(`Informe no máximo ${MAX_MAILBOXES} caixas.`);
  return out;
}

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
 */
export function parseMailSource(body = {}, existing = null, box, { forTest = false } = {}) {
  const type = body.type;
  if (!MAIL_TYPES[type]) throw bad('Escolha o tipo de conexão: Microsoft 365, Google Workspace ou IMAP.');
  const sameType = existing?.type === type;
  const prev = sameType ? existing.secrets || {} : {};
  const seal = (value) => box.seal(value);
  const data = {
    name: text(body.name, 'o nome da conexão', { required: !forTest, max: 200 }),
    type,
    description: text(body.description, 'a descrição', { max: 1000 }),
    scope: type !== 'imap' && body.scope === 'all' ? 'all' : 'list',
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

  if (type === 'graph') {
    const g = body.graph || {};
    const tenantId = text(g.tenantId, 'o ID do locatário', { required: true, max: 255 });
    if (!GUID_RE.test(tenantId) && !DOMAIN_RE.test(tenantId)) throw bad('ID do locatário inválido: use o GUID (ID do diretório) ou o domínio, ex.: empresa.onmicrosoft.com.');
    const clientId = text(g.clientId, 'o ID do cliente (aplicativo)', { required: true, max: 64 });
    if (!GUID_RE.test(clientId)) throw bad('ID do cliente inválido: use o "ID do aplicativo (cliente)" do registro do aplicativo.');
    data.graph = { tenantId, clientId };
    const secret = text(g.clientSecret, 'o segredo do cliente', { max: 2000 });
    secrets.clientSecret = secret ? seal(secret) : prev.clientSecret;
    if (!secrets.clientSecret) throw bad('Informe o segredo do cliente (valor do segredo criado no registro do aplicativo).');
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
    data.imap = { host, port, security, allowSelfSigned: Boolean(i.allowSelfSigned) };
    // Senhas salvas só valem para o mesmo destino e a mesma proteção: trocar o servidor, a porta
    // ou a segurança (ou passar a aceitar certificado não confiável) exige informá-las de novo, para
    // que uma senha salva não seja enviada a outro endereço nem sem a proteção com que foi cadastrada.
    const kept = sameType && sameImapEndpoint(existing.imap, data.imap) ? prev : {};
    const defaultPassword = text(i.defaultPassword, 'a senha padrão', { max: 1000 });
    secrets.defaultPassword = defaultPassword ? seal(defaultPassword) : kept.defaultPassword;
    secrets.passwords = {};
    const seen = new Set();
    for (const [index, m] of (Array.isArray(body.mailboxes) ? body.mailboxes : []).entries()) {
      const address = email(m?.address, `o e-mail da caixa ${index + 1}`);
      if (!address || seen.has(key(address))) continue;
      seen.add(key(address));
      const login = text(m?.login, `o login da caixa ${address}`, { max: 320 });
      data.mailboxes.push(login && login !== address ? { address, login } : { address });
      const password = text(m?.password, `a senha da caixa ${address}`, { max: 1000 });
      const saved = password ? seal(password) : kept.passwords?.[key(address)];
      if (saved) secrets.passwords[key(address)] = saved;
      else if (!secrets.defaultPassword) throw bad(`Informe a senha da caixa ${address} (ou uma senha padrão da conexão).`);
    }
    if (data.mailboxes.length > MAX_MAILBOXES) throw bad(`Informe no máximo ${MAX_MAILBOXES} caixas.`);
  } else if (data.scope === 'list') {
    data.mailboxes = addressList(body.mailboxes, 'o e-mail da caixa').map((address) => ({ address }));
  }

  if (data.scope === 'list' && data.mailboxes.length === 0) throw bad('Informe ao menos uma caixa de e-mail.');
  data.secrets = Object.fromEntries(Object.entries(secrets).filter(([, v]) => v && (typeof v === 'string' || Object.keys(v).length)));
  return data;
}

/** Conexão sem os segredos (apenas indica quais estão salvos). */
export function publicMailSource(source) {
  const { secrets = {}, ...rest } = source;
  const out = {
    ...rest,
    typeLabel: MAIL_TYPES[source.type] || source.type,
    mailboxes: (source.mailboxes || []).map((m) => ({ ...m, hasPassword: Boolean(secrets.passwords?.[key(m.address)]) })),
  };
  if (source.graph) out.graph = { ...source.graph, hasClientSecret: Boolean(secrets.clientSecret) };
  if (source.gmail) out.gmail = { ...source.gmail, hasPrivateKey: Boolean(secrets.privateKey) };
  if (source.imap) out.imap = { ...source.imap, hasDefaultPassword: Boolean(secrets.defaultPassword) };
  return out;
}

export function mailSourcesRouter({ store, endpoints = {} }) {
  const router = Router();

  const find = (id) => {
    const source = store.getMailSource(id);
    if (!source) throw new HttpError(404, 'Conexão de e-mail não encontrada.');
    return source;
  };

  router.get('/', (req, res) => {
    const list = store
      .listMailSources()
      .map(publicMailSource)
      .sort((a, b) => a.name.localeCompare(b.name, 'pt-BR'));
    res.json(list);
  });

  router.post('/', (req, res) => {
    const source = store.createMailSource(parseMailSource(req.body, null, store.secrets));
    res.status(201).json(publicMailSource(source));
  });

  // Testa a conexão com os dados do formulário (campos de senha vazios usam os segredos salvos).
  router.post('/test', async (req, res) => {
    const id = typeof req.body?.id === 'string' ? req.body.id : '';
    const existing = id ? find(id) : null;
    const data = parseMailSource(req.body, existing, store.secrets, { forTest: true });
    let secrets;
    try {
      secrets = store.openMailSecrets(data);
    } catch (err) {
      return res.json({ ok: false, message: err.message, details: [] });
    }
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(new Error('timeout')), 60000);
    try {
      const connector = createConnector({ ...data, id: existing?.id || 'teste', secrets }, { signal: controller.signal, endpoints });
      res.json(await connector.test());
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

  router.put('/:id', (req, res) => {
    const existing = find(req.params.id);
    res.json(publicMailSource(store.updateMailSource(existing.id, parseMailSource(req.body, existing, store.secrets))));
  });

  router.delete('/:id', (req, res) => {
    if (!store.deleteMailSource(req.params.id)) throw new HttpError(404, 'Conexão de e-mail não encontrada.');
    res.status(204).end();
  });

  return router;
}
