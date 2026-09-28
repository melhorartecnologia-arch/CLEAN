// OAuth 2.0 da Microsoft: aplicativo com certificado (asserção PS256), conta conectada (código de
// dispositivo, /me, caixas compartilhadas, renovação do token gravada cifrada) e IMAP com XOAUTH2 —
// contra os simuladores do Entra ID, do Graph e de um servidor IMAP.
import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import crypto from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { MailScanner } from '../src/mail/scanner.js';
import { GraphConnector } from '../src/mail/graph.js';
import { imapError } from '../src/mail/imap.js';
import { createCertificate, importCertificate } from '../src/cloud/certificate.js';
import { aadError, delegatedScopes, clientAssertion, MicrosoftAuth } from '../src/cloud/microsoft-auth.js';
import { ApiError } from '../src/mail/http.js';
import { mailDeletionScope } from '../src/scan/delete.js';
import { createApp } from '../src/app.js';
import { Store } from '../src/store.js';
import { ScanManager } from '../src/scan/manager.js';
import { startMockApis, imapTokenUser } from './helpers/mock-apis.js';
import { startFakeImap } from './helpers/fake-imap.js';

const TENANT = 'contoso.onmicrosoft.com';
const CLIENT = '11111111-2222-3333-4444-555555555555';
const SECRET = 'segredo-do-aplicativo';
const G = 'https://graph.microsoft.com/';
const IMAP_SCOPE = 'https://outlook.office.com/IMAP.AccessAsUser.All';
const READ_WRITE = [`${G}User.Read`, `${G}Mail.ReadWrite`, `${G}Mail.ReadWrite.Shared`];
const READ_ONLY = [`${G}User.Read`, `${G}Mail.Read`, `${G}Mail.Read.Shared`];
const TERMS = [{ id: 'l:conf', type: 'text', value: 'confidencial', listName: 'Sigilo' }];

function mail(subject, body) {
  return Buffer.from(
    [`From: Ana <ana@contoso.com>`, 'To: rh@contoso.com', `Subject: ${subject}`, 'Date: Thu, 25 Sep 2026 10:00:00 -0300', `Message-ID: <${crypto.randomUUID()}@contoso.com>`, 'Content-Type: text/plain; charset=utf-8', '', body, ''].join('\r\n'),
  );
}

function graphData() {
  const folders = [
    { id: 'inbox', displayName: 'Caixa de Entrada', wellKnown: 'inbox' },
    { id: 'trash', displayName: 'Itens Excluídos', wellKnown: 'deleteditems' },
  ];
  const box = (prefix, extra = {}) => ({
    folders,
    messages: {
      inbox: [
        { id: `${prefix}1`, received: '2026-09-20T10:00:00Z', raw: mail(`Relatório ${prefix}`, 'Documento CONFIDENCIAL.') },
        { id: `${prefix}2`, received: '2026-09-21T10:00:00Z', raw: mail(`Almoço ${prefix}`, 'Nada de mais.') },
      ],
      trash: [],
    },
    ...extra,
  });
  return {
    tenant: TENANT,
    clientId: CLIENT,
    secret: SECRET,
    users: [
      { id: 'u-ana', mail: 'ana@contoso.com', displayName: 'Ana Souza', ...box('a') },
      { id: 'u-rh', mail: 'rh@contoso.com', displayName: 'RH (compartilhada)', sharedWith: ['ana@contoso.com'], ...box('r') },
      { id: 'u-bia', mail: 'bia@contoso.com', displayName: 'Bia', ...box('b') },
    ],
    delegated: { signInAs: 'ana@contoso.com', pendingPolls: 1, refreshTokens: {}, issued: [] },
  };
}

let mocks;
let graph;
let root;
let cert;

before(async () => {
  root = fs.mkdtempSync(path.join(os.tmpdir(), 'clean-oauth-'));
  graph = graphData();
  mocks = await startMockApis({ graph });
  cert = await createCertificate({ commonName: 'CLEAN - Teste' });
});

after(async () => {
  await mocks?.close();
  fs.rmSync(root, { recursive: true, force: true });
});

/** Um token de atualização válido no simulador, para a conta e as permissões dadas. */
function refreshTokenFor(user, scopes) {
  const token = `rt-${crypto.randomUUID()}`;
  graph.delegated.refreshTokens[token] = { user, scopes };
  return token;
}

async function runMail(sources, endpoints = mocks.endpoints, options = {}) {
  const messages = [];
  const scanner = new MailScanner({ sources, terms: TERMS, options, endpoints }, (m) => messages.push(m));
  const stats = await scanner.run();
  return {
    stats,
    records: messages.filter((m) => m.type === 'results').flatMap((m) => m.records),
    errors: messages.filter((m) => m.type === 'errors').flatMap((m) => m.items),
    logs: messages.filter((m) => m.type === 'log'),
    credentials: messages.filter((m) => m.type === 'credentials'),
  };
}

const delegatedSource = (extra = {}) => ({
  id: 'src-conta',
  name: 'Conta da Ana',
  type: 'graph',
  scope: 'list',
  mailboxes: [{ address: 'ana@contoso.com' }, { address: 'rh@contoso.com' }],
  excludeMailboxes: [],
  excludeFolders: [],
  graph: { tenantId: TENANT, clientId: CLIENT, auth: 'delegated', account: { id: 'u-ana', username: 'ana@contoso.com', address: 'ana@contoso.com', name: 'Ana Souza', scopes: READ_WRITE, grantId: 'grant-1' } },
  secrets: { refreshToken: refreshTokenFor('ana@contoso.com', READ_WRITE) },
  ...extra,
});

// ---------------------------------------------------------------------------------------------

test('certificado: X.509 autoassinado válido, impressões digitais e importação em PEM', async () => {
  const x = new crypto.X509Certificate(cert.certificate.pem);
  assert.ok(x.verify(x.publicKey), 'assinado pela própria chave');
  assert.ok(x.checkPrivateKey(crypto.createPrivateKey(cert.privateKeyPem)), 'a chave privada é a do certificado');
  assert.equal(x.ca, false);
  assert.match(cert.certificate.thumbprint, /^[0-9A-F]{40}$/, 'SHA-1 em hexadecimal, como no Entra ID');
  assert.match(cert.certificate.thumbprint256, /^[\w-]{43}$/, 'SHA-256 em base64url (x5t#S256)');
  assert.equal(cert.certificate.subject, 'CN=CLEAN - Teste');
  const years = (Date.parse(cert.certificate.notAfter) - Date.now()) / (365.25 * 86400000);
  assert.ok(years > 1.9 && years < 2.1, `válido por 2 anos (${years})`);
  assert.ok(!cert.certificate.pem.includes('PRIVATE'), 'os dados públicos não têm a chave');

  const imported = importCertificate(`${cert.certificate.pem}\n${cert.privateKeyPem}`);
  assert.equal(imported.certificate.thumbprint, cert.certificate.thumbprint);
  const other = crypto.generateKeyPairSync('rsa', { modulusLength: 2048 }).privateKey.export({ type: 'pkcs8', format: 'pem' });
  assert.throws(() => importCertificate(`${cert.certificate.pem}${other}`), /não é a do certificado/);
  assert.throws(() => importCertificate(cert.certificate.pem), /não tem a chave privada/);
  assert.throws(() => importCertificate('qualquer coisa'), /não tem um certificado/);
  assert.throws(() => importCertificate(`${cert.certificate.pem}-----BEGIN ENCRYPTED PRIVATE KEY-----\nAA==\n-----END ENCRYPTED PRIVATE KEY-----`), /protegida por senha/);
});

test('asserção do cliente: JWT PS256 com x5t#S256, prazo de 10 minutos e assinatura conferível', () => {
  const jwt = clientAssertion({ clientId: CLIENT, audience: 'https://login/t/oauth2/v2.0/token', privateKey: cert.privateKeyPem, thumbprint256: cert.certificate.thumbprint256 });
  const [h, p, sig] = jwt.split('.');
  const header = JSON.parse(Buffer.from(h, 'base64url'));
  const claims = JSON.parse(Buffer.from(p, 'base64url'));
  assert.deepEqual(header, { alg: 'PS256', typ: 'JWT', 'x5t#S256': cert.certificate.thumbprint256 });
  assert.equal(claims.iss, CLIENT);
  assert.equal(claims.sub, CLIENT);
  assert.equal(claims.aud, 'https://login/t/oauth2/v2.0/token');
  assert.equal(claims.exp - claims.nbf, 600);
  assert.ok(claims.jti);
  const key = new crypto.X509Certificate(cert.certificate.pem).publicKey;
  assert.ok(crypto.verify('sha256', Buffer.from(`${h}.${p}`), { key, padding: crypto.constants.RSA_PKCS1_PSS_PADDING, saltLength: crypto.constants.RSA_PSS_SALTLEN_DIGEST }, Buffer.from(sig, 'base64url')));
});

test('permissões pedidas na entrada: leitura ou escrita, compartilhadas só com o locatário da organização', () => {
  assert.deepEqual(delegatedScopes('graph', { tenantId: TENANT }), [...READ_ONLY, 'openid', 'profile', 'offline_access']);
  assert.deepEqual(delegatedScopes('graph', { write: true, tenantId: 'organizations' }), [...READ_WRITE, 'openid', 'profile', 'offline_access']);
  assert.deepEqual(delegatedScopes('graph', { tenantId: 'consumers' }), [`${G}User.Read`, `${G}Mail.Read`, 'openid', 'profile', 'offline_access']);
  assert.deepEqual(delegatedScopes('imap', { tenantId: TENANT }), [IMAP_SCOPE, 'openid', 'profile', 'offline_access']);
});

test('erros do Entra ID viram a providência a tomar', () => {
  const err = (message, code = '') => new ApiError(message, { status: 400, code });
  assert.match(aadError(err("AADSTS7000218: The request body must contain the following parameter: 'client_assertion' or 'client_secret'.")).message, /Permitir fluxos de clientes públicos/);
  assert.match(aadError(err('AADSTS700027: Client assertion failed signature validation.')).message, /envie o arquivo do certificado/);
  assert.match(aadError(err('AADSTS700027: Client assertion failed signature validation. Reason - The key used is expired.')).message, /certificado do aplicativo venceu/);
  // Os identificadores do Entra ID ficam na mensagem, para o administrador achar a entrada nos logs.
  const traced = aadError(err('AADSTS7000215: Invalid client secret provided. Trace ID: 0a1b2c3d-1111-2222-3333-444455556666 Correlation ID: 9f8e7d6c-aaaa-bbbb-cccc-ddddeeeeffff Timestamp: 2026-09-28'));
  assert.match(traced.message, /Segredo do cliente inválido.*AADSTS7000215 · Trace ID 0a1b2c3d-1111-2222-3333-444455556666 · Correlation ID 9f8e7d6c-aaaa-bbbb-cccc-ddddeeeeffff/);
  assert.match(aadError(err('AADSTS700082: The refresh token has expired due to inactivity.', 'invalid_grant'), { delegated: true }).message, /90 dias.*Conectar conta/);
  assert.match(aadError(err('AADSTS99999: algo novo', 'invalid_grant'), { delegated: true }).message, /não vale mais.*Conectar conta/);
  assert.equal(aadError(err('AADSTS99999: algo novo', 'invalid_grant')), null, 'sem conta conectada, fica o erro original');
  assert.match(imapError({ authenticationFailed: true, responseText: 'LOGIN failed.' }, 'outlook.office365.com').message, /não aceitam mais senha.*OAuth 2\.0 da Microsoft/);
  assert.match(imapError({ authenticationFailed: true }, 'imap.empresa.com').message, /^Usuário ou senha recusados/);
  assert.match(imapError({ authenticationFailed: true }, 'outlook.office365.com', { oauth: true }).message, /recusou o login OAuth.*IMAP está habilitado/);
});

test('alcance da exclusão: igual ao de antes com aplicativo; muda com outra conta conectada', () => {
  const app = { type: 'graph', scope: 'all', mailboxes: [], graph: { tenantId: TENANT, clientId: CLIENT } };
  assert.equal(mailDeletionScope(app), `graph|${TENANT}|all|`, 'conexões já cadastradas: o mesmo valor (agendamentos confirmados continuam valendo)');
  assert.equal(mailDeletionScope({ ...app, graph: { ...app.graph, auth: 'certificate' } }), mailDeletionScope(app), 'trocar o segredo pelo certificado não muda o alcance');
  const ana = { ...app, scope: 'list', mailboxes: [{ address: 'rh@contoso.com' }], graph: { ...app.graph, auth: 'delegated', account: { id: 'u-ana' } } };
  assert.notEqual(mailDeletionScope(ana), mailDeletionScope({ ...ana, graph: { ...ana.graph, account: { id: 'u-bia' } } }));
  const imap = { type: 'imap', scope: 'list', mailboxes: [{ address: 'ana@contoso.com' }], imap: { host: 'outlook.office365.com', port: 993 } };
  assert.equal(mailDeletionScope(imap), 'imap|outlook.office365.com:993|list|ana@contoso.com');
  assert.notEqual(mailDeletionScope({ ...imap, imap: { ...imap.imap, auth: 'oauth' }, graph: ana.graph }), mailDeletionScope(imap));
});

test('Microsoft 365 com certificado: o Entra ID confere a asserção e a análise funciona', async () => {
  graph.certificatePem = cert.certificate.pem;
  const source = {
    id: 'src-cert',
    name: 'Aplicativo com certificado',
    type: 'graph',
    scope: 'all',
    mailboxes: [],
    excludeMailboxes: [],
    excludeFolders: [],
    graph: { tenantId: TENANT, clientId: CLIENT, auth: 'certificate', certificate: cert.certificate },
    secrets: { certificateKey: cert.privateKeyPem },
  };
  const before = graph.assertions || 0;
  const { records, errors } = await runMail([source]);
  assert.deepEqual(errors, []);
  assert.deepEqual(records.map((r) => `${r.mailbox}:${r.subject}`).sort(), ['ana@contoso.com:Relatório a', 'bia@contoso.com:Relatório b', 'rh@contoso.com:Relatório r']);
  assert.ok(graph.assertions > before, 'autenticou com a asserção assinada');

  // Outro certificado (não enviado ao aplicativo): erro claro, sem ler caixa nenhuma.
  const other = await createCertificate();
  const wrong = await runMail([{ ...source, graph: { ...source.graph, certificate: other.certificate }, secrets: { certificateKey: other.privateKeyPem } }]);
  assert.equal(wrong.records.length, 0);
  assert.match(wrong.errors[0].message, /certificado não foi reconhecido.*AADSTS700027/);
});

test('Microsoft 365 com a conta conectada: /me, caixa compartilhada, sem acesso às outras e token renovado', async () => {
  const source = delegatedSource({ mailboxes: [{ address: 'ana@contoso.com' }, { address: 'rh@contoso.com' }, { address: 'bia@contoso.com' }, { address: 'naoexiste@contoso.com' }] });
  graph.delegatedCalls = [];
  const { records, errors, credentials, stats } = await runMail([source]);
  assert.deepEqual(records.map((r) => `${r.mailbox}:${r.subject}`).sort(), ['ana@contoso.com:Relatório a', 'rh@contoso.com:Relatório r']);
  assert.ok(graph.delegatedCalls.some((c) => c.startsWith('GET /me/mailFolders')), 'a caixa da própria conta é lida por /me');
  assert.ok(graph.delegatedCalls.some((c) => c.startsWith('GET /users/rh@contoso.com/mailFolders')), 'a compartilhada, pelo endereço');
  assert.ok(!graph.delegatedCalls.some((c) => /^GET \/users(\?|$)/.test(c)), 'não lista os usuários do locatário');
  assert.equal(errors.length, 2);
  assert.equal(errors[0].path, 'bia@contoso.com');
  assert.match(errors[0].message, /conta conectada não tem acesso a ela/);
  // Endereço desconhecido (ou diferente do nome de logon): erro, e não "caixa ignorada".
  assert.equal(errors[1].path, 'naoexiste@contoso.com');
  assert.match(errors[1].message, /não foi encontrada para a conta conectada: confira o endereço/);
  assert.equal(stats.mailboxesSkipped, 0);
  assert.equal(stats.gaps, 2);
  // A Microsoft devolve um novo token de atualização a cada renovação: vai para o servidor gravar.
  assert.equal(credentials.length, 1);
  assert.equal(credentials[0].sourceId, 'src-conta');
  assert.equal(credentials[0].grantId, 'grant-1');
  assert.equal(credentials[0].refreshToken, graph.delegated.issued.at(-1));
  assert.notEqual(credentials[0].refreshToken, source.secrets.refreshToken);

  // Sem Mail.ReadWrite, a exclusão é recusada com a orientação de conectar de novo.
  const readOnly = new GraphConnector(delegatedSource({ graph: { ...source.graph, account: { ...source.graph.account, scopes: READ_ONLY } }, secrets: { refreshToken: refreshTokenFor('ana@contoso.com', READ_ONLY) } }), { endpoints: mocks.endpoints });
  const denied = await readOnly.deleteMessages({ address: 'ana@contoso.com' }, ['a2'], 'permanent');
  assert.match(denied.get('a2').error, /conecte a conta de novo com "Permitir excluir"/);
  const writer = new GraphConnector(delegatedSource(), { endpoints: mocks.endpoints });
  const ok = await writer.deleteMessages({ address: 'rh@contoso.com' }, ['r2'], 'trash');
  assert.deepEqual(ok.get('r2'), { ok: true });
  assert.ok(graph.users[1].messages.trash.some((m) => m.id === 'r2'), 'movida para Itens Excluídos da caixa compartilhada');

  // Autorização vencida (90 dias sem uso): pede para conectar a conta de novo — uma vez só, sem
  // repetir o pedido à Microsoft para cada caixa.
  const before = graph.tokenRequests.length;
  const revoked = await runMail([delegatedSource({ secrets: { refreshToken: 'rt-desconhecido' }, mailboxes: [{ address: 'ana@contoso.com' }, { address: 'rh@contoso.com' }, { address: 'bia@contoso.com' }] })]);
  assert.equal(revoked.records.length, 0);
  assert.equal(graph.tokenRequests.length - before, 1, 'um pedido de token só');
  assert.equal(revoked.errors.length, 2);
  assert.match(revoked.errors[0].message, /expirou por falta de uso.*Conectar conta/);
  assert.match(revoked.errors[1].message, /^2 caixas da conexão não foram analisadas: A autorização da conta conectada expirou/);
});

test('IMAP com OAuth da Microsoft: conta conectada e aplicativo (XOAUTH2), só em servidor da Microsoft', async () => {
  const oauthLogins = [];
  const inbox = () => ({ INBOX: [{ raw: mail('Planilha', 'dados confidencial'), date: new Date('2026-09-20T10:00:00Z') }] });
  const server = await startFakeImap(
    { 'ana@contoso.com': { password: 'nao-usada', folders: inbox() }, 'rh@contoso.com': { password: 'x', folders: inbox() } },
    {
      oauthLogins,
      // Como o Exchange Online: o token vale para a caixa da conta e para as que ela tem Acesso Total.
      oauth: (login, token) => {
        const who = imapTokenUser(token);
        return who === '*' || who === login || (login === 'rh@contoso.com' && who === 'ana@contoso.com');
      },
    },
  );
  try {
    const endpoints = { ...mocks.endpoints, microsoftImap: { hosts: ['127.0.0.1'], insecure: true } };
    const source = {
      id: 'src-imap',
      name: 'IMAP da Microsoft',
      type: 'imap',
      scope: 'list',
      mailboxes: [{ address: 'ana@contoso.com' }, { address: 'rh@contoso.com' }],
      excludeMailboxes: [],
      excludeFolders: [],
      imap: { host: '127.0.0.1', port: server.port, security: 'none', auth: 'oauth' },
      graph: { tenantId: TENANT, clientId: CLIENT, auth: 'delegated', account: { username: 'ana@contoso.com', address: 'ana@contoso.com', scopes: [IMAP_SCOPE], grantId: 'grant-imap' } },
      secrets: { refreshToken: refreshTokenFor('ana@contoso.com', [IMAP_SCOPE]) },
    };
    const delegated = await runMail([source], endpoints);
    assert.deepEqual(delegated.errors, []);
    assert.deepEqual(delegated.records.map((r) => r.mailbox).sort(), ['ana@contoso.com', 'rh@contoso.com']);
    assert.deepEqual(oauthLogins, ['ana@contoso.com', 'rh@contoso.com'], 'login por XOAUTH2 (sem senha)');
    assert.equal(delegated.credentials.length, 1, 'o token renovado vai para o servidor gravar');
    const imapRequest = graph.tokenRequests.filter((r) => r.grant === 'refresh_token').at(-1);
    assert.match(imapRequest.scope, /IMAP\.AccessAsUser\.All/);

    // Aplicativo (segredo do cliente): token do recurso outlook.office365.com.
    const app = await runMail([{ ...source, graph: { tenantId: TENANT, clientId: CLIENT, auth: 'secret' }, secrets: { clientSecret: SECRET } }], endpoints);
    assert.equal(app.records.length, 2);
    assert.ok(graph.tokenRequests.some((r) => r.grant === 'client_credentials' && r.scope === 'https://outlook.office365.com/.default'));

    // Fora dos servidores da Microsoft o token não é enviado.
    const elsewhere = await runMail([source], mocks.endpoints);
    assert.equal(elsewhere.records.length, 0);
    assert.match(elsewhere.errors[0].message, /só é usado nos servidores IMAP da Microsoft/);
    assert.equal(oauthLogins.length, 4, 'nenhum login novo');

    // Conta sem acesso à caixa: a mensagem explica o que conferir.
    const denied = await runMail([{ ...source, mailboxes: [{ address: 'rh@contoso.com' }], secrets: { refreshToken: refreshTokenFor('bia@contoso.com', [IMAP_SCOPE]) } }], endpoints);
    assert.match(denied.errors[0].message, /recusou o login OAuth/);
  } finally {
    await server.close();
  }
});

test('IMAP com OAuth: a sessão que cai (token vencido) continua numa nova, sem perder nem repetir mensagens', async () => {
  const oauthLogins = [];
  // Maiores que o limite da análise (1 MB): baixadas uma a uma, cada uma num FETCH.
  const big = (n) => ({ raw: mail(`Grande ${n}`, `confidencial ${'x'.repeat(1100000)}`), date: new Date('2026-09-20T10:00:00Z') });
  const small = (n) => ({ raw: mail(`Pequena ${n}`, 'confidencial'), date: new Date('2026-09-20T10:00:00Z') });
  // 1 FETCH para listar cada pasta + 1 por mensagem grande (baixada sozinha) ou por lote de pequenas:
  // com 3 FETCH por sessão, a sessão cai três vezes no meio da leitura.
  const server = await startFakeImap(
    { 'ana@contoso.com': { password: 'x', folders: { INBOX: [big(1), big(2), big(3), big(4)], Arquivo: [small(1), small(2)] } } },
    { oauthLogins, dropAfterFetches: 3, oauth: (login, token) => imapTokenUser(token) === login },
  );
  const lifetime = graph.tokenLifetime;
  try {
    graph.tokenLifetime = 20 * 60; // 20 minutos: menos que o mínimo de uma sessão (30), então cada sessão pede outro
    const endpoints = { ...mocks.endpoints, microsoftImap: { hosts: ['127.0.0.1'], insecure: true } };
    const refreshes = () => (graph.tokenRequests || []).filter((r) => r.grant === 'refresh_token').length;
    const before = refreshes();
    const messages = [];
    const scanner = new MailScanner(
      {
        sources: [
          {
            id: 'src-queda',
            name: 'IMAP que cai',
            type: 'imap',
            scope: 'list',
            mailboxes: [{ address: 'ana@contoso.com' }],
            excludeMailboxes: [],
            excludeFolders: [],
            imap: { host: '127.0.0.1', port: server.port, security: 'none', auth: 'oauth' },
            graph: { tenantId: TENANT, clientId: CLIENT, auth: 'delegated', account: { username: 'ana@contoso.com', address: 'ana@contoso.com', scopes: [IMAP_SCOPE], grantId: 'g-queda' } },
            secrets: { refreshToken: refreshTokenFor('ana@contoso.com', [IMAP_SCOPE]) },
          },
        ],
        terms: TERMS,
        options: { maxMessageSizeMB: 1 },
        endpoints,
      },
      (m) => messages.push(m),
    );
    const stats = await scanner.run();
    const records = messages.filter((m) => m.type === 'results').flatMap((m) => m.records);
    const errors = messages.filter((m) => m.type === 'errors').flatMap((m) => m.items);
    assert.deepEqual(errors, []);
    assert.deepEqual(records.map((r) => r.subject).sort(), ['Grande 1', 'Grande 2', 'Grande 3', 'Grande 4', 'Pequena 1', 'Pequena 2'], 'todas, uma vez cada');
    assert.equal(stats.messagesSeen, 6);
    assert.ok(oauthLogins.length >= 3, `novas sessões depois das quedas (${oauthLogins.length})`);
    assert.equal(refreshes() - before, oauthLogins.length, 'cada sessão com um token novo (o guardado vale menos de 30 min)');
  } finally {
    graph.tokenLifetime = lifetime;
    await server.close();
  }
});

/** Conexão IMAP com OAuth (conta conectada) no servidor de teste dado. */
const imapOAuthSource = (port) => ({
  id: 'src-imap-sessoes',
  name: 'IMAP sessões',
  type: 'imap',
  scope: 'list',
  mailboxes: [{ address: 'ana@contoso.com' }],
  excludeMailboxes: [],
  excludeFolders: [],
  imap: { host: '127.0.0.1', port, security: 'none', auth: 'oauth' },
  graph: { tenantId: TENANT, clientId: CLIENT, auth: 'delegated', account: { username: 'ana@contoso.com', address: 'ana@contoso.com', scopes: [IMAP_SCOPE], grantId: 'g-sessoes' } },
  secrets: { refreshToken: refreshTokenFor('ana@contoso.com', [IMAP_SCOPE]) },
});
const imapEndpoints = () => ({ ...mocks.endpoints, microsoftImap: { hosts: ['127.0.0.1'], insecure: true } });
const bigMessage = (n) => ({ raw: mail(`Grande ${n}`, `confidencial ${'x'.repeat(1100000)}`), date: new Date('2026-09-20T10:00:00Z') });
const smallMessage = (n) => ({ raw: mail(`Pequena ${n}`, 'confidencial'), date: new Date('2026-09-20T10:00:00Z') });

test('IMAP com OAuth: só desiste depois de quedas seguidas sem leitura; reconexão recusada é tentada de novo', async () => {
  const oauthLogins = [];
  let attempts = 0;
  // Cada sessão lê uma mensagem grande e cai (2 FETCH por sessão): 8 quedas, todas com leitura.
  const server = await startFakeImap(
    { 'ana@contoso.com': { password: 'x', folders: { INBOX: [1, 2, 3, 4, 5, 6, 7, 8].map(bigMessage), Arquivo: [smallMessage(1)] } } },
    {
      oauthLogins,
      dropAfterFetches: 2,
      // A 3ª tentativa de login é recusada uma vez (ex.: falha momentânea do Exchange Online).
      oauth: (login, token) => ++attempts !== 3 && imapTokenUser(token) === login,
    },
  );
  try {
    // Limite de 1 MB: cada mensagem grande é baixada sozinha (um FETCH).
    const { records, errors } = await runMail([imapOAuthSource(server.port)], imapEndpoints(), { maxMessageSizeMB: 1 });
    assert.deepEqual(errors, []);
    assert.equal(records.length, 9, 'todas, uma vez cada');
    assert.equal(new Set(records.map((r) => r.subject)).size, 9);
    assert.ok(oauthLogins.length >= 9, `uma sessão por mensagem grande (${oauthLogins.length})`);
  } finally {
    await server.close();
  }
  // Quedas seguidas sem ler nada: desiste da caixa, com a explicação.
  const stuck = await startFakeImap({ 'ana@contoso.com': { password: 'x', folders: { INBOX: [smallMessage(1)] } } }, { dropAfterFetches: 1, oauth: (login, token) => imapTokenUser(token) === login });
  try {
    const started = Date.now();
    const { records, errors } = await runMail([imapOAuthSource(stuck.port)], imapEndpoints());
    assert.equal(records.length, 0);
    assert.match(errors.at(-1).message, /caiu \(ou não pôde ser reaberta\) 5 vezes seguidas sem ler nenhuma mensagem/);
    assert.ok(Date.now() - started > 5000, 'com pausas entre as tentativas');
  } finally {
    await stuck.close();
  }
});

test('IMAP com OAuth: troca a sessão antes de o token vencer, e não fica trocando com tokens curtos', async () => {
  const oauthLogins = [];
  const server = await startFakeImap(
    { 'ana@contoso.com': { password: 'x', folders: { INBOX: [smallMessage(1), smallMessage(2)], Arquivo: [smallMessage(3)] } } },
    { oauthLogins, oauth: (login, token) => imapTokenUser(token) === login },
  );
  const lifetime = graph.tokenLifetime;
  try {
    // O 1º token vale 4 minutos (menos que a margem de 5): a sessão é trocada logo, com um token novo.
    graph.tokenLifetimes = [240];
    const renewed = await runMail([imapOAuthSource(server.port)], imapEndpoints());
    assert.deepEqual(renewed.errors, []);
    assert.equal(renewed.records.length, 3);
    assert.equal(oauthLogins.length, 2, 'uma troca planejada, sem queda');
    // Todos os tokens valem 4 minutos: trocar não adianta, e a leitura segue na mesma sessão.
    oauthLogins.length = 0;
    graph.tokenLifetime = 240;
    const short = await runMail([imapOAuthSource(server.port)], imapEndpoints());
    assert.deepEqual(short.errors, []);
    assert.equal(short.records.length, 3);
    assert.equal(oauthLogins.length, 2, 'uma troca só (o novo token também é curto)');
  } finally {
    graph.tokenLifetimes = null;
    graph.tokenLifetime = lifetime;
    await server.close();
  }
});

test('IMAP: a leitura de uma pasta para entre os lotes para trocar a sessão e continua de onde parou', async () => {
  const messages = Array.from({ length: 120 }, (_, i) => smallMessage(i + 1));
  const server = await startFakeImap({ 'ana@contoso.com': { password: 'p', folders: { INBOX: messages } } });
  const { ImapConnector } = await import('../src/mail/imap.js');
  const connector = new ImapConnector({ imap: { host: '127.0.0.1', port: server.port, security: 'none' }, secrets: { defaultPassword: 'p' }, mailboxes: [] });
  const folder = { path: 'INBOX', display: 'INBOX', all: false, inTrash: false };
  const state = { validity: null, delivered: new Set() };
  const options = { gmail: false, since: null, before: null, headersOnly: false, maxBytes: 50 * 1048576 };
  try {
    let session = await connector.connect({ address: 'ana@contoso.com' });
    const first = [];
    let batches = 0;
    await assert.rejects(async () => {
      // Depois do 1º lote (50 mensagens), o token "está para vencer".
      for await (const item of connector.folderMessages(session.client, folder, state, { ...options, renewDue: () => batches++ > 0 })) first.push(item.id);
    }, (err) => err.renewSession === true);
    assert.equal(first.length, 50);
    await session.dispose();
    session = await connector.connect({ address: 'ana@contoso.com' });
    const rest = [];
    for await (const item of connector.folderMessages(session.client, folder, state, options)) rest.push(item.id);
    await session.dispose();
    assert.equal(rest.length, 70, 'as que faltavam');
    assert.equal(new Set([...first, ...rest]).size, 120, 'sem repetir');
  } finally {
    await server.close();
  }
});

test('token: renovado quando vale menos que o pedido; certificado vencido nem chega a ser enviado', async () => {
  const lifetime = graph.tokenLifetime;
  try {
    graph.tokenLifetime = 20 * 60;
    const auth = new MicrosoftAuth({ tenantId: TENANT, clientId: CLIENT, auth: 'secret' }, { clientSecret: SECRET }, { login: mocks.endpoints.graphLogin });
    const count = () => (graph.tokenRequests || []).filter((r) => r.grant === 'client_credentials').length;
    const before = count();
    await auth.token('graph');
    await auth.token('graph');
    assert.equal(count() - before, 1, 'guardado enquanto vale mais de 2 minutos');
    await auth.token('graph', { minValidityMs: 30 * 60 * 1000 });
    assert.equal(count() - before, 2, 'renovado: vale menos de 30 minutos');
    assert.ok(auth.expiresAt('graph') > Date.now() + 19 * 60 * 1000);
  } finally {
    graph.tokenLifetime = lifetime;
  }
  const expired = { ...cert.certificate, notAfter: '2025-01-31T12:00:00.000Z' };
  const old = new MicrosoftAuth({ tenantId: TENANT, clientId: CLIENT, auth: 'certificate', certificate: expired }, { certificateKey: cert.privateKeyPem }, { login: mocks.endpoints.graphLogin });
  const before = (graph.tokenRequests || []).length;
  await assert.rejects(old.token('graph'), /certificado do aplicativo venceu em 31\/01\/2025: na conexão, gere um novo certificado/);
  await assert.rejects(old.token('graph'), /venceu em/);
  assert.equal((graph.tokenRequests || []).length, before, 'nenhum pedido com o certificado vencido');
});

// ---------------------------------------------------------------------------------------------
// API: entrada da conta, certificado, cadastro, teste, análise e exclusão

async function startApp() {
  const endpoints = { ...mocks.endpoints, microsoftImap: { hosts: ['127.0.0.1'], insecure: true } };
  const store = await new Store(path.join(root, `data-${crypto.randomUUID()}`)).init();
  const manager = new ScanManager(store, { mailEndpoints: endpoints });
  const app = createApp({ store, manager, config: { authUser: '', authPassword: '', mailEndpoints: endpoints } });
  const srv = await new Promise((resolve) => {
    const x = app.listen(0, '127.0.0.1', () => resolve(x));
  });
  const base = `http://127.0.0.1:${srv.address().port}`;
  const api = async (method, url, body) => {
    const res = await fetch(base + url, { method, headers: { 'Content-Type': 'application/json', 'X-CLEAN': '1' }, body: body === undefined ? undefined : JSON.stringify(body) });
    const type = res.headers.get('content-type') || '';
    return { status: res.status, data: type.includes('json') ? await res.json() : null };
  };
  const wait = async (id) => {
    let scan;
    for (let i = 0; i < 300; i++) {
      scan = (await api('GET', `/api/scans/${id}`)).data;
      if (!['queued', 'running'].includes(scan.status)) return scan;
      await new Promise((r) => setTimeout(r, 100));
    }
    return scan;
  };
  return { store, api, wait, close: async () => (srv.close(), await manager.shutdown()) };
}

/** Entrada pelo código de dispositivo até o fim (o simulador responde "pendente" na 1ª consulta). */
async function signIn(api, body, { onPending } = {}) {
  const started = await api('POST', '/api/mail-sources/oauth/device', body);
  assert.equal(started.status, 201, JSON.stringify(started.data));
  let state;
  for (let i = 0; i < 80; i++) {
    state = (await api('POST', '/api/mail-sources/oauth/device/status', { flowId: started.data.flowId })).data;
    if (state.status !== 'pending') break;
    onPending?.(state);
    await new Promise((r) => setTimeout(r, 150));
  }
  return { ...started.data, state };
}

test('API: conta conectada pelo código, token cifrado, teste, análise e exclusão gravam o token renovado', async () => {
  const app = await startApp();
  try {
    graph.delegated.pendingPolls = 1;
    const flow = await signIn(app.api, { type: 'graph', graph: { tenantId: TENANT, clientId: CLIENT }, allowDelete: true });
    assert.equal(flow.userCode, 'ABCD-EFGH');
    assert.equal(flow.verificationUri, 'https://microsoft.com/devicelogin');
    assert.equal(flow.state.status, 'connected', JSON.stringify(flow.state));
    assert.equal(flow.state.account.address, 'ana@contoso.com');
    assert.equal(flow.state.account.name, 'Ana Souza');
    assert.equal(flow.state.account.canDelete, true);
    const asked = graph.tokenRequests.find((r) => r.endpoint === 'devicecode').scope.split(' ');
    assert.deepEqual(asked, [...READ_WRITE, 'openid', 'profile', 'offline_access'], 'escrita pedida porque a exclusão está marcada');
    assert.ok(!JSON.stringify(flow).includes('rt-'), 'o token não vai para o navegador');

    const created = await app.api('POST', '/api/mail-sources', {
      name: 'Conta da Ana',
      type: 'graph',
      scope: 'all',
      mailboxes: '',
      allowDelete: true,
      graph: { tenantId: TENANT, clientId: CLIENT, auth: 'delegated', signIn: flow.flowId },
    });
    assert.equal(created.status, 201, JSON.stringify(created.data));
    const src = created.data;
    assert.equal(src.scope, 'list', 'com a conta conectada, as caixas são as da lista');
    assert.deepEqual(src.mailboxes.map((m) => m.address), ['ana@contoso.com'], 'sem caixas informadas: a da conta');
    assert.equal(src.graph.auth, 'delegated');
    assert.equal(src.graph.hasRefreshToken, true);
    assert.equal(src.graph.account.username, 'ana@contoso.com');
    assert.equal(src.graph.account.grantId, undefined);
    assert.ok(!JSON.stringify(src).includes('rt-') && !JSON.stringify(src).includes('enc:v1'));
    await app.store.saveNow();
    const db = fs.readFileSync(path.join(app.store.dataDir, 'db.json'), 'utf8');
    assert.ok(!db.includes('"rt-') && db.includes('enc:v1:'), 'token gravado cifrado');
    const reuse = await app.api('POST', '/api/mail-sources', { name: 'Outra', type: 'graph', graph: { tenantId: TENANT, clientId: CLIENT, auth: 'delegated', signIn: flow.flowId } });
    assert.equal(reuse.status, 400, 'a entrada é usada uma vez');

    const stored = () => app.store.openMailSecrets(app.store.getMailSource(src.id)).refreshToken;
    const tested = await app.api('POST', '/api/mail-sources/test', { ...src, graph: { tenantId: TENANT, clientId: CLIENT, auth: 'delegated' }, mailboxes: 'ana@contoso.com\nrh@contoso.com' });
    assert.equal(tested.data.ok, true, JSON.stringify(tested.data));
    assert.ok(tested.data.details.some((d) => /Conta conectada: Ana Souza — ana@contoso\.com/.test(d)));
    assert.ok(tested.data.details.some((d) => /rh@contoso\.com: \d+ mensagem/.test(d)));
    assert.equal(stored(), graph.delegated.issued.at(-1), 'o token renovado no teste foi gravado');

    // Caixa compartilhada incluída; análise em segundo plano (thread) grava o token renovado.
    const updated = await app.api('PUT', `/api/mail-sources/${src.id}`, { ...src, graph: { tenantId: TENANT, clientId: CLIENT, auth: 'delegated' }, mailboxes: 'ana@contoso.com\nrh@contoso.com' });
    assert.equal(updated.status, 200, JSON.stringify(updated.data));
    const list = await app.api('POST', '/api/lists', { name: 'Sigilo', terms: [{ type: 'text', value: 'confidencial' }] });
    const before = stored();
    const scan = await app.api('POST', '/api/scans', { kind: 'mail', sourceIds: [src.id], listIds: [list.data.id] });
    assert.equal(scan.status, 201);
    const done = await app.wait(scan.data.id);
    assert.equal(done.status, 'completed', JSON.stringify(done));
    assert.equal(done.stats.messagesMatched, 2);
    assert.notEqual(stored(), before);
    assert.equal(stored(), graph.delegated.issued.at(-1), 'o token renovado na análise foi gravado');

    const results = await app.api('GET', `/api/scans/${scan.data.id}/results`);
    const rh = results.data.items.find((r) => r.mailbox === 'rh@contoso.com');
    const removed = await app.api('POST', `/api/scans/${scan.data.id}/results/${rh.id}/delete`, { confirm: true, method: 'permanent' });
    assert.equal(removed.status, 200, JSON.stringify(removed.data));
    assert.equal(removed.data.deletion.status, 'deleted');
    assert.equal(stored(), graph.delegated.issued.at(-1), 'o token renovado na exclusão foi gravado');

    // Uma renovação de uma entrada antiga não substitui a da conta conectada depois.
    assert.equal(app.store.saveRefreshToken(src.id, 'outra-entrada', 'rt-velho'), false);
    assert.equal(stored(), graph.delegated.issued.at(-1));

    // Trocar o aplicativo exige conectar de novo; a exclusão exige uma entrada com escrita.
    const otherApp = await app.api('PUT', `/api/mail-sources/${src.id}`, { ...src, graph: { tenantId: TENANT, clientId: '99999999-2222-3333-4444-555555555555', auth: 'delegated' } });
    assert.equal(otherApp.status, 400);
    assert.match(otherApp.data.error, /conecte a conta novamente/);
    const readOnly = await signIn(app.api, { id: src.id, type: 'graph', graph: { tenantId: TENANT, clientId: CLIENT }, allowDelete: false });
    assert.equal(readOnly.state.account.canDelete, false);
    const noWrite = await app.api('PUT', `/api/mail-sources/${src.id}`, { ...src, allowDelete: true, graph: { tenantId: TENANT, clientId: CLIENT, auth: 'delegated', signIn: readOnly.flowId } });
    assert.equal(noWrite.status, 400);
    assert.match(noWrite.data.error, /autorizou somente a leitura/);
    const testedReadOnly = await app.api('POST', '/api/mail-sources/test', { ...src, allowDelete: true, graph: { tenantId: TENANT, clientId: CLIENT, auth: 'delegated', signIn: readOnly.flowId } });
    assert.equal(testedReadOnly.data.ok, false, 'o teste avisa o que o salvar recusaria');
    assert.match(testedReadOnly.data.message, /leitura das caixas funciona, mas a exclusão está permitida/);

    // Autorização revogada: o teste explica o que fazer.
    graph.delegated.revoked = true;
    const revoked = await app.api('POST', '/api/mail-sources/test', { ...src, graph: { tenantId: TENANT, clientId: CLIENT, auth: 'delegated' } });
    assert.equal(revoked.data.ok, false);
    assert.match(revoked.data.message, /revogada.*Conectar conta/);
  } finally {
    graph.delegated.revoked = false;
    await app.close();
  }
});

test('API: entrada recusada pelo registro do aplicativo e locatário genérico só para a conta conectada', async () => {
  const app = await startApp();
  try {
    graph.delegated.publicClientDisabled = true;
    graph.delegated.pendingPolls = 0;
    const flow = await signIn(app.api, { type: 'graph', graph: { tenantId: TENANT, clientId: CLIENT } });
    assert.equal(flow.state.status, 'failed');
    assert.match(flow.state.error, /Permitir fluxos de clientes públicos/);
    graph.delegated.publicClientDisabled = false;

    const generic = await app.api('POST', '/api/mail-sources', { name: 'X', type: 'graph', graph: { tenantId: 'organizations', clientId: CLIENT, auth: 'secret', clientSecret: SECRET } });
    assert.equal(generic.status, 400);
    assert.match(generic.data.error, /vale só para a conta conectada/);
    const personal = await signIn(app.api, { type: 'graph', graph: { tenantId: 'consumers', clientId: CLIENT } });
    assert.equal(personal.state.status, 'connected');
    const scopes = graph.tokenRequests.filter((r) => r.endpoint === 'devicecode').at(-1).scope;
    assert.ok(!scopes.includes('.Shared'), 'contas pessoais não têm caixas compartilhadas');
    const unknown = await app.api('POST', '/api/mail-sources/oauth/device/status', { flowId: 'nao-existe' });
    assert.equal(unknown.data.status, 'failed');

    // Falha passageira da Microsoft (503) numa consulta: a entrada continua, e termina.
    graph.delegated.flakyPolls = 2;
    graph.delegated.pendingPolls = 1;
    const notices = [];
    const flaky = await signIn(app.api, { type: 'graph', graph: { tenantId: TENANT, clientId: CLIENT } }, { onPending: (state) => state.notice && notices.push(state.notice) });
    assert.equal(flaky.state.status, 'connected', JSON.stringify(flaky.state));
    assert.match(notices[0], /Falha momentânea ao consultar a Microsoft/);

    // Entrada ainda em andamento, cancelada ou começada em outra conexão: não é salva.
    graph.delegated.pendingPolls = 50;
    const started = await app.api('POST', '/api/mail-sources/oauth/device', { type: 'graph', graph: { tenantId: TENANT, clientId: CLIENT } });
    const body = { name: 'X', type: 'graph', graph: { tenantId: TENANT, clientId: CLIENT, auth: 'delegated', signIn: started.data.flowId } };
    const unfinished = await app.api('POST', '/api/mail-sources', body);
    assert.equal(unfinished.status, 400);
    assert.match(unfinished.data.error, /ainda não terminou/);
    assert.equal((await app.api('POST', '/api/mail-sources/oauth/device/cancel', { flowId: started.data.flowId })).status, 204);
    const cancelled = await app.api('POST', '/api/mail-sources', body);
    assert.match(cancelled.data.error, /não está mais disponível/);
    graph.delegated.pendingPolls = 0;
    const source = await app.api('POST', '/api/mail-sources', { name: 'Segredo', type: 'graph', scope: 'all', graph: { tenantId: TENANT, clientId: CLIENT, auth: 'secret', clientSecret: SECRET } });
    const forOther = await signIn(app.api, { id: source.data.id, type: 'graph', graph: { tenantId: TENANT, clientId: CLIENT } });
    assert.equal(forOther.state.status, 'connected');
    const elsewhere = await app.api('POST', '/api/mail-sources', { name: 'Outra', type: 'graph', graph: { tenantId: TENANT, clientId: CLIENT, auth: 'delegated', signIn: forOther.flowId } });
    assert.equal(elsewhere.status, 400);
    assert.match(elsewhere.data.error, /feita em outra conexão/);
    const here = await app.api('PUT', `/api/mail-sources/${source.data.id}`, { name: 'Segredo', type: 'graph', graph: { tenantId: TENANT, clientId: CLIENT, auth: 'delegated', signIn: forOther.flowId } });
    assert.equal(here.status, 200, JSON.stringify(here.data));
  } finally {
    graph.delegated.publicClientDisabled = false;
    graph.delegated.pendingPolls = 1;
    graph.delegated.flakyPolls = 0;
    await app.close();
  }
});

test('API: certificado gerado no CLEAN (chave só no servidor), teste, cadastro e repositório ligado', async () => {
  const app = await startApp();
  try {
    const longName = 'Caixas do departamento financeiro e contábil da matriz em São Paulo';
    const long = await app.api('POST', '/api/mail-sources/certificate', { name: longName });
    assert.equal(long.status, 201, 'nome longo: o nome do certificado é cortado');
    assert.equal(long.data.certificate.subject, `CN=CLEAN - ${longName.slice(0, 56).trim()}`, 'o nome do certificado (CN) tem no máximo 64 caracteres');
    const generated = await app.api('POST', '/api/mail-sources/certificate', { name: 'E-mail', replaces: long.data.certificateId });
    assert.equal(generated.status, 201);
    const replaced = await app.api('POST', '/api/mail-sources', { name: 'Y', type: 'graph', scope: 'all', graph: { tenantId: TENANT, clientId: CLIENT, auth: 'certificate', certificateId: long.data.certificateId } });
    assert.equal(replaced.status, 400, 'o certificado substituído saiu da memória');
    const { certificateId, certificate } = generated.data;
    assert.match(certificate.thumbprint, /^[0-9A-F]{40}$/);
    assert.equal(certificate.subject, 'CN=CLEAN - E-mail');
    assert.ok(certificate.pem.startsWith('-----BEGIN CERTIFICATE-----'));
    assert.ok(!JSON.stringify(generated.data).includes('PRIVATE KEY'), 'a chave privada não sai do servidor');
    graph.certificatePem = certificate.pem; // "enviado" ao registro do aplicativo

    const body = { name: 'Certificado', type: 'graph', scope: 'all', graph: { tenantId: TENANT, clientId: CLIENT, auth: 'certificate', certificateId } };
    const tested = await app.api('POST', '/api/mail-sources/test', body);
    assert.equal(tested.data.ok, true, JSON.stringify(tested.data));
    assert.ok(tested.data.details.some((d) => /Entra ID \(certificado\): OK/.test(d)));
    const created = await app.api('POST', '/api/mail-sources', body);
    assert.equal(created.status, 201, JSON.stringify(created.data));
    assert.equal(created.data.graph.certificate.thumbprint, certificate.thumbprint);
    assert.equal(created.data.graph.hasCertificateKey, true);
    await app.store.saveNow();
    const db = fs.readFileSync(path.join(app.store.dataDir, 'db.json'), 'utf8');
    assert.ok(!db.includes('PRIVATE KEY'), 'a chave privada é gravada cifrada');
    const kept = await app.api('PUT', `/api/mail-sources/${created.data.id}`, { ...body, certificateId: undefined, graph: { tenantId: TENANT, clientId: CLIENT, auth: 'certificate' } });
    assert.equal(kept.status, 200, 'sem certificado novo, mantém o salvo');
    assert.equal(kept.data.graph.certificate.thumbprint, certificate.thumbprint);

    // Repositório do OneDrive ligado à conexão: usa o mesmo certificado.
    const repo = await app.api('POST', '/api/repositories', { name: 'OneDrive', type: 'onedrive', scope: 'all', credentialsFrom: created.data.id });
    assert.equal(repo.status, 201, JSON.stringify(repo.data));
    assert.equal(repo.data.graph.auth, 'certificate');
    assert.equal(repo.data.graph.hasCertificateKey, true);
    assert.ok(app.store.openRepositorySecrets(app.store.getRepository(repo.data.id)).certificateKey.includes('PRIVATE KEY'));

    // A conexão passa para a conta conectada: o repositório deixa de estar ligado (permissões só de e-mail).
    const flow = await signIn(app.api, { id: created.data.id, type: 'graph', graph: { tenantId: TENANT, clientId: CLIENT } });
    const delegated = await app.api('PUT', `/api/mail-sources/${created.data.id}`, { ...body, graph: { tenantId: TENANT, clientId: CLIENT, auth: 'delegated', signIn: flow.flowId } });
    assert.equal(delegated.status, 200, JSON.stringify(delegated.data));
    assert.equal(app.store.getRepository(repo.data.id).credentialsFrom, null);
    // O repositório desligado continua com o certificado e pode ser editado sem informar um segredo.
    const renamed = await app.api('PUT', `/api/repositories/${repo.data.id}`, { name: 'OneDrive (certificado)', type: 'onedrive', scope: 'all', graph: { tenantId: TENANT, clientId: CLIENT, clientSecret: '' } });
    assert.equal(renamed.status, 200, JSON.stringify(renamed.data));
    assert.equal(renamed.data.graph.auth, 'certificate');
    assert.equal(renamed.data.graph.hasCertificateKey, true);
    const toSecret = await app.api('PUT', `/api/repositories/${repo.data.id}`, { name: 'OneDrive', type: 'onedrive', scope: 'all', graph: { tenantId: TENANT, clientId: CLIENT, clientSecret: SECRET } });
    assert.equal(toSecret.data.graph.auth, undefined, 'com um segredo informado, passa a usar o segredo');
    assert.equal(toSecret.data.graph.hasClientSecret, true);
    const link = await app.api('POST', '/api/repositories', { name: 'SharePoint', type: 'sharepoint', scope: 'all', credentialsFrom: created.data.id });
    assert.equal(link.status, 400);
    assert.match(link.data.error, /usa uma conta conectada/);

    // Importação de um certificado existente (PEM com a chave) e erros.
    const imported = await app.api('POST', '/api/mail-sources/certificate', { pem: `${cert.certificate.pem}${cert.privateKeyPem}` });
    assert.equal(imported.status, 201);
    assert.equal(imported.data.certificate.thumbprint, cert.certificate.thumbprint);
    const invalid = await app.api('POST', '/api/mail-sources/certificate', { pem: cert.certificate.pem });
    assert.equal(invalid.status, 400);
    const expired = await app.api('POST', '/api/mail-sources', { ...body, name: 'Y', graph: { ...body.graph, certificateId: 'nao-existe' } });
    assert.equal(expired.status, 400);
    assert.match(expired.data.error, /gere ou importe o certificado de novo/);
  } finally {
    await app.close();
  }
});

test('API: IMAP com OAuth — só servidores da Microsoft, sem senhas, e a caixa da conta por padrão', async () => {
  const app = await startApp();
  try {
    const flow = await signIn(app.api, { type: 'imap', graph: { tenantId: TENANT, clientId: CLIENT } });
    assert.equal(flow.state.status, 'connected');
    const scopes = graph.tokenRequests.filter((r) => r.endpoint === 'devicecode').at(-1).scope;
    assert.equal(scopes, `${IMAP_SCOPE} openid profile offline_access`);
    const body = { name: 'IMAP OAuth', type: 'imap', imap: { host: 'imap.empresa.com', security: 'tls', auth: 'oauth' }, mailboxes: [], graph: { tenantId: TENANT, clientId: CLIENT, auth: 'delegated', signIn: flow.flowId } };
    const elsewhere = await app.api('POST', '/api/mail-sources', body);
    assert.equal(elsewhere.status, 400);
    assert.match(elsewhere.data.error, /só é usado nos servidores IMAP da Microsoft/);
    const selfSigned = await app.api('POST', '/api/mail-sources', { ...body, imap: { host: 'outlook.office365.com', security: 'tls', allowSelfSigned: true, auth: 'oauth' } });
    assert.equal(selfSigned.status, 400);
    const created = await app.api('POST', '/api/mail-sources', { ...body, imap: { host: 'outlook.office365.com', security: 'tls', auth: 'oauth' } });
    assert.equal(created.status, 201, JSON.stringify(created.data));
    assert.equal(created.data.imap.auth, 'oauth');
    assert.equal(created.data.imap.port, 993);
    assert.deepEqual(created.data.mailboxes.map((m) => m.address), ['ana@contoso.com']);
    assert.equal(created.data.graph.account.canDelete, true, 'o IMAP permite alterar a caixa');
    // Uma conta conectada para o IMAP não serve para o Microsoft Graph (e vice-versa).
    const asGraph = await app.api('POST', '/api/mail-sources', { name: 'Z', type: 'graph', graph: { tenantId: TENANT, clientId: CLIENT, auth: 'delegated', signIn: flow.flowId } });
    assert.equal(asGraph.status, 400);

    // Com OAuth, o login do XOAUTH2 é o e-mail da caixa: um login de senha deixado na lista não é guardado.
    const app2 = await app.api('POST', '/api/mail-sources', {
      name: 'IMAP aplicativo',
      type: 'imap',
      imap: { host: 'outlook.office365.com', security: 'tls', auth: 'oauth' },
      mailboxes: [{ address: 'financeiro@contoso.com', login: 'CONTOSO\\svc\\financeiro' }],
      graph: { tenantId: TENANT, clientId: CLIENT, auth: 'secret', clientSecret: SECRET },
    });
    assert.equal(app2.status, 201, JSON.stringify(app2.data));
    assert.deepEqual(app2.data.mailboxes.map((m) => m.login), [undefined]);
    // O segredo do mesmo aplicativo continua valendo ao passar para o tipo Microsoft 365 (e de volta).
    const asM365 = await app.api('PUT', `/api/mail-sources/${app2.data.id}`, { name: 'Microsoft 365', type: 'graph', scope: 'all', graph: { tenantId: TENANT, clientId: CLIENT, auth: 'secret', clientSecret: '' } });
    assert.equal(asM365.status, 200, JSON.stringify(asM365.data));
    assert.equal(asM365.data.graph.hasClientSecret, true);
    const otherApp = await app.api('PUT', `/api/mail-sources/${app2.data.id}`, { name: 'Microsoft 365', type: 'graph', scope: 'all', graph: { tenantId: TENANT, clientId: '99999999-2222-3333-4444-555555555555', auth: 'secret', clientSecret: '' } });
    assert.equal(otherApp.status, 400, 'outro aplicativo: o segredo precisa ser informado de novo');
  } finally {
    await app.close();
  }
});
