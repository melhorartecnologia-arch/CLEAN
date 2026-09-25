// Aplicação Express: API REST em /api e interface web (arquivos estáticos de /public).
import crypto from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import express from 'express';
import { PROJECT_ROOT } from './config.js';
import { repositoriesRouter } from './routes/repositories.js';
import { listsRouter } from './routes/lists.js';
import { scansRouter } from './routes/scans.js';
import { mailSourcesRouter } from './routes/mail-sources.js';
import { HttpError } from './routes/validate.js';
import { PRESETS, VALIDATORS } from './scan/presets.js';
import { DEFAULT_OPTIONS } from './scan/scanner.js';
import { DEFAULT_EXCLUDES } from './scan/walker.js';
import { MAIL_DEFAULT_OPTIONS } from './mail/scanner.js';
import { MAIL_TYPES } from './mail/connectors.js';

// Versão exibida na interface (lida do package.json).
const VERSION = JSON.parse(fs.readFileSync(path.join(PROJECT_ROOT, 'package.json'), 'utf8')).version;

function sameSecret(a, b) {
  const ha = crypto.createHash('sha256').update(String(a)).digest();
  const hb = crypto.createHash('sha256').update(String(b)).digest();
  return crypto.timingSafeEqual(ha, hb);
}

/** Autenticação HTTP Basic (ativada quando AUTH_USER e AUTH_PASSWORD estão definidos). */
function basicAuth(user, password) {
  return (req, res, next) => {
    const header = req.headers.authorization || '';
    const [scheme, encoded] = header.split(' ');
    if (scheme === 'Basic' && encoded) {
      const decoded = Buffer.from(encoded, 'base64').toString('utf8');
      const sep = decoded.indexOf(':');
      if (sep !== -1 && sameSecret(decoded.slice(0, sep), user) && sameSecret(decoded.slice(sep + 1), password)) return next();
    }
    res.setHeader('WWW-Authenticate', 'Basic realm="CLEAN", charset="UTF-8"');
    res.status(401).send('Autenticação necessária.');
  };
}

function securityHeaders(req, res, next) {
  res.setHeader('X-Content-Type-Options', 'nosniff');
  res.setHeader('X-Frame-Options', 'DENY');
  res.setHeader('Referrer-Policy', 'no-referrer');
  res.setHeader(
    'Content-Security-Policy',
    "default-src 'self'; script-src 'self'; style-src 'self'; img-src 'self' data:; connect-src 'self'; object-src 'none'; base-uri 'none'; frame-ancestors 'none'; form-action 'self'",
  );
  next();
}

/** Nome do host sem a porta ("servidor:3000" -> "servidor", "[::1]:3000" -> "[::1]"). */
export function hostName(value) {
  const host = String(value || '').trim().toLowerCase();
  if (host.startsWith('[')) return host.slice(0, host.indexOf(']') + 1);
  const colon = host.indexOf(':');
  return colon === -1 || host.indexOf(':', colon + 1) !== -1 ? host : host.slice(0, colon);
}

/** localhost, nome e endereços IP desta máquina e os nomes configurados em ALLOWED_HOSTS. */
export function allowedHostNames(extra = []) {
  const names = new Set(['localhost', '127.0.0.1', '[::1]', ...extra]);
  const machine = os.hostname().toLowerCase();
  names.add(machine);
  if (process.env.USERDNSDOMAIN) names.add(`${machine}.${process.env.USERDNSDOMAIN.toLowerCase()}`);
  for (const list of Object.values(os.networkInterfaces())) {
    for (const item of list || []) names.add(item.family === 'IPv6' || item.family === 6 ? `[${item.address.toLowerCase()}]` : item.address);
  }
  return names;
}

/**
 * Proteção contra "DNS rebinding": uma página de outro site que faça seu domínio apontar para este
 * servidor chega com um Host desconhecido e é recusada (inclusive nas leituras).
 */
function hostGuard(allowed) {
  if (allowed.has('*')) return (req, res, next) => next();
  return (req, res, next) => {
    const host = hostName(req.get('Host'));
    if (allowed.has(host)) return next();
    res
      .status(403)
      .type('text/plain; charset=utf-8')
      .send(`Acesso recusado: o endereço "${host}" não está autorizado. Inclua-o em ALLOWED_HOSTS no arquivo .env do CLEAN.`);
  };
}

/**
 * Proteção contra CSRF: requisições que alteram dados precisam do cabeçalho X-CLEAN (que um site
 * de terceiros não consegue enviar sem CORS) e, se houver Origin, ele deve ser um endereço autorizado.
 */
function csrfGuard(allowed) {
  return (req, res, next) => {
    if (['GET', 'HEAD', 'OPTIONS'].includes(req.method)) return next();
    if (req.get('X-CLEAN') !== '1') return next(new HttpError(403, 'Requisição recusada (cabeçalho X-CLEAN ausente).'));
    const origin = req.get('Origin');
    if (origin) {
      let host = null;
      try {
        host = hostName(new URL(origin).host);
      } catch {
        // Origin inválido
      }
      const ok = host && (allowed.has('*') || allowed.has(host) || host === hostName(req.get('Host')));
      if (!ok) return next(new HttpError(403, 'Requisição de outra origem recusada.'));
    }
    next();
  };
}

export function createApp({ store, manager, config }) {
  const app = express();
  const allowed = allowedHostNames(config.allowedHosts || []);
  app.disable('x-powered-by');
  app.set('trust proxy', false);
  app.use(hostGuard(allowed));
  app.use(securityHeaders);
  if (config.authUser && config.authPassword) app.use(basicAuth(config.authUser, config.authPassword));

  const api = express.Router();
  api.use(express.json({ limit: '10mb' }));
  api.use((req, res, next) => {
    res.setHeader('Cache-Control', 'no-store');
    next();
  });
  api.use(csrfGuard(allowed));

  api.get('/info', (req, res) => {
    res.json({
      version: VERSION,
      platform: process.platform,
      hostname: os.hostname(),
      user: os.userInfo().username,
      node: process.version,
      dataDir: store.dataDir,
      auth: Boolean(config.authUser && config.authPassword),
      presets: PRESETS,
      validators: Object.fromEntries(Object.entries(VALIDATORS).map(([k, v]) => [k, v.label])),
      defaults: DEFAULT_OPTIONS,
      defaultExcludes: DEFAULT_EXCLUDES,
      mailDefaults: MAIL_DEFAULT_OPTIONS,
      mailTypes: MAIL_TYPES,
    });
  });
  api.use('/repositories', repositoriesRouter({ store }));
  api.use('/lists', listsRouter({ store }));
  api.use('/mail-sources', mailSourcesRouter({ store, endpoints: config.mailEndpoints }));
  api.use('/scans', scansRouter({ store, manager }));
  api.use((req, res, next) => next(new HttpError(404, 'Rota não encontrada.')));
  // eslint-disable-next-line no-unused-vars
  api.use((err, req, res, next) => {
    const status = err.status || err.statusCode || 500;
    if (status >= 500) console.error('[CLEAN]', err);
    const message = err.type === 'entity.parse.failed' ? 'JSON inválido.' : status >= 500 ? 'Erro interno no servidor.' : err.message;
    res.status(status).json({ error: message });
  });
  app.use('/api', api);

  app.use(express.static(path.join(PROJECT_ROOT, 'public'), { index: 'index.html', maxAge: 0 }));
  return app;
}
