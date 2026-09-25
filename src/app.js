// Aplicação Express: API REST em /api e interface web (arquivos estáticos de /public).
import crypto from 'node:crypto';
import os from 'node:os';
import path from 'node:path';
import express from 'express';
import { PROJECT_ROOT } from './config.js';
import { repositoriesRouter } from './routes/repositories.js';
import { listsRouter } from './routes/lists.js';
import { scansRouter } from './routes/scans.js';
import { HttpError } from './routes/validate.js';
import { PRESETS, VALIDATORS } from './scan/presets.js';
import { DEFAULT_OPTIONS } from './scan/scanner.js';
import { DEFAULT_EXCLUDES } from './scan/walker.js';

const VERSION = '1.0.0';

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

/**
 * Proteção contra CSRF: requisições que alteram dados precisam do cabeçalho X-CLEAN (que um site
 * de terceiros não consegue enviar sem CORS) e, se houver Origin, ele deve ser o próprio servidor.
 */
function csrfGuard(req, res, next) {
  if (['GET', 'HEAD', 'OPTIONS'].includes(req.method)) return next();
  if (req.get('X-CLEAN') !== '1') return next(new HttpError(403, 'Requisição recusada (cabeçalho X-CLEAN ausente).'));
  const origin = req.get('Origin');
  if (origin) {
    let host = null;
    try {
      host = new URL(origin).host;
    } catch {
      // Origin inválido
    }
    if (host !== req.get('Host')) return next(new HttpError(403, 'Requisição de outra origem recusada.'));
  }
  next();
}

export function createApp({ store, manager, config }) {
  const app = express();
  app.disable('x-powered-by');
  app.set('trust proxy', false);
  app.use(securityHeaders);
  if (config.authUser && config.authPassword) app.use(basicAuth(config.authUser, config.authPassword));

  const api = express.Router();
  api.use(express.json({ limit: '10mb' }));
  api.use((req, res, next) => {
    res.setHeader('Cache-Control', 'no-store');
    next();
  });
  api.use(csrfGuard);

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
    });
  });
  api.use('/repositories', repositoriesRouter({ store }));
  api.use('/lists', listsRouter({ store }));
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
