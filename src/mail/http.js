// Requisições HTTP às APIs de e-mail (Microsoft Graph e Gmail): tempo limite, novas tentativas em
// limites de taxa (429) e falhas temporárias (5xx), mensagens de erro legíveis e download com limite
// de tamanho.

export class ApiError extends Error {
  constructor(message, { status = 0, code = '', retryable = false } = {}) {
    super(message);
    this.name = 'ApiError';
    this.status = status;
    this.code = code;
    this.retryable = retryable;
  }
}

const RETRY_STATUS = new Set([408, 429, 500, 502, 503, 504]);

/** Espera ms milissegundos (interrompida se o sinal for abortado). */
export function sleep(ms, signal) {
  return new Promise((resolve, reject) => {
    if (signal?.aborted) return reject(signal.reason);
    const timer = setTimeout(() => {
      signal?.removeEventListener('abort', onAbort);
      resolve();
    }, ms);
    const onAbort = () => {
      clearTimeout(timer);
      reject(signal.reason);
    };
    signal?.addEventListener('abort', onAbort, { once: true });
  });
}

function retryAfter(res, attempt) {
  const header = res?.headers?.get('retry-after');
  if (header) {
    const seconds = Number(header);
    if (Number.isFinite(seconds)) return Math.min(Math.max(seconds, 1), 300) * 1000;
    const date = Date.parse(header);
    if (!Number.isNaN(date)) return Math.min(Math.max(date - Date.now(), 1000), 300000);
  }
  return Math.min(1000 * 2 ** attempt, 60000) + Math.floor(Math.random() * 500);
}

/** Mensagem de erro de Graph ({ error: { code, message } }) ou Google ({ error: { message } } / OAuth). */
function describeError(status, body) {
  let data = null;
  try {
    data = JSON.parse(body);
  } catch {
    // resposta sem JSON
  }
  const err = data?.error;
  if (err && typeof err === 'object') {
    // Google informa o motivo em errors[].reason (ex.: userRateLimitExceeded) ou details[].reason.
    const reason = String(err.errors?.[0]?.reason || err.details?.find?.((d) => d?.reason)?.reason || '');
    return { code: String(err.code || err.status || ''), message: String(err.message || err.status || ''), reason };
  }
  if (typeof err === 'string') return { code: err, message: String(data.error_description || err), reason: '' };
  const text = String(body || '').replace(/\s+/g, ' ').trim().slice(0, 300);
  return { code: '', message: text || `HTTP ${status}`, reason: '' };
}

/** 429/5xx e o 403 de limite de taxa do Google (userRateLimitExceeded, rateLimitExceeded). */
function isRetryable(status, { reason, message }) {
  if (RETRY_STATUS.has(status)) return true;
  const text = `${reason} ${message}`;
  return status === 403 && /rate.?limit|too many requests/i.test(text) && !/daily/i.test(text);
}

function isAbort(err, signal) {
  return signal?.aborted || err?.name === 'AbortError';
}

/**
 * Faz uma requisição com novas tentativas.
 * options: { method, headers, form (objeto → x-www-form-urlencoded), json, signal, timeoutMs,
 *            retries, type: 'json' | 'buffer' | 'text', maxBytes (para 'buffer'), onRetry }
 * Com type 'buffer' retorna { data: Buffer, truncated, size } — o corpo é lido até maxBytes.
 */
export async function request(url, options = {}) {
  const { method = 'GET', form, json, signal, timeoutMs = 120000, retries = 6, type = 'json', maxBytes = Infinity, onRetry } = options;
  const headers = { ...(options.headers || {}) };
  let body;
  if (form) {
    body = new URLSearchParams(form).toString();
    headers['Content-Type'] = 'application/x-www-form-urlencoded';
  } else if (json !== undefined) {
    body = JSON.stringify(json);
    headers['Content-Type'] = 'application/json';
  }
  if (type === 'json') headers.Accept ||= 'application/json';
  for (let attempt = 0; ; attempt++) {
    if (signal?.aborted) throw signal.reason;
    // Tempo limite por inatividade: renovado a cada parte recebida de um download, para que
    // mensagens grandes em redes lentas não sejam interrompidas no meio.
    const controller = new AbortController();
    let timedOut = false;
    let timer = null;
    const arm = () => {
      clearTimeout(timer);
      timer = setTimeout(() => {
        timedOut = true;
        controller.abort(new Error('timeout'));
      }, timeoutMs);
    };
    arm();
    const combined = signal ? AbortSignal.any([signal, controller.signal]) : controller.signal;
    let res;
    try {
      res = await fetch(url, { method, headers, body, signal: combined, redirect: 'follow' });
      if (res.ok) {
        if (type === 'buffer') return await readLimited(res, maxBytes, arm);
        // Respostas JSON também renovam o tempo limite a cada parte (ex.: mensagens do Gmail).
        const text = (await readLimited(res, Infinity, arm)).data.toString('utf8');
        if (type === 'text') return text;
        return text ? JSON.parse(text) : null;
      }
      const text = await res.text().catch(() => '');
      const described = describeError(res.status, text);
      const error = new ApiError(described.message, { status: res.status, code: described.code, retryable: isRetryable(res.status, described) });
      if (!error.retryable || attempt >= retries) throw error;
      const wait = retryAfter(res, attempt);
      onRetry?.({ attempt: attempt + 1, wait, error });
      await sleep(wait, signal);
    } catch (err) {
      if (isAbort(err, signal) && signal?.aborted) throw signal.reason ?? err;
      if (err instanceof ApiError) throw err;
      // Tempo esgotado ou falha de rede: tenta de novo.
      const error = new ApiError(timedOut ? `Tempo esgotado ao acessar ${new URL(url).host}.` : networkMessage(err, url), { code: err?.cause?.code || err?.code || '', retryable: true });
      if (attempt >= retries) throw error;
      const wait = retryAfter(null, attempt);
      onRetry?.({ attempt: attempt + 1, wait, error });
      await sleep(wait, signal);
    } finally {
      clearTimeout(timer);
    }
  }
}

function networkMessage(err, url) {
  const code = err?.cause?.code || err?.code || '';
  const host = new URL(url).host;
  const certificate = `Certificado de ${host} não reconhecido (${code}). Se a rede usa um proxy com inspeção de HTTPS, informe o certificado da empresa na variável NODE_EXTRA_CA_CERTS (veja "Credenciais e rede" no LEIA-ME).`;
  const known = {
    ENOTFOUND: `Endereço ${host} não encontrado (DNS). Verifique a conexão com a internet deste servidor.`,
    ECONNREFUSED: `Conexão recusada por ${host}.`,
    ECONNRESET: `A conexão com ${host} foi interrompida.`,
    ETIMEDOUT: `Tempo esgotado ao conectar em ${host}.`,
    UND_ERR_CONNECT_TIMEOUT: `Tempo esgotado ao conectar em ${host}. Se a rede exige proxy, veja a seção "Proxy" do LEIA-ME.`,
    SELF_SIGNED_CERT_IN_CHAIN: certificate,
    UNABLE_TO_VERIFY_LEAF_SIGNATURE: certificate,
    UNABLE_TO_GET_ISSUER_CERT_LOCALLY: certificate,
    DEPTH_ZERO_SELF_SIGNED_CERT: certificate,
    CERT_HAS_EXPIRED: certificate,
  };
  return known[code] || `Falha de rede ao acessar ${host}: ${err?.cause?.message || err?.message || err}`;
}

/** Lê o corpo da resposta até maxBytes; o restante é descartado (conexão encerrada). */
async function readLimited(res, maxBytes, onChunk = () => {}) {
  const declared = Number(res.headers.get('content-length')) || 0;
  const chunks = [];
  let size = 0;
  let truncated = false;
  const reader = res.body.getReader();
  try {
    for (;;) {
      const { value, done } = await reader.read();
      if (done) break;
      onChunk();
      if (size + value.length > maxBytes) {
        chunks.push(Buffer.from(value.buffer, value.byteOffset, maxBytes - size));
        size = maxBytes;
        truncated = true;
        break;
      }
      chunks.push(Buffer.from(value.buffer, value.byteOffset, value.byteLength));
      size += value.length;
    }
  } finally {
    if (truncated) await reader.cancel().catch(() => {});
    else reader.releaseLock();
  }
  return { data: Buffer.concat(chunks, size), truncated, size: Math.max(declared, size) };
}

/**
 * Executa fn para cada item de um iterador (assíncrono) com até `limit` execuções simultâneas e
 * entrega os resultados conforme ficam prontos. Resultados undefined não são entregues.
 */
export async function* pool(source, limit, fn) {
  const iterator = source[Symbol.asyncIterator] ? source[Symbol.asyncIterator]() : source[Symbol.iterator]();
  const running = new Map();
  let seq = 0;
  let exhausted = false;
  try {
    for (;;) {
      while (!exhausted && running.size < limit) {
        const { value, done } = await iterator.next();
        if (done) {
          exhausted = true;
          break;
        }
        const key = seq++;
        running.set(
          key,
          Promise.resolve()
            .then(() => fn(value))
            .then(
              (result) => ({ key, result }),
              (error) => ({ key, error }),
            ),
        );
      }
      if (running.size === 0) return;
      const { key, result, error } = await Promise.race(running.values());
      running.delete(key);
      if (error) throw error;
      if (result !== undefined) yield result;
    }
  } finally {
    await iterator.return?.();
  }
}
