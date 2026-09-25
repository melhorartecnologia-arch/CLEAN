// Mensagens MIME: e-mails .eml e páginas da web salvas em arquivo único (.mht/.mhtml).
// Decodifica quoted-printable e base64, conjuntos de caracteres e cabeçalhos codificados (RFC 2047).
import { decoderFor, decodeText } from './text.js';
import { htmlToText } from './xml.js';
import { compact, isoDate } from './ooxml.js';

export const MIME_EXTENSIONS = new Set(['.eml', '.mht', '.mhtml', '.mhtm']);

const MAX_DEPTH = 6;

function splitHeaderBody(raw) {
  const m = /\r?\n\r?\n/.exec(raw);
  return m ? { head: raw.slice(0, m.index), body: raw.slice(m.index + m[0].length) } : { head: raw, body: '' };
}

function parseHeaders(head) {
  const headers = {};
  for (const line of head.replace(/\r?\n[ \t]+/g, ' ').split(/\r?\n/)) {
    const idx = line.indexOf(':');
    if (idx <= 0) continue;
    const name = line.slice(0, idx).trim().toLowerCase();
    if (!(name in headers)) headers[name] = line.slice(idx + 1).trim();
  }
  return headers;
}

function decodeQuotedPrintable(text) {
  const s = text.replace(/=\r?\n/g, '');
  const out = Buffer.alloc(s.length);
  let n = 0;
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    if (c === 61 && /^[0-9A-Fa-f]{2}$/.test(s.substr(i + 1, 2))) {
      out[n++] = parseInt(s.substr(i + 1, 2), 16);
      i += 2;
    } else {
      out[n++] = c & 0xff;
    }
  }
  return out.subarray(0, n);
}

/** Parâmetro de cabeçalho (charset, boundary, name, filename), inclusive na forma RFC 2231 (filename*=). */
function param(value, name) {
  const v = value || '';
  const extended = new RegExp(`(?:^|;)\\s*${name}\\*(?:0\\*)?=\\s*([^;]+)`, 'i').exec(v);
  if (extended) {
    const [charset, , encoded = ''] = extended[1].trim().replace(/^"|"$/g, '').split("'");
    try {
      const bytes = Buffer.from(encoded.replace(/%([0-9A-Fa-f]{2})/g, (_, h) => String.fromCharCode(parseInt(h, 16))), 'latin1');
      return decoderFor((charset || 'utf-8').toLowerCase()).decode(bytes);
    } catch {
      return encoded;
    }
  }
  const m = new RegExp(`(?:^|;)\\s*${name}\\s*=\\s*(?:"([^"]*)"|([^;\\s]+))`, 'i').exec(v);
  return m ? (m[1] ?? m[2]) : '';
}

/** Decodifica palavras codificadas (=?utf-8?B?...?= / =?iso-8859-1?Q?...?=) em cabeçalhos. */
export function decodeHeader(value) {
  return String(value || '').replace(/=\?([^?]+)\?([bqBQ])\?([^?]*)\?=(?:\s+(?==\?))?/g, (all, charset, enc, text) => {
    try {
      const bytes = enc.toUpperCase() === 'B' ? Buffer.from(text, 'base64') : decodeQuotedPrintable(text.replace(/_/g, ' '));
      return decoderFor(charset.toLowerCase()).decode(bytes);
    } catch {
      return all;
    }
  });
}

function decodeBody(body, encoding) {
  const cte = String(encoding || '').toLowerCase();
  if (cte === 'base64') return Buffer.from(body.replace(/[^A-Za-z0-9+/=]/g, ''), 'base64');
  if (cte === 'quoted-printable') return decodeQuotedPrintable(body);
  return Buffer.from(body, 'latin1');
}

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function walk(raw, depth, out) {
  const { head, body } = splitHeaderBody(raw);
  const h = parseHeaders(head);
  const type = (h['content-type'] || 'text/plain').toLowerCase();
  const disposition = h['content-disposition'] || '';
  const filename = decodeHeader(param(disposition, 'filename') || param(h['content-type'], 'name'));

  if (type.startsWith('multipart/') && depth < MAX_DEPTH) {
    const boundary = param(h['content-type'], 'boundary');
    if (boundary) {
      const parts = body.split(new RegExp(`(?:^|\\r?\\n)--${escapeRegExp(boundary)}(?:--)?[ \\t]*(?:\\r?\\n|$)`));
      const inner = { texts: [], attachments: out.attachments };
      for (const part of parts.slice(1)) if (part.trim()) walk(part, depth + 1, inner);
      // Em multipart/alternative o mesmo texto vem em versões diferentes: fica só uma.
      if (type.startsWith('multipart/alternative')) {
        const plain = inner.texts.filter((t) => !t.html);
        out.texts.push(...(plain.length ? plain : inner.texts));
      } else {
        out.texts.push(...inner.texts);
      }
      return;
    }
  }
  if (type.startsWith('message/rfc822') && depth < MAX_DEPTH) {
    walk(body, depth + 1, out);
    return;
  }
  if (/attachment/i.test(disposition) || (filename && !type.startsWith('text/'))) {
    if (filename) out.attachments.push(filename);
    return;
  }
  if (type.startsWith('text/')) {
    const bytes = decodeBody(body, h['content-transfer-encoding']);
    const charset = param(h['content-type'], 'charset');
    let text = charset ? decoderFor(charset.toLowerCase()).decode(bytes) : decodeText(bytes);
    const html = type.startsWith('text/html');
    if (html) text = htmlToText(text);
    out.texts.push({ html, text });
  }
}

/** Extrai o texto de uma mensagem MIME. Retorna { segments, metadata }. */
export function mimeExtract(buf) {
  const raw = buf.toString('latin1');
  const { head } = splitHeaderBody(raw);
  const h = parseHeaders(head);
  const out = { texts: [], attachments: [] };
  walk(raw, 0, out);
  const header = [
    ['Assunto', h.subject],
    ['De', h.from],
    ['Para', h.to],
    ['Cc', h.cc],
  ]
    .filter(([, v]) => v)
    .map(([k, v]) => `${k}: ${decodeHeader(v)}`)
    .join('\n');
  const segments = [];
  if (header) segments.push({ label: 'Cabeçalho', text: header });
  const body = out.texts.map((t) => t.text).join('\n');
  if (body.trim()) segments.push({ label: 'Mensagem', text: body });
  if (out.attachments.length) segments.push({ label: 'Anexos', text: out.attachments.join('\n') });
  const metadata = compact({
    author: h.from ? decodeHeader(h.from) : null,
    title: h.subject ? decodeHeader(h.subject) : null,
    created: h.date ? isoDate(h.date) : null,
  });
  return { segments, metadata };
}
