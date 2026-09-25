// Mensagens MIME: e-mails (.eml e mensagens baixadas das caixas de e-mail) e páginas da web salvas
// em arquivo único (.mht/.mhtml). Decodifica base64 e quoted-printable, conjuntos de caracteres e
// cabeçalhos codificados (RFC 2047 e RFC 2231) e separa o corpo da mensagem dos anexos.
import { decoderFor, decodeText } from './text.js';
import { htmlToText } from './xml.js';
import { isoDate } from './ooxml.js';

export const MIME_EXTENSIONS = new Set(['.eml', '.mht', '.mhtml', '.mhtm']);

const MAX_DEPTH = 12; // partes multipart aninhadas
const MAX_PARTS = 1000;
const MAX_HEADER = 256 * 1024;

const utf8Fatal = new TextDecoder('utf-8', { fatal: true });

// ---------------------------------------------------------------------------------------------
// Cabeçalhos

/** Bytes acima de 0x7F sem codificação (comuns em cabeçalhos): UTF-8 ou, se inválido, Windows-1252. */
function rawHeaderText(value) {
  if (!/[\x80-\xff]/.test(value)) return value;
  const bytes = Buffer.from(value, 'latin1');
  try {
    return utf8Fatal.decode(bytes);
  } catch {
    return decoderFor('windows-1252').decode(bytes);
  }
}

function decodeQuotedPrintable(text) {
  const s = text.replace(/=[ \t]*\r?\n/g, '');
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

function decodeBytes(charset, bytes) {
  try {
    return decoderFor(charset).decode(bytes);
  } catch {
    return decoderFor('windows-1252').decode(bytes);
  }
}

/**
 * Decodifica palavras codificadas (=?utf-8?B?...?= / =?iso-8859-1?Q?...?=). Palavras vizinhas são
 * unidas antes da decodificação, pois um caractere pode estar dividido entre duas delas.
 */
export function decodeHeader(value) {
  const text = rawHeaderText(String(value ?? ''));
  if (text.indexOf('=?') === -1) return text;
  const re = /=\?([^?\s]+)\?([bBqQ])\?([^?\s]*)\?=/g;
  let out = '';
  let last = 0;
  let pending = null;
  const flush = () => {
    if (pending) out += decodeBytes(pending.charset, Buffer.concat(pending.bytes));
    pending = null;
  };
  let m;
  while ((m = re.exec(text)) !== null) {
    const between = text.slice(last, m.index);
    // Espaços entre duas palavras codificadas não fazem parte do texto (RFC 2047, seção 6.2).
    if (!(pending && /^[ \t\r\n]*$/.test(between))) {
      flush();
      out += between;
    }
    const charset = m[1].split('*')[0].toLowerCase();
    const bytes = m[2].toUpperCase() === 'B' ? Buffer.from(m[3], 'base64') : decodeQuotedPrintable(m[3].replace(/_/g, ' '));
    if (pending && pending.charset !== charset) flush();
    pending ||= { charset, bytes: [] };
    pending.bytes.push(bytes);
    last = re.lastIndex;
  }
  flush();
  return out + text.slice(last);
}

function parseHeaderBlock(block) {
  const headers = Object.create(null);
  const unfolded = block.replace(/\r?\n(?=[ \t])/g, '');
  for (const line of unfolded.split(/\r?\n/)) {
    const idx = line.indexOf(':');
    if (idx <= 0) continue;
    const name = line.slice(0, idx).trim().toLowerCase();
    if (!name || /\s/.test(name)) continue;
    (headers[name] ||= []).push(line.slice(idx + 1).trim());
  }
  return headers;
}

/** Divide "valor; a=1; b="x;y"" pelos ponto e vírgulas que não estão entre aspas. */
function splitParams(value) {
  const parts = [];
  let current = '';
  let quoted = false;
  for (let i = 0; i < value.length; i++) {
    const c = value[i];
    if (c === '\\' && quoted && i + 1 < value.length) {
      current += c + value[++i];
    } else if (c === '"') {
      quoted = !quoted;
      current += c;
    } else if (c === ';' && !quoted) {
      parts.push(current);
      current = '';
    } else {
      current += c;
    }
  }
  parts.push(current);
  return parts;
}

function unquote(value) {
  const v = value.trim();
  if (v.length >= 2 && v.startsWith('"') && v.endsWith('"')) return v.slice(1, -1).replace(/\\(.)/g, '$1');
  return v;
}

/**
 * Valor e parâmetros de Content-Type / Content-Disposition, inclusive parâmetros RFC 2231
 * (filename*=utf-8''rel%C3%B3rio.pdf, divididos em filename*0*=, filename*1*=...).
 */
export function parseStructuredHeader(value) {
  const [main, ...rest] = splitParams(String(value || ''));
  const params = Object.create(null);
  const sections = Object.create(null);
  for (const item of rest) {
    const eq = item.indexOf('=');
    if (eq <= 0) continue;
    const key = item.slice(0, eq).trim().toLowerCase();
    const val = unquote(item.slice(eq + 1));
    const m = /^([^*]+)(?:\*(\d+))?(\*)?$/.exec(key);
    if (!m) continue;
    if (m[2] === undefined && !m[3]) {
      if (!(m[1] in params)) params[m[1]] = val;
      continue;
    }
    (sections[m[1]] ||= []).push({ n: Number(m[2] || 0), extended: Boolean(m[3]), value: val });
  }
  for (const [name, list] of Object.entries(sections)) {
    list.sort((a, b) => a.n - b.n);
    if (!list.some((section) => section.extended)) {
      // Continuações simples (name*0=, name*1=): o texto unido ainda pode ter palavras RFC 2047.
      params[name] = list.map((section) => section.value).join('');
      continue;
    }
    let charset = 'utf-8';
    const bytes = [];
    list.forEach((section, i) => {
      let v = section.value;
      if (section.extended) {
        if (i === 0) {
          const parts = v.split("'");
          if (parts.length >= 3) {
            charset = parts[0] || 'utf-8';
            v = parts.slice(2).join("'");
          }
        }
        bytes.push(Buffer.from(v.replace(/%([0-9A-Fa-f]{2})/g, (_, h) => String.fromCharCode(parseInt(h, 16))), 'latin1'));
      } else {
        bytes.push(Buffer.from(v, 'latin1'));
      }
    });
    params[name] = decodeBytes(charset.toLowerCase(), Buffer.concat(bytes));
    params[`${name}*`] = true; // já decodificado
  }
  for (const key of Object.keys(params)) {
    if (key.endsWith('*')) continue;
    if (!params[`${key}*`]) params[key] = decodeHeader(params[key]);
  }
  return { value: main.trim().toLowerCase(), params };
}

// ---------------------------------------------------------------------------------------------
// Endereços

/** Divide uma lista de endereços pelas vírgulas fora de aspas, <> e comentários (). */
function splitAddressList(value) {
  const items = [];
  let current = '';
  let quoted = false;
  let angle = 0;
  let comment = 0;
  for (let i = 0; i < value.length; i++) {
    const c = value[i];
    if (quoted) {
      if (c === '\\' && i + 1 < value.length) current += c + value[++i];
      else {
        if (c === '"') quoted = false;
        current += c;
      }
      continue;
    }
    if (c === '"') quoted = true;
    else if (c === '(') comment++;
    else if (c === ')' && comment) comment--;
    else if (c === '<' && !comment) angle++;
    else if (c === '>' && angle) angle--;
    else if ((c === ',' || c === ';') && !angle && !comment) {
      items.push(current);
      current = '';
      continue;
    }
    current += c;
  }
  items.push(current);
  return items;
}

/** Lê endereços como "Nome <email>", "email (Nome)" ou "email". Retorna [{ name, address }]. */
export function parseAddresses(values) {
  const out = [];
  for (const value of [].concat(values || [])) {
    for (let item of splitAddressList(String(value))) {
      item = item.trim();
      // Grupo "Equipe: a@x, b@y;" — descarta o nome do grupo.
      const group = /^[^"<>@(]*:\s*/.exec(item);
      if (group && !item.slice(0, group[0].length).includes('@')) item = item.slice(group[0].length);
      if (!item) continue;
      let name = '';
      let address = '';
      const angle = /<([^<>]*)>\s*(?:\([^)]*\)\s*)?$/.exec(item);
      if (angle) {
        address = angle[1].trim();
        name = item.slice(0, angle.index).trim();
      } else {
        const comment = /\(([^)]*)\)/.exec(item);
        name = comment ? comment[1] : '';
        address = item.replace(/\([^)]*\)/g, '').trim();
      }
      name = decodeHeader(unquote(name)).trim();
      address = decodeHeader(address).replace(/^mailto:/i, '').trim();
      if (!address && !name) continue;
      out.push({ name, address });
    }
  }
  return out;
}

export function formatAddress(entry) {
  if (!entry) return '';
  if (entry.name && entry.address && entry.name.toLowerCase() !== entry.address.toLowerCase()) return `${entry.name} <${entry.address}>`;
  return entry.address || entry.name || '';
}

// ---------------------------------------------------------------------------------------------
// Estrutura da mensagem

/** Localiza cabeçalho e corpo de uma parte entre as posições start e end do texto. */
function readEntity(raw, start, end) {
  let headEnd;
  let bodyStart;
  if (raw.startsWith('\r\n', start)) {
    headEnd = start;
    bodyStart = start + 2;
  } else if (raw.charCodeAt(start) === 10) {
    headEnd = start;
    bodyStart = start + 1;
  } else {
    const firstLine = raw.slice(start, Math.min(end, start + 1000)).split(/\r?\n/, 1)[0];
    if (!/^[\x21-\x39\x3b-\x7e]+[ \t]*:/.test(firstLine)) {
      // Parte sem cabeçalho (mensagem malformada): tudo é corpo.
      headEnd = start;
      bodyStart = start;
    } else {
      const re = /\r?\n\r?\n/g;
      re.lastIndex = start;
      const m = re.exec(raw);
      if (!m || m.index >= end) {
        headEnd = end;
        bodyStart = end;
      } else {
        headEnd = m.index;
        bodyStart = Math.min(end, m.index + m[0].length);
      }
    }
  }
  const headers = parseHeaderBlock(raw.slice(start, Math.min(headEnd, start + MAX_HEADER)));
  return { headers, bodyStart, bodyEnd: end };
}

/** Posições das partes de um corpo multipart. */
function splitMultipart(raw, start, end, boundary) {
  const delimiter = `--${boundary}`;
  const parts = [];
  let partStart = -1;
  let pos = start;
  while (pos < end) {
    const idx = raw.indexOf(delimiter, pos);
    if (idx === -1 || idx + delimiter.length > end) break;
    if (idx !== start && raw.charCodeAt(idx - 1) !== 10) {
      pos = idx + 1;
      continue;
    }
    const after = idx + delimiter.length;
    const close = raw.startsWith('--', after);
    let lineEnd = raw.indexOf('\n', after);
    if (lineEnd === -1 || lineEnd > end) lineEnd = end;
    // O restante da linha só pode ter espaços (senão é outro separador que começa igual).
    if (!/^[ \t\r]*$/.test(raw.slice(close ? after + 2 : after, lineEnd))) {
      pos = idx + 1;
      continue;
    }
    if (partStart !== -1) {
      let partEnd = idx;
      if (partEnd > partStart && raw.charCodeAt(partEnd - 1) === 10) partEnd--;
      if (partEnd > partStart && raw.charCodeAt(partEnd - 1) === 13) partEnd--;
      parts.push({ start: partStart, end: partEnd });
    }
    if (close) return parts;
    partStart = Math.min(end, lineEnd + 1);
    pos = partStart;
  }
  if (partStart !== -1 && partStart < end) parts.push({ start: partStart, end });
  return parts;
}

function decodeTransfer(raw, start, end, encoding) {
  const body = raw.slice(start, end);
  const cte = String(encoding || '').trim().toLowerCase();
  if (cte === 'base64') return Buffer.from(body.replace(/[^A-Za-z0-9+/]/g, ''), 'base64');
  if (cte === 'quoted-printable') return decodeQuotedPrintable(body);
  return Buffer.from(body, 'latin1');
}

function decodeTextPart(bytes, charset) {
  const cs = String(charset || '').trim().toLowerCase();
  if (!cs || cs === 'us-ascii' || cs === 'ascii' || cs === 'iso-8859-1' || cs === 'latin1') return decodeText(bytes);
  return decodeBytes(cs, bytes);
}

const EXTENSIONS = {
  'application/pdf': '.pdf',
  'application/zip': '.zip',
  'application/msword': '.doc',
  'application/vnd.ms-excel': '.xls',
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document': '.docx',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': '.xlsx',
  'application/vnd.openxmlformats-officedocument.presentationml.presentation': '.pptx',
  'application/ms-tnef': '.dat',
  'application/vnd.ms-outlook': '.msg',
  'image/png': '.png',
  'image/jpeg': '.jpg',
  'image/gif': '.gif',
  'text/calendar': '.ics',
  'text/csv': '.csv',
  'text/plain': '.txt',
  'text/html': '.html',
  'text/xml': '.xml',
  'message/rfc822': '.eml',
  'message/global': '.eml',
  'application/pkcs7-signature': '.p7s',
  'application/x-pkcs7-signature': '.p7s',
  'application/pkcs7-mime': '.p7m',
  'application/x-pkcs7-mime': '.p7m',
};

function nestedSubject(bytes) {
  const raw = bytes.subarray(0, 64 * 1024).toString('latin1');
  const entity = readEntity(raw, 0, raw.length);
  return decodeHeader(entity.headers.subject?.[0] || '').trim();
}

/** Nome seguro para exibição (sem quebras de linha nem caracteres de controle). */
function cleanName(name) {
  // eslint-disable-next-line no-control-regex
  return String(name || '').replace(/[\x00-\x1f\x7f]+/g, ' ').trim().slice(0, 255);
}

function walk(ctx, start, end, depth, defaultType) {
  if (++ctx.parts > MAX_PARTS) return;
  const { raw } = ctx;
  const entity = readEntity(raw, start, end);
  const h = entity.headers;
  const ct = parseStructuredHeader(h['content-type']?.[0] || defaultType);
  const type = ct.value || 'text/plain';
  const disposition = parseStructuredHeader(h['content-disposition']?.[0] || '');
  const encoding = h['content-transfer-encoding']?.[0];

  if (type.startsWith('multipart/') && ct.params.boundary && depth < MAX_DEPTH) {
    const parts = splitMultipart(raw, entity.bodyStart, entity.bodyEnd, ct.params.boundary);
    if (type === 'multipart/signed') {
      // A primeira parte é o conteúdo; a segunda, a assinatura.
      if (parts[0]) walk(ctx, parts[0].start, parts[0].end, depth + 1, 'text/plain');
      return;
    }
    if (type === 'multipart/encrypted') {
      ctx.encrypted = true;
      return;
    }
    if (type === 'multipart/alternative') {
      // Versões diferentes do mesmo texto: fica a que tiver mais conteúdo.
      let best = null;
      for (const part of parts) {
        const texts = ctx.texts;
        ctx.texts = [];
        walk(ctx, part.start, part.end, depth + 1, 'text/plain');
        const found = ctx.texts;
        ctx.texts = texts;
        const size = found.reduce((sum, t) => sum + t.text.trim().length, 0);
        if (size > 0 && (!best || size > best.size)) best = { found, size };
      }
      if (best) ctx.texts.push(...best.found);
      return;
    }
    const childDefault = type === 'multipart/digest' ? 'message/rfc822' : 'text/plain';
    for (const part of parts) walk(ctx, part.start, part.end, depth + 1, childDefault);
    return;
  }

  // Na raiz, application/pkcs7-mime é a própria mensagem em S/MIME. Dentro de outra parte, é um
  // arquivo assinado ou cifrado anexado (ex.: contrato.pdf.p7m da ICP-Brasil) e entra como anexo,
  // assim como as assinaturas .p7s anexadas (a assinatura de um multipart/signed nem é visitada).
  if ((type === 'application/pkcs7-mime' || type === 'application/x-pkcs7-mime') && depth === 0) {
    const smime = String(ct.params['smime-type'] || '').toLowerCase();
    if (smime === 'signed-data' || smime === 'certs-only') ctx.opaqueSigned = true;
    else ctx.encrypted = true;
    return;
  }

  const incomplete = ctx.truncated && entity.bodyEnd >= raw.length;
  const filename = cleanName(disposition.params.filename || ct.params.name);
  const isBodyText = (type === 'text/plain' || type === 'text/html') && !filename && disposition.value !== 'attachment';
  if (isBodyText) {
    const bytes = decodeTransfer(raw, entity.bodyStart, entity.bodyEnd, encoding);
    let text = decodeTextPart(bytes, ct.params.charset);
    const html = type === 'text/html';
    if (html) text = htmlToText(text);
    ctx.texts.push({ html, text, incomplete });
    return;
  }

  const message = type === 'message/rfc822' || type === 'message/global';
  // Só imagens contam como "embutidas no corpo": o Apple Mail envia anexos comuns (PDF etc.)
  // com Content-Disposition: inline.
  const inline = type.startsWith('image/') && (disposition.value === 'inline' || (!disposition.value && Boolean(h['content-id'])));
  // Parte omitida pelo conector (anexo de mensagem acima do limite de tamanho, não baixado). O
  // cabeçalho só vale com o código gerado pelo próprio conector: uma mensagem recebida não consegue
  // usá-lo para esconder um anexo da análise.
  const omitted = Boolean(ctx.omittedToken) && h['x-clean-omitted']?.[0] === ctx.omittedToken;
  let data = null;
  let size;
  if (ctx.decodeAttachments || message) {
    data = decodeTransfer(raw, entity.bodyStart, entity.bodyEnd, encoding);
    size = data.length;
  } else {
    const length = entity.bodyEnd - entity.bodyStart;
    size = String(encoding || '').toLowerCase().includes('base64') ? Math.floor(length * 0.74) : length;
  }
  if (omitted) size = Number(h['x-clean-size']?.[0]) || size;
  let name = filename;
  if (!name && message && data) name = nestedSubject(data);
  if (!name) {
    const location = /[^/\\]+$/.exec(String(h['content-location']?.[0] || '').split(/[?#]/)[0])?.[0];
    name = cleanName(location) || (inline && type.startsWith('image/') ? 'imagem-embutida' : 'anexo-sem-nome');
  }
  const ext = EXTENSIONS[type];
  if (ext && !name.toLowerCase().endsWith(ext) && (message || !/\.[a-z0-9]{1,5}$/i.test(name))) name += ext;
  ctx.attachments.push({
    name,
    contentType: type,
    size,
    inline,
    message,
    incomplete: incomplete || omitted,
    contentId: String(h['content-id']?.[0] || '').replace(/^<|>$/g, '') || null,
    data: ctx.decodeAttachments ? data : null,
  });
}

/**
 * Lê uma mensagem MIME. Retorna cabeçalhos decodificados, textos do corpo e anexos:
 *   { subject, from, sender, to, cc, bcc, replyTo, date, messageId, texts: [{ html, text }],
 *     attachments: [{ name, contentType, size, inline, message, incomplete, data }],
 *     encrypted, opaqueSigned, headers }
 * truncated: a mensagem foi cortada (a última parte fica marcada como incompleta).
 * decodeAttachments: false dispensa a decodificação dos anexos (só nomes e tamanhos aproximados).
 * omittedToken: código das partes marcadas pelo conector como não baixadas (X-Clean-Omitted).
 */
export function parseMime(buf, { truncated = false, decodeAttachments = true, omittedToken = null } = {}) {
  let raw = Buffer.isBuffer(buf) ? buf.toString('latin1') : String(buf);
  // Arquivos no formato mbox começam com a linha "From remetente data".
  if (raw.startsWith('From ')) raw = raw.slice(raw.indexOf('\n') + 1);
  const ctx = { raw, truncated, decodeAttachments, omittedToken, parts: 0, texts: [], attachments: [], encrypted: false, opaqueSigned: false };
  const top = readEntity(raw, 0, raw.length);
  walk(ctx, 0, raw.length, 0, 'text/plain');
  const h = top.headers;
  return {
    headers: h,
    subject: decodeHeader(h.subject?.[0] || '').trim(),
    from: parseAddresses(h.from),
    sender: parseAddresses(h.sender),
    to: parseAddresses(h.to),
    cc: parseAddresses(h.cc),
    bcc: parseAddresses(h.bcc),
    replyTo: parseAddresses(h['reply-to']),
    date: isoDate(String(h.date?.[0] || '').replace(/\s*\([^)]*\)\s*$/, '')),
    messageId: String(h['message-id']?.[0] || '').trim() || null,
    texts: ctx.texts,
    attachments: ctx.attachments,
    encrypted: ctx.encrypted,
    opaqueSigned: ctx.opaqueSigned,
  };
}
