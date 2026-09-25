// Utilitários leves para extrair texto de XML (Office Open XML, OpenDocument) e HTML.

const XML_ENTITIES = { lt: '<', gt: '>', amp: '&', quot: '"', apos: "'" };

// Entidades HTML do bloco Latin-1 (U+00A0 a U+00FF), na ordem dos códigos.
const LATIN1_NAMES = (
  'nbsp iexcl cent pound curren yen brvbar sect uml copy ordf laquo not shy reg macr deg plusmn ' +
  'sup2 sup3 acute micro para middot cedil sup1 ordm raquo frac14 frac12 frac34 iquest Agrave Aacute ' +
  'Acirc Atilde Auml Aring AElig Ccedil Egrave Eacute Ecirc Euml Igrave Iacute Icirc Iuml ETH Ntilde ' +
  'Ograve Oacute Ocirc Otilde Ouml times Oslash Ugrave Uacute Ucirc Uuml Yacute THORN szlig agrave ' +
  'aacute acirc atilde auml aring aelig ccedil egrave eacute ecirc euml igrave iacute icirc iuml eth ' +
  'ntilde ograve oacute ocirc otilde ouml divide oslash ugrave uacute ucirc uuml yacute thorn yuml'
).split(' ');

const HTML_ENTITIES = {
  ...XML_ENTITIES,
  ...Object.fromEntries(LATIN1_NAMES.map((name, i) => [name, String.fromCharCode(0xa0 + i)])),
  ndash: '–',
  mdash: '—',
  lsquo: '‘',
  rsquo: '’',
  sbquo: '‚',
  ldquo: '“',
  rdquo: '”',
  bdquo: '„',
  hellip: '…',
  bull: '•',
  euro: '€',
  trade: '™',
  OElig: 'Œ',
  oelig: 'œ',
  Scaron: 'Š',
  scaron: 'š',
};

function fromCodePoint(code) {
  if (!Number.isFinite(code) || code < 0 || code > 0x10ffff) return '';
  return String.fromCodePoint(code);
}

function decodeWith(table, value) {
  if (!value || value.indexOf('&') === -1) return value;
  return value.replace(/&(#x[0-9a-f]+|#\d+|[a-z][a-z0-9]*);/gi, (all, ent) => {
    if (ent[0] === '#') {
      return fromCodePoint(ent[1] === 'x' || ent[1] === 'X' ? parseInt(ent.slice(2), 16) : parseInt(ent.slice(1), 10));
    }
    return table[ent] ?? all;
  });
}

export const decodeXmlEntities = (value) => decodeWith(XML_ENTITIES, value);
export const decodeHtmlEntities = (value) => decodeWith(HTML_ENTITIES, value);

/** Valor de um atributo dentro do texto de uma tag de abertura. */
export function attr(tag, name) {
  const re = new RegExp(`\\s${name.replace(/[.:]/g, '\\$&')}\\s*=\\s*("([^"]*)"|'([^']*)')`);
  const m = re.exec(tag);
  return m ? decodeXmlEntities(m[2] ?? m[3]) : null;
}

/** Texto do primeiro elemento com o nome informado (ex.: "cp:lastModifiedBy"). */
export function tagText(xml, name) {
  if (!xml) return null;
  const n = name.replace(/[.:]/g, '\\$&');
  const m = new RegExp(`<${n}(?:\\s[^>]*)?>([\\s\\S]*?)</${n}>`).exec(xml);
  if (!m) return null;
  const value = decodeXmlEntities(m[1].replace(/<[^>]*>/g, '')).trim();
  return value || null;
}

const TAG_NAME = /^\/?\s*([^\s/>]+)/;

/**
 * Converte XML em texto percorrendo tags e nós de texto.
 * rules.text: tags cujo conteúdo é texto (null = todo nó de texto conta)
 * rules.skip: tags cujo conteúdo deve ser ignorado
 * rules.close: { tag: 'caractere emitido ao fechar' }
 * rules.empty: { tag: 'caractere emitido para a tag' } (tags vazias ou de abertura)
 * rules.onOpen: função opcional (nome, tagCompleta) que pode devolver texto a emitir
 */
export function xmlToText(xml, rules) {
  const textTags = rules.text ? new Set(rules.text) : null;
  const skipTags = new Set(rules.skip || []);
  const closeMap = rules.close || {};
  const emptyMap = rules.empty || {};
  const out = [];
  let inText = 0;
  let inSkip = 0;
  const tokenRe = /<([^>]*)>|([^<]+)/g;
  let m;
  while ((m = tokenRe.exec(xml)) !== null) {
    if (m[2] !== undefined) {
      if (inSkip === 0 && (textTags === null || inText > 0)) out.push(decodeXmlEntities(m[2]));
      continue;
    }
    const tag = m[1];
    const first = tag.charCodeAt(0);
    if (first === 63 /* ? */ || first === 33 /* ! */) continue;
    const closing = first === 47; /* / */
    const selfClosing = !closing && tag.charCodeAt(tag.length - 1) === 47;
    const name = TAG_NAME.exec(tag)?.[1] ?? '';
    if (closing) {
      if (skipTags.has(name)) inSkip = Math.max(0, inSkip - 1);
      else if (textTags && textTags.has(name)) inText = Math.max(0, inText - 1);
      if (inSkip === 0 && closeMap[name] !== undefined) out.push(closeMap[name]);
      continue;
    }
    if (skipTags.has(name)) {
      if (!selfClosing) inSkip++;
      continue;
    }
    if (inSkip > 0) continue;
    if (emptyMap[name] !== undefined) out.push(typeof emptyMap[name] === 'function' ? emptyMap[name](tag) : emptyMap[name]);
    if (rules.onOpen) {
      const extra = rules.onOpen(name, tag);
      if (extra) out.push(extra);
    }
    if (!selfClosing && textTags && textTags.has(name)) inText++;
    if (selfClosing && closeMap[name] !== undefined) out.push(closeMap[name]);
  }
  return out.join('');
}

/**
 * Pós-processa texto em que o fim de parágrafo foi marcado com \u0001: parágrafos no fim de uma
 * célula de tabela não quebram a linha (a célula termina com tab e a linha da tabela com \n).
 */
export function tidyParagraphs(text) {
  return text.replace(/\u0001(?=\t)/g, '').replace(/\u0001/g, '\n').replace(/\t+\n/g, '\n');
}

const BLOCK_CLOSE = new Set(['p', 'div', 'tr', 'li', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'table', 'section', 'article', 'blockquote', 'pre', 'dd', 'dt', 'title']);
const RAW_TEXT = new Set(['style', 'script']);
const TAG_START = /^<(\/?)([a-zA-Z][a-zA-Z0-9:-]*)/;

/**
 * Converte HTML em texto simples. Percorre o texto uma única vez (tempo linear mesmo com HTML
 * malformado ou malicioso, ex.: milhares de "<!--" sem fechamento).
 */
export function htmlToText(html) {
  const out = [];
  const n = html.length;
  let i = 0;
  let nextGt = -1; // posição do próximo ">" (reaproveitada entre tags)
  let lower = null; // cópia em minúsculas, só se houver <style> ou <script>
  while (i < n) {
    const lt = html.indexOf('<', i);
    if (lt === -1) {
      out.push(html.slice(i));
      break;
    }
    if (lt > i) out.push(html.slice(i, lt));
    if (html.startsWith('<!--', lt)) {
      const end = html.indexOf('-->', lt + 4);
      out.push(' ');
      i = end === -1 ? n : end + 3;
      continue;
    }
    if (nextGt !== -2 && nextGt <= lt) nextGt = html.indexOf('>', lt + 1);
    if (nextGt === -1) nextGt = -2; // não há mais ">": o restante é texto
    const m = nextGt >= 0 ? TAG_START.exec(html.slice(lt, Math.min(nextGt + 1, lt + 80))) : null;
    const first = html.charCodeAt(lt + 1);
    if (!m && nextGt >= 0 && (first === 33 || first === 63 || first === 47)) {
      i = nextGt + 1; // <!doctype>, <?xml ?>, </ ...>
      continue;
    }
    if (!m) {
      out.push('<'); // "<" solto é texto (ex.: "a < b")
      i = lt + 1;
      continue;
    }
    i = nextGt + 1;
    const closing = m[1] === '/';
    const name = m[2].toLowerCase();
    if (!closing && RAW_TEXT.has(name)) {
      lower ||= html.toLowerCase();
      const close = lower.indexOf(`</${name}`, i);
      const gt = close === -1 ? -1 : html.indexOf('>', close);
      i = gt === -1 ? n : gt + 1;
      out.push(' ');
      continue;
    }
    if (name === 'br') out.push('\n');
    else if (closing && BLOCK_CLOSE.has(name)) out.push('\n');
    else if (closing && (name === 'td' || name === 'th')) out.push('\t');
  }
  return decodeHtmlEntities(out.join(''));
}
