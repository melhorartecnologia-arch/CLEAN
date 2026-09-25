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

/** Converte HTML em texto simples. */
export function htmlToText(html) {
  const body = html
    .replace(/<!--[\s\S]*?-->/g, ' ')
    .replace(/<style\b[\s\S]*?<\/style\s*>/gi, ' ')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/(p|div|tr|li|h[1-6]|table|section|article|blockquote|pre|dd|dt|title)\s*>/gi, '\n')
    .replace(/<\/(td|th)\s*>/gi, '\t')
    .replace(/<[^>]*>/g, '');
  return decodeHtmlEntities(body);
}
