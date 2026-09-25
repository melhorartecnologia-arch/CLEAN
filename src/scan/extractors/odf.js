// OpenDocument (LibreOffice): .odt, .ods e .odp.
import { xmlToText, tagText, attr, decodeXmlEntities, tidyParagraphs } from './xml.js';
import { compact, isoDate } from './ooxml.js';

export function detectOdf(zip) {
  const mime = zip.text('mimetype')?.trim() || '';
  if (mime.includes('opendocument.text')) return 'odt';
  if (mime.includes('opendocument.spreadsheet')) return 'ods';
  if (mime.includes('opendocument.presentation')) return 'odp';
  if (mime.includes('opendocument.graphics')) return 'odg';
  return zip.has('content.xml') && zip.has('META-INF/manifest.xml') ? 'odt' : null;
}

export function odfIsEncrypted(zip) {
  return /encryption-data/.test(zip.text('META-INF/manifest.xml') || '');
}

export function odfMetadata(zip) {
  const meta = zip.text('meta.xml');
  return compact({
    author: tagText(meta, 'meta:initial-creator'),
    lastModifiedBy: tagText(meta, 'dc:creator'),
    created: isoDate(tagText(meta, 'meta:creation-date')),
    modified: isoDate(tagText(meta, 'dc:date')),
    title: tagText(meta, 'dc:title'),
    application: tagText(meta, 'meta:generator'),
  });
}

const RULES = {
  text: null,
  skip: [
    'office:automatic-styles',
    'office:font-face-decls',
    'office:scripts',
    'office:forms',
    'dc:creator',
    'dc:date',
    'meta:creator-initials',
    'table:named-expressions',
    'text:note-citation',
  ],
  // \u0001 marca fim de parágrafo (ver tidyParagraphs).
  close: { 'text:p': '\u0001', 'text:h': '\u0001', 'table:table-cell': '\t', 'table:table-row': '\n' },
  empty: {
    'text:note': '\n',
    'text:tab': '\t',
    'text:line-break': '\n',
    'text:s': (tag) => ' '.repeat(Math.min(Number(attr(tag, 'text:c')) || 1, 100)),
  },
  // Linhas vazias repetidas ("number-rows-repeated") mantêm a numeração das linhas da planilha.
  onOpen: (name, tag) => {
    if (name !== 'table:table-row') return null;
    const repeated = Number(attr(tag, 'table:number-rows-repeated')) || 1;
    return repeated > 1 && repeated <= 10000 ? '\n'.repeat(repeated - 1) : null;
  },
};

/** Parágrafos dentro de células viram texto da célula (sem quebrar a linha da tabela). */
function toText(xml) {
  return tidyParagraphs(xmlToText(xml, RULES));
}

// Planilhas: cada linha da planilha é uma linha de texto; parágrafos e quebras dentro de uma
// célula viram espaço para não deslocar a numeração das linhas.
const SHEET_RULES = {
  ...RULES,
  close: { 'text:p': ' ', 'text:h': ' ', 'table:table-cell': '\t', 'table:table-row': '\n' },
  empty: { ...RULES.empty, 'text:line-break': ' ' },
};

function sheetToText(xml) {
  return xmlToText(xml, SHEET_RULES)
    .replace(/ +\t/g, '\t')
    .replace(/[ \t]+\n/g, '\n');
}

/** A numeração de linhas só é confiável se não houver conteúdo após um bloco enorme de linhas vazias. */
function exactRows(tableXml) {
  for (const m of tableXml.matchAll(/table:number-rows-repeated="(\d+)"/g)) {
    if (Number(m[1]) > 10000 && tableXml.indexOf('<text:p', m.index) !== -1) return false;
  }
  return true;
}

export function odfSegments(zip, kind) {
  const xml = zip.text('content.xml');
  if (!xml) return [];
  if (kind === 'ods') {
    const segments = [];
    for (const m of xml.matchAll(/<table:table\b([^>]*)>([\s\S]*?)<\/table:table>/g)) {
      const name = decodeXmlEntities(/\btable:name="([^"]*)"/.exec(m[1])?.[1] || '');
      segments.push({ label: `Planilha "${name}"`, text: sheetToText(m[2]), lines: exactRows(m[2]) });
    }
    return segments;
  }
  if (kind === 'odp' || kind === 'odg') {
    const segments = [];
    let i = 0;
    for (const m of xml.matchAll(/<draw:page\b[^>]*>([\s\S]*?)<\/draw:page>/g)) {
      i++;
      segments.push({ label: kind === 'odp' ? `Slide ${i}` : `Página ${i}`, text: toText(m[1]) });
    }
    return segments;
  }
  return [{ label: 'Documento', text: toText(xml) }];
}
