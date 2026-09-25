// Office Open XML: Word (.docx), Excel (.xlsx) e PowerPoint (.pptx), incluindo variantes com macro.
import { readRels } from './zip.js';
import { xmlToText, tagText, decodeXmlEntities, tidyParagraphs } from './xml.js';

/** Metadados de docProps/core.xml e docProps/app.xml. */
export function ooxmlMetadata(zip) {
  const core = zip.text('docProps/core.xml');
  const app = zip.text('docProps/app.xml');
  return compact({
    author: tagText(core, 'dc:creator'),
    lastModifiedBy: tagText(core, 'cp:lastModifiedBy'),
    created: isoDate(tagText(core, 'dcterms:created')),
    modified: isoDate(tagText(core, 'dcterms:modified')),
    title: tagText(core, 'dc:title'),
    application: tagText(app, 'Application'),
    company: tagText(app, 'Company'),
  });
}

/** Detecta o tipo (docx/xlsx/pptx) pela parte principal do pacote. */
export function detectOoxml(zip) {
  if (zip.has('word/document.xml')) return 'docx';
  if (zip.has('xl/workbook.xml')) return 'xlsx';
  if (zip.has('ppt/presentation.xml')) return 'pptx';
  return null;
}

// mc:Fallback repete o conteúdo de mc:Choice (ex.: caixas de texto em VML) e contaria em dobro.
const WORD_RULES = {
  text: ['w:t', 'w:delText'],
  skip: ['mc:Fallback'],
  close: { 'w:p': '\u0001', 'w:tc': '\t', 'w:tr': '\n' }, // \u0001: fim de parágrafo (ver tidyParagraphs)
  empty: { 'w:tab': '\t', 'w:br': '\n', 'w:cr': '\n', 'w:noBreakHyphen': '-' },
};

const WORD_PARTS = [
  [/^word\/document\.xml$/, 'Documento'],
  [/^word\/header\d*\.xml$/, 'Cabeçalho'],
  [/^word\/footer\d*\.xml$/, 'Rodapé'],
  [/^word\/footnotes\.xml$/, 'Notas de rodapé'],
  [/^word\/endnotes\.xml$/, 'Notas de fim'],
  [/^word\/comments\.xml$/, 'Comentários'],
];

export function docxSegments(zip) {
  const segments = [];
  const names = zip.names();
  for (const [re, label] of WORD_PARTS) {
    for (const name of names.filter((n) => re.test(n)).sort(naturalCompare)) {
      const xml = zip.text(name);
      if (xml) segments.push({ label, text: tidyParagraphs(xmlToText(xml, WORD_RULES)) });
    }
  }
  return segments;
}

const DRAWING_RULES = {
  text: ['a:t', 'p:text'],
  skip: ['mc:Fallback'],
  close: { 'a:p': '\u0001', 'a:tc': '\t', 'a:tr': '\n' },
  empty: { 'a:br': '\n' },
};

export function pptxSegments(zip) {
  const segments = [];
  const pres = zip.text('ppt/presentation.xml') || '';
  const rels = readRels(zip, 'ppt/_rels/presentation.xml.rels', 'ppt');
  const slides = [];
  const re = /<p:sldId\b[^>]*\br:id="([^"]*)"/g;
  let m;
  while ((m = re.exec(pres)) !== null) {
    const rel = rels.get(m[1]);
    if (rel) slides.push(rel.target);
  }
  if (slides.length === 0) {
    slides.push(...zip.names().filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n)).sort(naturalCompare));
  }
  slides.forEach((slidePath, i) => {
    const xml = zip.text(slidePath);
    if (!xml) return;
    let text = tidyParagraphs(xmlToText(xml, DRAWING_RULES));
    const dir = slidePath.slice(0, slidePath.lastIndexOf('/'));
    const file = slidePath.slice(slidePath.lastIndexOf('/') + 1);
    const slideRels = readRels(zip, `${dir}/_rels/${file}.rels`, dir);
    for (const rel of slideRels.values()) {
      if (rel.type.endsWith('/notesSlide')) {
        const notes = zip.text(rel.target);
        if (notes) text += `\n${tidyParagraphs(xmlToText(notes, DRAWING_RULES))}`;
      }
    }
    segments.push({ label: `Slide ${i + 1}`, text });
  });
  for (const name of zip.names().filter((n) => /^ppt\/comments\/.*\.xml$/.test(n)).sort(naturalCompare)) {
    const xml = zip.text(name);
    if (xml) segments.push({ label: 'Comentários', text: tidyParagraphs(xmlToText(xml, DRAWING_RULES)) });
  }
  return segments;
}

// As expressões aceitam prefixo de namespace opcional (<x:c>, <x:row>...), usado por alguns geradores.
const CELL_RE = /<(?:\w+:)?row\b([^>]*?)(\/?)>|<(?:\w+:)?c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:\w+:)?c>)/g;
const ROW_NUM_RE = /\br="(\d+)"/;
const CELL_TYPE_RE = /\bt="([^"]*)"/;
const V_RE = /<(?:\w+:)?v>([^<]*)<\/(?:\w+:)?v>/;
const T_RE = /<(?:\w+:)?t(?:\s[^>]*)?>([^<]*)<\/(?:\w+:)?t>/g;
const PHONETIC_RE = /<(?:\w+:)?rPh\b[\s\S]*?<\/(?:\w+:)?rPh>/g;

function richText(xml) {
  let out = '';
  for (const m of xml.replace(PHONETIC_RE, '').matchAll(T_RE)) out += m[1];
  return decodeXmlEntities(out);
}

function sharedStrings(zip, path) {
  const xml = path ? zip.text(path) : null;
  if (!xml) return [];
  const list = [];
  const re = /<(?:\w+:)?si>([\s\S]*?)<\/(?:\w+:)?si>|<(?:\w+:)?si\/>/g;
  let m;
  while ((m = re.exec(xml)) !== null) list.push(m[1] ? richText(m[1]) : '');
  return list;
}

/** Texto de uma planilha: uma linha de texto por linha da planilha, células separadas por tab. */
function sheetText(xml, sst) {
  const out = [];
  let line = 1;
  let lastRow = 0;
  let exact = true;
  let firstCell = true;
  let m;
  CELL_RE.lastIndex = 0;
  while ((m = CELL_RE.exec(xml)) !== null) {
    if (m[1] !== undefined) {
      const r = Number(ROW_NUM_RE.exec(m[1])?.[1]) || lastRow + 1;
      lastRow = r;
      if (r > line) {
        if (r - line <= 20000) {
          out.push('\n'.repeat(r - line));
        } else {
          out.push('\n');
          exact = false;
        }
        line = r;
      }
      firstCell = true;
      continue;
    }
    const type = CELL_TYPE_RE.exec(m[3])?.[1] || 'n';
    const inner = m[4] || '';
    let value = '';
    if (type === 's') {
      const idx = Number(V_RE.exec(inner)?.[1]);
      value = Number.isInteger(idx) ? sst[idx] || '' : '';
    } else if (type === 'inlineStr') {
      value = richText(inner);
    } else if (type !== 'b' && type !== 'e') {
      value = decodeXmlEntities(V_RE.exec(inner)?.[1] || '');
    }
    if (!value) continue;
    // Quebras de linha dentro da célula (Alt+Enter) não podem deslocar a numeração das linhas.
    value = value.replace(/\r\n|[\r\n]/g, ' ');
    out.push(firstCell ? value : `\t${value}`);
    firstCell = false;
  }
  return { text: out.join(''), exact };
}

export function xlsxSegments(zip) {
  const segments = [];
  const wb = zip.text('xl/workbook.xml') || '';
  const rels = readRels(zip, 'xl/_rels/workbook.xml.rels', 'xl');
  let sstPath = null;
  for (const rel of rels.values()) if (rel.type.endsWith('/sharedStrings')) sstPath = rel.target;
  const sst = sharedStrings(zip, sstPath || (zip.has('xl/sharedStrings.xml') ? 'xl/sharedStrings.xml' : null));
  const re = /<(?:\w+:)?sheet\b([^>]*)\/?>/g;
  let m;
  while ((m = re.exec(wb)) !== null) {
    const name = decodeXmlEntities(/\bname="([^"]*)"/.exec(m[1])?.[1] || '');
    const rid = /\s\w+:id="([^"]*)"/.exec(m[1])?.[1];
    const rel = rid && rels.get(rid);
    if (!rel) continue;
    const xml = zip.text(rel.target);
    if (!xml) continue;
    const { text, exact } = sheetText(xml, sst);
    segments.push({ label: `Planilha "${name}"`, text, lines: exact });
  }
  for (const name of zip.names().filter((n) => /^xl\/(comments\d*|threadedComments\/[^/]+)\.xml$/.test(n)).sort(naturalCompare)) {
    const xml = zip.text(name);
    if (xml) segments.push({ label: 'Comentários', text: xmlToText(xml, { text: ['t', 'text'], close: { comment: '\n', threadedComment: '\n' } }) });
  }
  return segments;
}

export function naturalCompare(a, b) {
  return a.localeCompare(b, undefined, { numeric: true });
}

export function isoDate(value) {
  if (!value) return null;
  const d = new Date(value);
  return Number.isNaN(d.getTime()) ? null : d.toISOString();
}

export function compact(obj) {
  return Object.fromEntries(Object.entries(obj).filter(([, v]) => v !== null && v !== undefined && v !== ''));
}
