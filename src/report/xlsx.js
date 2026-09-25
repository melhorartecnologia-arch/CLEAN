// Gerador de planilhas .xlsx (Office Open XML) em fluxo contínuo, sem dependências além do fflate.
// As linhas são geradas e compactadas aos poucos, então relatórios grandes não esgotam a memória;
// abas com mais linhas do que o Excel suporta continuam em abas adicionais ("Nome (2)").
import { once } from 'node:events';
import { Zip, ZipDeflate, strToU8 } from 'fflate';

const MAX_CELL = 32767;
export const MAX_ROWS = 1_048_576; // limite de linhas de uma planilha do Excel
const BATCH = 2000;

/** Remove caracteres proibidos em XML 1.0 e escapa os especiais. */
export function xmlEscape(value) {
  return String(value)
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\ufffe\uffff]/g, '')
    .replace(/[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/g, '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

export function columnName(index) {
  let n = index + 1;
  let name = '';
  while (n > 0) {
    const rest = (n - 1) % 26;
    name = String.fromCharCode(65 + rest) + name;
    n = Math.floor((n - 1) / 26);
  }
  return name;
}

/** Data -> número de série do Excel (no fuso horário local do servidor). */
function excelDate(date) {
  return (date.getTime() - date.getTimezoneOffset() * 60000) / 86400000 + 25569;
}

// Índices em cellXfs (ver STYLES)
const STYLE = { default: 0, header: 1, date: 2, wrap: 3, bold: 4, title: 5, int: 6 };

const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="1"><numFmt numFmtId="164" formatCode="dd/mm/yyyy hh:mm"/></numFmts>
<fonts count="4"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font><font><b/><sz val="11"/><name val="Calibri"/><family val="2"/></font><font><b/><sz val="14"/><name val="Calibri"/><family val="2"/></font></fonts>
<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1F4E79"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="7">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="3" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;

function cellXml(ref, cell) {
  if (cell === null || cell === undefined || cell === '') return '';
  let value = cell;
  let style = null;
  if (typeof cell === 'object' && !(cell instanceof Date)) {
    value = cell.v;
    style = cell.s || null;
    if (value === null || value === undefined) return style ? `<c r="${ref}" s="${STYLE[style]}"/>` : '';
  }
  if (value instanceof Date) {
    if (Number.isNaN(value.getTime())) return '';
    return `<c r="${ref}" s="${STYLE[style || 'date']}"><v>${excelDate(value)}</v></c>`;
  }
  if (typeof value === 'number' && Number.isFinite(value)) {
    const s = style ? ` s="${STYLE[style]}"` : Number.isInteger(value) ? ` s="${STYLE.int}"` : '';
    return `<c r="${ref}"${s}><v>${value}</v></c>`;
  }
  let text = String(value);
  if (text.length > MAX_CELL) text = `${text.slice(0, MAX_CELL - 1)}…`;
  const s = style ? ` s="${STYLE[style]}"` : '';
  return `<c r="${ref}" t="inlineStr"${s}><is><t xml:space="preserve">${xmlEscape(text)}</t></is></c>`;
}

function rowXml(row, r) {
  let cells = '';
  for (let c = 0; c < row.length; c++) cells += cellXml(`${columnName(c)}${r}`, row[c]);
  return `<row r="${r}">${cells}</row>`;
}

function sheetHead(sheet) {
  const parts = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'];
  if (sheet.header) {
    parts.push('<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>');
  } else {
    parts.push('<sheetViews><sheetView workbookViewId="0"/></sheetViews>');
  }
  if (sheet.cols && sheet.cols.length) {
    parts.push('<cols>');
    sheet.cols.forEach((w, i) => parts.push(`<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`));
    parts.push('</cols>');
  }
  parts.push('<sheetData>');
  return parts.join('');
}

/** Nome de planilha válido no Excel (até 31 caracteres, sem : \ / ? * [ ]). */
function sheetName(name, used) {
  const base = String(name || 'Planilha').replace(/[:\\/?*[\]]/g, ' ').slice(0, 31).trim() || 'Planilha';
  let candidate = base;
  for (let i = 2; used.has(candidate.toLowerCase()); i++) candidate = `${base.slice(0, 27)} (${i})`;
  used.add(candidate.toLowerCase());
  return candidate;
}

/**
 * Grava um .xlsx em `out` (fluxo gravável, ex.: a resposta HTTP), respeitando o controle de fluxo.
 * sheets: [{ name, cols?: [larguras], header?: [células], rows: iterável de linhas }]
 *   Com header, a primeira linha é congelada e recebe filtro automático.
 * célula: texto | número | Date | { v, s: 'header'|'date'|'wrap'|'bold'|'title'|'int' }
 */
export async function writeXlsx(sheets, out, { title = 'Relatório', creator = 'CLEAN', maxRows = MAX_ROWS } = {}) {
  let failure = null;
  let finished;
  const done = new Promise((resolve) => {
    finished = resolve;
  });
  const zip = new Zip((err, data, final) => {
    if (err) {
      failure = err;
      finished();
      return;
    }
    out.write(Buffer.from(data));
    if (final) finished();
  });
  const drain = async () => {
    if (out.writableNeedDrain) await once(out, 'drain');
    if (failure) throw failure;
  };
  const addFile = (name, content) => {
    const file = new ZipDeflate(name, { level: 6 });
    zip.add(file);
    file.push(strToU8(content), true);
  };

  const used = new Set();
  const parts = []; // { name, autoFilter }
  for (const sheet of sheets) {
    const header = sheet.header || null;
    const limit = maxRows - (header ? 1 : 0);
    let part = null;
    let written = 0;
    let chunk = '';
    const open = () => {
      const index = parts.length + 1;
      const name = sheetName(parts.some((p) => p.base === sheet.name) ? `${sheet.name} (${parts.filter((p) => p.base === sheet.name).length + 1})` : sheet.name, used);
      part = { index, name, base: sheet.name, rows: 0, width: header ? header.length : 1, file: new ZipDeflate(`xl/worksheets/sheet${index}.xml`, { level: 6 }) };
      parts.push(part);
      zip.add(part.file);
      chunk = sheetHead(sheet);
      if (header) {
        chunk += rowXml(header.map((h) => (typeof h === 'object' ? h : { v: h, s: 'header' })), 1);
        part.rows = 1;
      }
      written = 0;
    };
    const close = () => {
      const filter = header && part.rows > 1 ? `<autoFilter ref="A1:${columnName(part.width - 1)}${part.rows}"/>` : '';
      part.autoFilter = filter ? `A1:${columnName(part.width - 1)}${part.rows}` : null;
      part.file.push(strToU8(`${chunk}</sheetData>${filter}</worksheet>`), true);
      chunk = '';
    };
    open();
    let pending = 0;
    for (const row of sheet.rows) {
      if (written >= limit) {
        close();
        await drain();
        open();
      }
      part.rows++;
      written++;
      if (row.length > part.width) part.width = row.length;
      chunk += rowXml(row, part.rows);
      if (++pending >= BATCH) {
        part.file.push(strToU8(chunk));
        chunk = '';
        pending = 0;
        await drain();
      }
    }
    close();
    await drain();
  }

  const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  addFile(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${parts
      .map((p) => `<Override PartName="/xl/worksheets/sheet${p.index}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`)
      .join('')}<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`,
  );
  addFile(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>',
  );
  addFile(
    'docProps/core.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${xmlEscape(title)}</dc:title><dc:creator>${xmlEscape(creator)}</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`,
  );
  addFile(
    'docProps/app.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>CLEAN</Application></Properties>',
  );
  const definedNames = parts
    .map((p, i) => (p.autoFilter ? `<definedName name="_xlnm._FilterDatabase" localSheetId="${i}" hidden="1">'${xmlEscape(p.name.replace(/'/g, "''"))}'!${p.autoFilter.replace(/([A-Z]+)(\d+)/g, '$$$1$$$2')}</definedName>` : ''))
    .join('');
  addFile(
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView/></bookViews><sheets>${parts
      .map((p) => `<sheet name="${xmlEscape(p.name)}" sheetId="${p.index}" r:id="rId${p.index}"/>`)
      .join('')}</sheets>${definedNames ? `<definedNames>${definedNames}</definedNames>` : ''}</workbook>`,
  );
  addFile(
    'xl/_rels/workbook.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${parts
      .map((p) => `<Relationship Id="rId${p.index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${p.index}.xml"/>`)
      .join('')}<Relationship Id="rId${parts.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`,
  );
  addFile('xl/styles.xml', STYLES);
  zip.end();
  await done;
  if (failure) throw failure;
}

/** Gera o .xlsx inteiro em memória (útil para arquivos pequenos e testes). */
export async function buildXlsx(sheets, options) {
  const chunks = [];
  const sink = {
    writableNeedDrain: false,
    write(chunk) {
      chunks.push(chunk);
      return true;
    },
  };
  await writeXlsx(sheets, sink, options);
  return Buffer.concat(chunks);
}
