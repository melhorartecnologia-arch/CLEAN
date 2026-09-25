// Gerador mínimo de planilhas .xlsx (Office Open XML), sem dependências além do fflate.
import { zipSync, strToU8 } from 'fflate';

const MAX_CELL = 32767;

/** Remove caracteres proibidos em XML 1.0 e escapa os especiais. */
export function xmlEscape(value) {
  return String(value)
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g, '')
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

function sheetXml(sheet) {
  const rows = sheet.rows || [];
  const width = Math.max(1, ...rows.map((r) => r.length), (sheet.cols || []).length);
  const parts = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'];
  parts.push(`<dimension ref="A1:${columnName(width - 1)}${Math.max(rows.length, 1)}"/>`);
  if (sheet.freezeRow) {
    const top = sheet.freezeRow + 1;
    parts.push(`<sheetViews><sheetView workbookViewId="0"><pane ySplit="${sheet.freezeRow}" topLeftCell="A${top}" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A${top}" sqref="A${top}"/></sheetView></sheetViews>`);
  } else {
    parts.push('<sheetViews><sheetView workbookViewId="0"/></sheetViews>');
  }
  if (sheet.cols && sheet.cols.length) {
    parts.push('<cols>');
    sheet.cols.forEach((w, i) => parts.push(`<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`));
    parts.push('</cols>');
  }
  parts.push('<sheetData>');
  rows.forEach((row, r) => {
    const cells = row.map((cell, c) => cellXml(`${columnName(c)}${r + 1}`, cell)).join('');
    parts.push(`<row r="${r + 1}">${cells}</row>`);
  });
  parts.push('</sheetData>');
  if (sheet.autoFilter && rows.length > 0) parts.push(`<autoFilter ref="${sheet.autoFilter}"/>`);
  parts.push('</worksheet>');
  return parts.join('');
}

/** Nome de planilha válido no Excel (até 31 caracteres, sem : \ / ? * [ ]). */
function sheetName(name, used) {
  let base = String(name || 'Planilha').replace(/[:\\/?*[\]]/g, ' ').slice(0, 31).trim() || 'Planilha';
  let candidate = base;
  for (let i = 2; used.has(candidate.toLowerCase()); i++) candidate = `${base.slice(0, 28)} ${i}`;
  used.add(candidate.toLowerCase());
  return candidate;
}

/**
 * sheets: [{ name, cols: [larguras], rows: [[célula]], freezeRow?: n, autoFilter?: 'A1:D10' }]
 * célula: texto | número | Date | { v, s: 'header'|'date'|'wrap'|'bold'|'title'|'int' }
 */
export function buildXlsx(sheets, { title = 'Relatório', creator = 'CLEAN' } = {}) {
  const used = new Set();
  const names = sheets.map((s) => sheetName(s.name, used));
  const files = {};
  const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  files['[Content_Types].xml'] = strToU8(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${sheets
      .map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`)
      .join('')}<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`,
  );
  files['_rels/.rels'] = strToU8(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>',
  );
  files['docProps/core.xml'] = strToU8(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${xmlEscape(title)}</dc:title><dc:creator>${xmlEscape(creator)}</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`,
  );
  files['docProps/app.xml'] = strToU8(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>CLEAN</Application></Properties>',
  );
  const definedNames = sheets
    .map((s, i) => (s.autoFilter && (s.rows || []).length ? `<definedName name="_xlnm._FilterDatabase" localSheetId="${i}" hidden="1">'${names[i].replace(/'/g, "''")}'!${s.autoFilter.replace(/([A-Z]+)(\d+)/g, '$$$1$$$2')}</definedName>` : ''))
    .join('');
  files['xl/workbook.xml'] = strToU8(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView/></bookViews><sheets>${names
      .map((n, i) => `<sheet name="${xmlEscape(n)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`)
      .join('')}</sheets>${definedNames ? `<definedNames>${definedNames}</definedNames>` : ''}</workbook>`,
  );
  files['xl/_rels/workbook.xml.rels'] = strToU8(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets
      .map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`)
      .join('')}<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`,
  );
  files['xl/styles.xml'] = strToU8(STYLES);
  sheets.forEach((sheet, i) => {
    files[`xl/worksheets/sheet${i + 1}.xml`] = strToU8(sheetXml(sheet));
  });
  return Buffer.from(zipSync(files, { level: 6 }));
}
