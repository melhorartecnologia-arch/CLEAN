// Rich Text Format (.rtf): texto e metadados do grupo \info (autor e "operator" = última alteração).
import { decoderFor, codePageLabel } from './text.js';

// Destinos cujo conteúdo não é texto do documento.
const SKIP = new Set([
  'fonttbl', 'colortbl', 'stylesheet', 'listtable', 'listoverridetable', 'revtbl', 'rsidtbl', 'pict',
  'object', 'objdata', 'themedata', 'colorschememapping', 'latentstyles', 'datastore', 'xmlnstbl',
  'fldinst', 'filetbl', 'generator', 'pgdsctbl', 'mmathPr', 'sp', 'nonshppict', 'template',
  'ftnsep', 'ftnsepc', 'aftnsep', 'aftnsepc', 'protusertbl', 'passwordhash', 'userprops', 'docvar',
]);
// Destinos marcados com \* que contêm texto visível.
const KEEP_STARRED = new Set(['shpinst', 'shptxt', 'annotation', 'footnote']);
const INFO_FIELDS = { author: 'author', operator: 'lastModifiedBy', title: 'title', company: 'company' };
const INFO_DATES = { creatim: 'created', revtim: 'modified' };
const SYMBOLS = {
  par: '\n', line: '\n', sect: '\n', page: '\n', row: '\n', tab: '\t', cell: '\t', emdash: '—',
  endash: '–', bullet: '•', lquote: '‘', rquote: '’', ldblquote: '“', rdblquote: '”', emspace: ' ',
  enspace: ' ', qmspace: ' ',
};

export function isRtf(buf) {
  return buf.length > 5 && buf.toString('latin1', 0, 5) === '{\\rtf';
}

export function rtfExtract(buf) {
  const src = buf.toString('latin1');
  const n = src.length;
  let codePage = 'windows-1252';
  const meta = {};
  const main = { text: '', bytes: [] };
  const stack = [];
  let state = { skip: false, uc: 1, sink: main, field: null, date: null, info: false };
  let groupStart = false;
  let starred = false;
  let skipChars = 0;
  const wordRe = /([a-zA-Z]{1,32})(-?\d{1,10})? ?/y;
  const runRe = /[^\\{}\r\n]+/y;

  const flush = (sink) => {
    if (sink && sink.bytes.length) {
      sink.text += decoderFor(codePage).decode(Uint8Array.from(sink.bytes));
      sink.bytes.length = 0;
    }
  };
  const active = () => !state.skip && state.sink;
  const putBytes = (str) => {
    if (!active()) return;
    for (let k = 0; k < str.length; k++) state.sink.bytes.push(str.charCodeAt(k));
  };
  const putText = (t) => {
    if (!active()) return;
    flush(state.sink);
    state.sink.text += t;
  };

  let i = 0;
  while (i < n) {
    const ch = src.charCodeAt(i);
    if (ch === 123 /* { */) {
      stack.push(state);
      state = { ...state, field: null, date: null };
      groupStart = true;
      starred = false;
      skipChars = 0;
      i++;
      continue;
    }
    if (ch === 125 /* } */) {
      if (state.field) {
        flush(state.sink);
        const value = state.sink.text.trim();
        if (value) {
          if (state.field === 'author' && meta.author) meta.lastModifiedBy ??= value;
          else meta[state.field] ??= value;
        }
      }
      if (state.date && state.date.yr) {
        const d = state.date;
        const date = new Date(d.yr, (d.mo || 1) - 1, d.dy || 1, d.hr || 0, d.min || 0);
        if (!Number.isNaN(date.getTime())) meta[d.field] ??= date.toISOString();
      }
      state = stack.pop() || state;
      groupStart = false;
      skipChars = 0;
      i++;
      continue;
    }
    if (ch === 13 || ch === 10) {
      i++;
      continue;
    }
    if (ch !== 92 /* \ */) {
      runRe.lastIndex = i;
      const run = runRe.exec(src)[0];
      i += run.length;
      groupStart = false;
      let start = 0;
      if (skipChars > 0) {
        start = Math.min(skipChars, run.length);
        skipChars -= start;
      }
      if (start < run.length) putBytes(run.slice(start));
      continue;
    }
    // Barra invertida: símbolo de controle ou palavra de controle
    const next = src[i + 1];
    if (next === "'") {
      const byte = parseInt(src.substr(i + 2, 2), 16);
      i += 4;
      groupStart = false;
      if (skipChars > 0) {
        skipChars--;
        continue;
      }
      if (!Number.isNaN(byte) && active()) state.sink.bytes.push(byte);
      continue;
    }
    if (next === '\\' || next === '{' || next === '}') {
      i += 2;
      groupStart = false;
      if (skipChars > 0) skipChars--;
      else putBytes(next);
      continue;
    }
    if (next === '*') {
      i += 2;
      if (groupStart) starred = true;
      continue;
    }
    if (next === '~' || next === '_' || next === '-' || next === '\r' || next === '\n') {
      i += 2;
      groupStart = false;
      if (next === '~') putText('\u00a0');
      else if (next === '_') putText('-');
      else if (next !== '-') putText('\n');
      continue;
    }
    wordRe.lastIndex = i + 1;
    const m = wordRe.exec(src);
    if (!m) {
      i += 2;
      continue;
    }
    i += 1 + m[0].length;
    const word = m[1];
    const param = m[2] !== undefined ? parseInt(m[2], 10) : null;
    const atGroupStart = groupStart;
    groupStart = false;

    if (word === 'u' && param !== null) {
      putText(String.fromCharCode(param < 0 ? param + 65536 : param));
      skipChars = state.uc;
      continue;
    }
    skipChars = 0;
    if (word === 'bin') {
      i += Math.max(0, param || 0);
      continue;
    }
    if (word === 'uc' && param !== null) {
      state.uc = param;
      continue;
    }
    if (word === 'ansicpg' && param) {
      flush(main);
      codePage = codePageLabel(param);
      continue;
    }
    if (atGroupStart) {
      if (starred && !KEEP_STARRED.has(word)) {
        state.skip = true;
        continue;
      }
      if (SKIP.has(word)) {
        state.skip = true;
        continue;
      }
      if (word === 'info') {
        state.info = true;
        state.sink = null;
        continue;
      }
      if (state.info) {
        if (INFO_FIELDS[word]) {
          state.field = INFO_FIELDS[word];
          state.sink = { text: '', bytes: [] };
        } else if (INFO_DATES[word]) {
          state.date = { field: INFO_DATES[word] };
        } else {
          state.sink = null;
        }
        continue;
      }
    }
    if (state.date && ['yr', 'mo', 'dy', 'hr', 'min'].includes(word)) {
      state.date[word] = param;
      continue;
    }
    if (SYMBOLS[word]) putText(SYMBOLS[word]);
  }
  flush(main);
  return { text: main.text, metadata: meta };
}
