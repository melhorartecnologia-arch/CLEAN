// Decodificação de arquivos de texto com detecção de codificação (UTF-8, UTF-16, Windows-1252).

const utf8Fatal = new TextDecoder('utf-8', { fatal: true });
const decoders = new Map();

export function decoderFor(label) {
  let d = decoders.get(label);
  if (!d) {
    try {
      d = new TextDecoder(label);
    } catch {
      d = new TextDecoder('windows-1252');
    }
    decoders.set(label, d);
  }
  return d;
}

/** Nome de codificação do TextDecoder para uma página de código do Windows. */
export function codePageLabel(cp) {
  const map = {
    65001: 'utf-8',
    1200: 'utf-16le',
    1201: 'utf-16be',
    874: 'windows-874',
    932: 'shift_jis',
    936: 'gbk',
    949: 'euc-kr',
    950: 'big5',
    10000: 'macintosh',
    20127: 'us-ascii',
    28591: 'iso-8859-1',
    28592: 'iso-8859-2',
    28605: 'iso-8859-15',
  };
  if (map[cp]) return map[cp];
  if (cp >= 1250 && cp <= 1258) return `windows-${cp}`;
  return 'windows-1252';
}

/** Proporção de bytes nulos em posições pares/ímpares (indica UTF-16 sem BOM). */
function utf16Guess(buf) {
  const len = Math.min(buf.length, 4096) & ~1;
  if (len < 16) return null;
  let evenZeros = 0;
  let oddZeros = 0;
  let printableLe = 0;
  let printableBe = 0;
  const printable = (b) => b === 9 || b === 10 || b === 13 || (b >= 32 && b !== 127);
  for (let i = 0; i < len; i += 2) {
    if (buf[i] === 0) evenZeros++;
    if (buf[i + 1] === 0) oddZeros++;
    if (buf[i + 1] === 0 && printable(buf[i])) printableLe++;
    if (buf[i] === 0 && printable(buf[i + 1])) printableBe++;
  }
  const pairs = len / 2;
  if (oddZeros / pairs > 0.3 && evenZeros / pairs < 0.05 && printableLe / oddZeros > 0.9) return 'utf-16le';
  if (evenZeros / pairs > 0.3 && oddZeros / pairs < 0.05 && printableBe / evenZeros > 0.9) return 'utf-16be';
  return null;
}

/**
 * Remove uma sequência UTF-8 incompleta no fim do buffer (leitura parcial de um arquivo grande),
 * para que o corte no meio de um caractere não faça o texto todo parecer Windows-1252.
 */
function trimPartialUtf8(buf) {
  let i = buf.length - 1;
  let continuation = 0;
  while (i >= 0 && continuation < 3 && (buf[i] & 0xc0) === 0x80) {
    i--;
    continuation++;
  }
  if (i < 0) return buf;
  const lead = buf[i];
  const needed = lead >= 0xf0 ? 4 : lead >= 0xe0 ? 3 : lead >= 0xc0 ? 2 : 1;
  return needed > continuation + 1 ? buf.subarray(0, i) : buf;
}

/** Decodifica um buffer de texto escolhendo a codificação mais provável. */
export function decodeText(buf, { partial = false } = {}) {
  if (buf.length >= 3 && buf[0] === 0xef && buf[1] === 0xbb && buf[2] === 0xbf) {
    return decoderFor('utf-8').decode(buf.subarray(3));
  }
  if (buf.length >= 2 && buf[0] === 0xff && buf[1] === 0xfe) return decoderFor('utf-16le').decode(buf.subarray(2));
  if (buf.length >= 2 && buf[0] === 0xfe && buf[1] === 0xff) return decoderFor('utf-16be').decode(buf.subarray(2));
  const utf16 = utf16Guess(buf);
  if (utf16) return decoderFor(utf16).decode(buf);
  try {
    return utf8Fatal.decode(partial ? trimPartialUtf8(buf) : buf);
  } catch {
    return decoderFor('windows-1252').decode(buf);
  }
}

/** Indica se o início do arquivo parece texto (e não binário). */
export function looksLikeText(buf) {
  const len = Math.min(buf.length, 8192);
  if (len === 0) return false;
  if ((buf[0] === 0xff && buf[1] === 0xfe) || (buf[0] === 0xfe && buf[1] === 0xff)) return true;
  if (utf16Guess(buf)) return true;
  let control = 0;
  for (let i = 0; i < len; i++) {
    const b = buf[i];
    if (b === 0) return false;
    if (b < 32 && b !== 9 && b !== 10 && b !== 13 && b !== 12 && b !== 27) control++;
  }
  return control / len < 0.02;
}

/**
 * Extrai sequências legíveis de um binário (8 bits e UTF-16LE), como o utilitário "strings".
 * Usado como alternativa para formatos binários antigos que não conseguimos interpretar.
 */
export function binaryStrings(buf, minLength = 4) {
  const parts = [];
  const hasWord = /[\p{L}\p{N}]{2}/u;
  // Sequências de 8 bits (Windows-1252 imprimível)
  let start = -1;
  for (let i = 0; i <= buf.length; i++) {
    const b = i < buf.length ? buf[i] : 0;
    if ((b >= 32 && b < 127) || b >= 160 || b === 9) {
      if (start === -1) start = i;
    } else {
      if (start !== -1 && i - start >= minLength) {
        const s = decoderFor('windows-1252').decode(buf.subarray(start, i));
        if (hasWord.test(s)) parts.push(s);
      }
      start = -1;
    }
  }
  // Sequências UTF-16LE (somente alfabetos latino/grego/cirílico e pontuação comum, para não
  // confundir texto de 8 bits lido de dois em dois bytes com ideogramas).
  for (const offset of [0, 1]) {
    start = -1;
    for (let i = offset; i < buf.length; i += 2) {
      const code = i + 1 < buf.length ? buf[i] | (buf[i + 1] << 8) : 0;
      const printable =
        (code >= 32 && code < 127) || (code >= 160 && code < 0x0590) || (code >= 0x2000 && code <= 0x20cf) || code === 9;
      if (printable) {
        if (start === -1) start = i;
        continue;
      }
      if (start !== -1 && (i - start) / 2 >= minLength) {
        const s = buf.toString('utf16le', start, i);
        if (hasWord.test(s)) parts.push(s);
      }
      start = -1;
    }
    if (start !== -1 && (buf.length - start) / 2 >= minLength) {
      const s = buf.toString('utf16le', start, buf.length - ((buf.length - start) % 2));
      if (hasWord.test(s)) parts.push(s);
    }
  }
  return parts.join('\n');
}
