// Tipo real de um arquivo pelos primeiros bytes (assinatura), para encontrar arquivos renomeados
// (ex.: um vídeo salvo como "relatorio.pdf"). Só formatos com assinatura inequívoca: textos,
// scripts e documentos antigos do Office (o mesmo contêiner de outros formatos) ficam de fora.

/** Bytes lidos do início de cada arquivo. */
export const HEADER_BYTES = 512;

const ascii = (buf, start, end) => buf.toString('latin1', start, Math.min(end, buf.length));
const startsWith = (buf, bytes, offset = 0) => buf.length >= offset + bytes.length && bytes.every((b, i) => buf[offset + i] === b);

// Marcas do contêiner ISO (MP4, MOV, 3GP, M4A, HEIC...): "ftyp" no byte 4 e a marca principal no 8.
const FTYP_AUDIO = new Set(['M4A ', 'M4B ', 'M4P ', 'F4A ', 'F4B ']);
const FTYP_IMAGE = new Set(['heic', 'heix', 'hevc', 'hevx', 'heim', 'heis', 'mif1', 'msf1', 'avif', 'avis']);

/**
 * O início do arquivo parece texto (sem caracteres de controle além de tabulação, quebras de linha
 * e ESC)? Nesse caso, só o PDF é reconhecido: assinaturas curtas em texto ("MZ", "ID3", "BZh"...)
 * seriam coincidência.
 */
function looksLikeText(buf) {
  for (const b of buf) if (b < 0x09 || (b > 0x0d && b < 0x20 && b !== 0x1b)) return false;
  return true;
}

/**
 * MP3 sem marca ID3: cabeçalho de quadro MPEG válido (sincronismo, versão, camada II ou III e
 * taxas). A camada I fica de fora: "FF FE" é também o início de um texto em UTF-16.
 */
function mp3Frame(buf) {
  if (buf.length < 4 || buf[0] !== 0xff || (buf[1] & 0xe0) !== 0xe0) return false;
  const version = (buf[1] >> 3) & 0x03;
  const layer = (buf[1] >> 1) & 0x03;
  const bitrate = (buf[2] >> 4) & 0x0f;
  const rate = (buf[2] >> 2) & 0x03;
  return version !== 1 && (layer === 1 || layer === 2) && bitrate !== 0x0f && bitrate !== 0 && rate !== 3;
}

/** Marca ID3 (MP3 e outros áudios): versão 2.2 a 2.4 e tamanho em bytes de 7 bits. */
const id3 = (buf) => ascii(buf, 0, 3) === 'ID3' && buf.length >= 10 && [2, 3, 4].includes(buf[3]) && buf[4] !== 0xff && [6, 7, 8, 9].every((i) => buf[i] < 0x80);

/** Arquivo ZIP: o primeiro item diz se é um documento do Office, um Java ou um Android. */
function zipKind(buf) {
  const zip = { category: 'archive', format: 'ZIP', container: 'zip' };
  if (buf.length < 30) return zip;
  const nameLength = buf.readUInt16LE(26);
  const name = ascii(buf, 30, 30 + nameLength);
  if (name === '[Content_Types].xml' || /^(_rels|docProps|word|xl|ppt)\//.test(name)) return { category: null, format: 'Documento do Office (OOXML)', container: 'zip' };
  if (name === 'mimetype') return { category: null, format: 'OpenDocument ou EPUB', container: 'zip' };
  if (name === 'AndroidManifest.xml' || name === 'classes.dex' || name === 'resources.arsc') return { category: 'executable', format: 'Aplicativo Android (APK)', container: 'zip' };
  if (name.startsWith('META-INF/')) return { category: 'executable', format: 'Java (JAR)', container: 'zip' };
  return zip;
}

/**
 * Categoria e formato pelo conteúdo ({ category, format, container? }), ou null se o formato não
 * for reconhecido. category null: formato reconhecido, mas sem categoria confiável (ex.: Office).
 * container: formato que outros tipos também usam por dentro (ver sameContainer). size: tamanho do
 * arquivo (confere o BMP).
 */
export function detectType(buf, size = null) {
  if (!buf || buf.length < 4) return null;
  if (looksLikeText(buf)) return ascii(buf, 0, 5) === '%PDF-' ? { category: 'document', format: 'PDF' } : null;
  // Vídeo e áudio
  if (ascii(buf, 4, 8) === 'ftyp') {
    const brand = ascii(buf, 8, 12);
    if (FTYP_AUDIO.has(brand)) return { category: 'audio', format: 'Áudio MPEG-4 (M4A)', container: 'mp4' };
    if (FTYP_IMAGE.has(brand)) return { category: 'image', format: 'Imagem HEIF/HEIC' };
    if (brand.startsWith('3g')) return { category: 'video', format: 'Vídeo 3GP', container: 'mp4' };
    if (brand === 'qt  ') return { category: 'video', format: 'Vídeo QuickTime (MOV)', container: 'mp4' };
    return { category: 'video', format: 'Vídeo MPEG-4 (MP4)', container: 'mp4' };
  }
  if (startsWith(buf, [0x1a, 0x45, 0xdf, 0xa3])) return { category: 'video', format: 'Vídeo Matroska/WebM', container: 'matroska' };
  if (ascii(buf, 0, 4) === 'RIFF') {
    const kind = ascii(buf, 8, 12);
    if (kind === 'AVI ') return { category: 'video', format: 'Vídeo AVI' };
    if (kind === 'WAVE') return { category: 'audio', format: 'Áudio WAV' };
    if (kind === 'WEBP') return { category: 'image', format: 'Imagem WebP' };
  }
  if (startsWith(buf, [0x30, 0x26, 0xb2, 0x75, 0x8e, 0x66, 0xcf, 0x11])) return { category: 'video', format: 'Windows Media (WMV/WMA)', container: 'asf' };
  if (ascii(buf, 0, 3) === 'FLV' && buf[3] === 0x01) return { category: 'video', format: 'Vídeo Flash (FLV)' };
  if (startsWith(buf, [0x00, 0x00, 0x01, 0xba]) || startsWith(buf, [0x00, 0x00, 0x01, 0xb3])) return { category: 'video', format: 'Vídeo MPEG' };
  if (buf.length > 376 && buf[0] === 0x47 && buf[188] === 0x47 && buf[376] === 0x47) return { category: 'video', format: 'Vídeo MPEG-TS' };
  if (id3(buf) || mp3Frame(buf)) return { category: 'audio', format: 'Áudio MP3' };
  if (ascii(buf, 0, 4) === 'fLaC') return { category: 'audio', format: 'Áudio FLAC' };
  if (ascii(buf, 0, 4) === 'OggS') return { category: 'audio', format: 'Áudio/vídeo Ogg', container: 'ogg' };
  if (ascii(buf, 0, 4) === 'MThd') return { category: 'audio', format: 'MIDI' };
  if (ascii(buf, 0, 6) === '#!AMR') return { category: 'audio', format: 'Áudio AMR' };
  // Imagens
  if (startsWith(buf, [0xff, 0xd8, 0xff])) return { category: 'image', format: 'Imagem JPEG' };
  if (startsWith(buf, [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])) return { category: 'image', format: 'Imagem PNG' };
  if (ascii(buf, 0, 6) === 'GIF87a' || ascii(buf, 0, 6) === 'GIF89a') return { category: 'image', format: 'Imagem GIF' };
  if (ascii(buf, 0, 2) === 'BM' && buf.length >= 6 && size !== null && buf.readUInt32LE(2) === size) return { category: 'image', format: 'Imagem BMP' };
  if (startsWith(buf, [0x49, 0x49, 0x2a, 0x00]) || startsWith(buf, [0x4d, 0x4d, 0x00, 0x2a])) return { category: 'image', format: 'Imagem TIFF (ou foto RAW)' };
  if (ascii(buf, 0, 4) === '8BPS') return { category: 'image', format: 'Photoshop (PSD)' };
  // Executáveis
  if (ascii(buf, 0, 2) === 'MZ') return { category: 'executable', format: 'Executável do Windows' };
  if (startsWith(buf, [0x7f, 0x45, 0x4c, 0x46])) return { category: 'executable', format: 'Executável Linux (ELF)' };
  // Compactados (o ZIP pode ser um documento do Office, um Java ou um Android)
  if (startsWith(buf, [0x50, 0x4b, 0x03, 0x04])) return zipKind(buf);
  if (ascii(buf, 0, 6) === 'Rar!\x1a\x07') return { category: 'archive', format: 'RAR' };
  if (startsWith(buf, [0x37, 0x7a, 0xbc, 0xaf, 0x27, 0x1c])) return { category: 'archive', format: '7-Zip' };
  if (startsWith(buf, [0x1f, 0x8b, 0x08])) return { category: 'archive', format: 'GZIP' };
  if (ascii(buf, 0, 3) === 'BZh' && buf[3] >= 0x31 && buf[3] <= 0x39) return { category: 'archive', format: 'BZIP2' };
  if (startsWith(buf, [0xfd, 0x37, 0x7a, 0x58, 0x5a, 0x00])) return { category: 'archive', format: 'XZ' };
  if (ascii(buf, 0, 4) === 'MSCF') return { category: 'archive', format: 'CAB' };
  if (startsWith(buf, [0x28, 0xb5, 0x2f, 0xfd])) return { category: 'archive', format: 'Zstandard' };
  if (ascii(buf, 257, 262) === 'ustar') return { category: 'archive', format: 'TAR', container: 'tar' };
  // Imagens de disco e máquinas virtuais
  if (ascii(buf, 0, 8) === 'vhdxfile') return { category: 'disk', format: 'Disco virtual (VHDX)' };
  if (ascii(buf, 0, 8) === 'conectix') return { category: 'disk', format: 'Disco virtual (VHD)' };
  if (ascii(buf, 0, 4) === 'KDMV') return { category: 'disk', format: 'Disco virtual (VMDK)' };
  if (ascii(buf, 0, 4) === 'QFI\xfb') return { category: 'disk', format: 'Disco virtual (QCOW)' };
  if (ascii(buf, 0, 40).includes('VirtualBox Disk Image')) return { category: 'disk', format: 'Disco virtual (VDI)' };
  // Documentos, e-mails e bancos de dados com assinatura própria
  if (ascii(buf, 0, 5) === '%PDF-') return { category: 'document', format: 'PDF' };
  if (ascii(buf, 0, 4) === '!BDN') return { category: 'email', format: 'Arquivo de dados do Outlook (PST/OST)' };
  if (ascii(buf, 0, 16) === 'SQLite format 3\x00') return { category: 'database', format: 'SQLite' };
  if (/^Standard (Jet|ACE) DB/.test(ascii(buf, 4, 19))) return { category: 'database', format: 'Microsoft Access' };
  return null;
}

/**
 * Extensões de outros tipos que usam o mesmo formato por dentro: não são arquivos renomeados (um
 * .m4a é MP4, um .mka é Matroska, um .docx ou um .jar é ZIP, um .ova é TAR).
 */
const SAME_CONTAINER = {
  mp4: ['.mp4', '.m4v', '.mov', '.3gp', '.3g2', '.f4v', '.m4a', '.m4b', '.m4p', '.m4r', '.aac', '.alac', '.3ga'],
  matroska: ['.mkv', '.webm', '.mka'],
  ogg: ['.ogg', '.oga', '.ogv', '.opus', '.spx'],
  asf: ['.asf', '.wmv', '.wma'],
  zip: [
    ...['.docx', '.docm', '.dotx', '.dotm', '.xlsx', '.xlsm', '.xlsb', '.xltx', '.xltm', '.pptx', '.pptm', '.ppsx', '.ppsm', '.potx', '.potm', '.vsdx', '.xps', '.oxps'],
    ...['.odt', '.ods', '.odp', '.odg', '.ott', '.ots', '.otp', '.epub', '.pages', '.numbers', '.key'],
    ...['.jar', '.war', '.ear', '.apk', '.aab', '.appx', '.appxbundle', '.msix', '.msixbundle', '.xpi', '.vsix', '.nupkg', '.kmz', '.3mf'],
  ],
  tar: ['.ova'],
};

/** O conteúdo reconhecido é o formato próprio da extensão (e não um arquivo renomeado)? */
export const sameContainer = (detected, extension) => Boolean(detected?.container && SAME_CONTAINER[detected.container]?.includes(extension));
