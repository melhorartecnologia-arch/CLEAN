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

/** MP3 sem marca ID3: cabeçalho de quadro MPEG válido (sincronismo, versão, camada e taxas). */
function mp3Frame(buf) {
  if (buf.length < 4 || buf[0] !== 0xff || (buf[1] & 0xe0) !== 0xe0) return false;
  const version = (buf[1] >> 3) & 0x03;
  const layer = (buf[1] >> 1) & 0x03;
  const bitrate = (buf[2] >> 4) & 0x0f;
  const rate = (buf[2] >> 2) & 0x03;
  return version !== 1 && layer !== 0 && bitrate !== 0x0f && bitrate !== 0 && rate !== 3;
}

/** Arquivo ZIP: o primeiro item diz se é um documento do Office, um Java ou um Android. */
function zipKind(buf) {
  if (buf.length < 30) return { category: 'archive', format: 'ZIP' };
  const nameLength = buf.readUInt16LE(26);
  const name = ascii(buf, 30, 30 + nameLength);
  if (name === '[Content_Types].xml' || /^(_rels|docProps|word|xl|ppt)\//.test(name)) return { category: null, format: 'Documento do Office (OOXML)' };
  if (name === 'mimetype') return { category: null, format: 'OpenDocument ou EPUB' };
  if (name === 'AndroidManifest.xml' || name === 'classes.dex' || name === 'resources.arsc') return { category: 'executable', format: 'Aplicativo Android (APK)' };
  if (name.startsWith('META-INF/')) return { category: 'executable', format: 'Java (JAR)' };
  return { category: 'archive', format: 'ZIP' };
}

/**
 * Categoria e formato pelo conteúdo ({ category, format }), ou null se o formato não for
 * reconhecido. category null: formato reconhecido, mas sem categoria confiável (ex.: Office).
 * size: tamanho do arquivo (confere o BMP).
 */
export function detectType(buf, size = null) {
  if (!buf || buf.length < 4) return null;
  // Vídeo e áudio
  if (ascii(buf, 4, 8) === 'ftyp') {
    const brand = ascii(buf, 8, 12);
    if (FTYP_AUDIO.has(brand)) return { category: 'audio', format: 'Áudio MPEG-4 (M4A)' };
    if (FTYP_IMAGE.has(brand)) return { category: 'image', format: 'Imagem HEIF/HEIC' };
    if (brand.startsWith('3g')) return { category: 'video', format: 'Vídeo 3GP' };
    if (brand === 'qt  ') return { category: 'video', format: 'Vídeo QuickTime (MOV)' };
    return { category: 'video', format: 'Vídeo MPEG-4 (MP4)' };
  }
  if (startsWith(buf, [0x1a, 0x45, 0xdf, 0xa3])) return { category: 'video', format: 'Vídeo Matroska/WebM' };
  if (ascii(buf, 0, 4) === 'RIFF') {
    const kind = ascii(buf, 8, 12);
    if (kind === 'AVI ') return { category: 'video', format: 'Vídeo AVI' };
    if (kind === 'WAVE') return { category: 'audio', format: 'Áudio WAV' };
    if (kind === 'WEBP') return { category: 'image', format: 'Imagem WebP' };
  }
  if (startsWith(buf, [0x30, 0x26, 0xb2, 0x75, 0x8e, 0x66, 0xcf, 0x11])) return { category: 'video', format: 'Windows Media (WMV/WMA)' };
  if (ascii(buf, 0, 3) === 'FLV' && buf[3] === 0x01) return { category: 'video', format: 'Vídeo Flash (FLV)' };
  if (startsWith(buf, [0x00, 0x00, 0x01, 0xba]) || startsWith(buf, [0x00, 0x00, 0x01, 0xb3])) return { category: 'video', format: 'Vídeo MPEG' };
  if (buf.length > 376 && buf[0] === 0x47 && buf[188] === 0x47 && buf[376] === 0x47) return { category: 'video', format: 'Vídeo MPEG-TS' };
  if (ascii(buf, 0, 3) === 'ID3' || mp3Frame(buf)) return { category: 'audio', format: 'Áudio MP3' };
  if (ascii(buf, 0, 4) === 'fLaC') return { category: 'audio', format: 'Áudio FLAC' };
  if (ascii(buf, 0, 4) === 'OggS') return { category: 'audio', format: 'Áudio/vídeo Ogg' };
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
  if (ascii(buf, 257, 262) === 'ustar') return { category: 'archive', format: 'TAR' };
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
