// Tipo real de um arquivo pelos primeiros bytes (assinatura), para encontrar arquivos renomeados
// (ex.: um vídeo salvo como "relatorio.pdf"). Só formatos com assinatura inequívoca: textos,
// scripts e documentos antigos do Office (o mesmo contêiner de outros formatos) ficam de fora, e as
// assinaturas curtas são conferidas com o que vem depois delas (cabeçalho PE, dois quadros de MP3,
// vários pacotes de MPEG-TS...).

/** Bytes lidos do início de cada arquivo (um bloco do disco: ler 512 bytes custaria o mesmo). */
export const HEADER_BYTES = 4096;

/**
 * Versão do reconhecimento: muda quando as assinaturas mudam. Faz parte dos critérios confirmados
 * nos agendamentos com exclusão automática (uma atualização pede nova confirmação).
 */
export const SIGNATURE_VERSION = 2;

const ascii = (buf, start, end) => buf.toString('latin1', start, Math.min(end, buf.length));
const startsWith = (buf, bytes, offset = 0) => buf.length >= offset + bytes.length && bytes.every((b, i) => buf[offset + i] === b);

// Marcas do contêiner ISO (MP4, MOV, 3GP, M4A, HEIC, CR3...): "ftyp" no byte 4 e a marca principal
// no 8. Marcas fora destas listas: formato MPEG-4 sem categoria (não é classificado).
const FTYP_VIDEO = new Set(['isom', 'iso2', 'iso3', 'iso4', 'iso5', 'iso6', 'iso7', 'iso8', 'iso9', 'mp41', 'mp42', 'avc1', 'M4V ', 'M4VH', 'M4VP', 'mmp4', 'f4v ', 'F4V ', 'XAVC', 'MSNV', 'dash', 'qt  ']);
const FTYP_AUDIO = new Set(['M4A ', 'M4B ', 'M4P ', 'F4A ', 'F4B ']);
const FTYP_IMAGE = new Set(['heic', 'heix', 'hevc', 'hevx', 'heim', 'heis', 'hevm', 'hevs', 'mif1', 'mif2', 'msf1', 'miaf', 'avif', 'avis', 'avio', 'crx ']);

/**
 * O início do arquivo parece texto (ASCII/UTF-8 sem caracteres de controle além de tabulação,
 * quebras de linha e ESC, ou UTF-16)? Nesse caso, só o PDF é reconhecido: assinaturas curtas num
 * texto ("MZ", "ID3", "BZh"...) seriam coincidência.
 */
function looksLikeText(buf) {
  let control = false;
  for (const b of buf) {
    if (b < 0x09 || (b > 0x0d && b < 0x20 && b !== 0x1b)) {
      control = true;
      break;
    }
  }
  if (!control) return true;
  // UTF-16 (com ou sem BOM): um byte de cada par é zero em quase todo o texto latino.
  const n = Math.min(buf.length, 512) & ~1;
  if (n < 16) return false;
  let evenZero = 0;
  let oddZero = 0;
  for (let i = 0; i < n; i += 2) {
    if (buf[i] === 0) evenZero++;
    if (buf[i + 1] === 0) oddZero++;
  }
  const pairs = n / 2;
  return (oddZero >= pairs * 0.9 && evenZero <= pairs * 0.05) || (evenZero >= pairs * 0.9 && oddZero <= pairs * 0.05);
}

// Quadros de MP3 (MPEG áudio, camadas II e III): taxas em kbit/s e frequências em Hz.
const MP3_BITRATES = {
  v1l2: [0, 32, 48, 56, 64, 80, 96, 112, 128, 160, 192, 224, 256, 320, 384],
  v1l3: [0, 32, 40, 48, 56, 64, 80, 96, 112, 128, 160, 192, 224, 256, 320],
  v2: [0, 8, 16, 24, 32, 40, 48, 56, 64, 80, 96, 112, 128, 144, 160],
};
const MP3_RATES = { 3: [44100, 48000, 32000], 2: [22050, 24000, 16000], 0: [11025, 12000, 8000] };

/** Tamanho do quadro de MP3 que começa em `at` (0 se não houver um cabeçalho válido ali). */
function mp3FrameLength(buf, at) {
  if (at + 4 > buf.length || buf[at] !== 0xff || (buf[at + 1] & 0xe0) !== 0xe0) return 0;
  const version = (buf[at + 1] >> 3) & 0x03; // 3: MPEG-1; 2: MPEG-2; 0: MPEG-2.5; 1: reservado
  const layer = (buf[at + 1] >> 1) & 0x03; // 1: camada III; 2: camada II (a camada I, "FF FE", é também o início de um texto UTF-16)
  const bitrate = (buf[at + 2] >> 4) & 0x0f;
  const rate = (buf[at + 2] >> 2) & 0x03;
  const padding = (buf[at + 2] >> 1) & 0x01;
  if (version === 1 || (layer !== 1 && layer !== 2) || bitrate === 0 || bitrate === 0x0f || rate === 3) return 0;
  const table = version === 3 ? (layer === 2 ? MP3_BITRATES.v1l2 : MP3_BITRATES.v1l3) : MP3_BITRATES.v2;
  const factor = layer === 1 && version !== 3 ? 72 : 144;
  return Math.floor((factor * table[bitrate] * 1000) / MP3_RATES[version][rate]) + padding;
}

/** MP3 sem marca ID3: dois quadros seguidos (um só cabeçalho seria coincidência em dados binários). */
function mp3Frames(buf) {
  const length = mp3FrameLength(buf, 0);
  return length > 0 && mp3FrameLength(buf, length) > 0;
}

/** Marca ID3 (MP3 e outros áudios): versão 2.2 a 2.4 e tamanho em bytes de 7 bits. */
const id3 = (buf) => ascii(buf, 0, 3) === 'ID3' && buf.length >= 10 && [2, 3, 4].includes(buf[3]) && buf[4] !== 0xff && [6, 7, 8, 9].every((i) => buf[i] < 0x80);

/** Pacotes de MPEG-TS (188 bytes) ou M2TS (192, com 4 bytes de horário antes): 5 ou mais seguidos. */
function transportStream(buf) {
  for (const [first, size] of [
    [0, 188],
    [4, 192],
  ]) {
    const packets = Math.floor((buf.length - first) / size);
    if (packets < 5) continue;
    let ok = true;
    for (let i = 0; i < packets && ok; i++) ok = buf[first + i * size] === 0x47;
    if (ok) return true;
  }
  return false;
}

/** Executável do Windows: "MZ" e a assinatura do cabeçalho novo (PE, NE, LE ou LX) no endereço indicado. */
function windowsExecutable(buf) {
  if (ascii(buf, 0, 2) !== 'MZ' || buf.length < 0x40) return false;
  const at = buf.readUInt32LE(0x3c);
  if (at < 0x40 || at + 4 > buf.length) return false;
  const sig = ascii(buf, at, at + 4);
  return sig === 'PE\0\0' || ['NE', 'LE', 'LX'].includes(sig.slice(0, 2));
}

/** Arquivo ZIP: o primeiro item diz se é um documento do Office, um Java ou um Android. */
function zipKind(buf) {
  const zip = { category: 'archive', format: 'ZIP', container: 'zip' };
  if (buf.length < 30) return zip;
  const nameLength = buf.readUInt16LE(26);
  const name = ascii(buf, 30, 30 + nameLength);
  if (name === '[Content_Types].xml' || /^(_rels|docProps|word|xl|ppt|customXml)\//.test(name)) return { category: null, format: 'Documento do Office (OOXML)', container: 'zip' };
  if (name === 'mimetype') return { category: null, format: 'OpenDocument ou EPUB', container: 'zip' };
  if (name === 'AndroidManifest.xml' || name === 'classes.dex' || name === 'resources.arsc') return { category: 'executable', format: 'Aplicativo Android (APK)', container: 'zip' };
  if (name.startsWith('META-INF/')) return { category: 'executable', format: 'Java (JAR)', container: 'zip' };
  return zip;
}

/**
 * Categoria e formato pelo conteúdo ({ category, format, container? }), ou null se o formato não
 * for reconhecido. category null: formato reconhecido, mas sem categoria confiável (ex.: Office,
 * MPEG-4 de marca desconhecida). container: formato que outros tipos também usam por dentro (ver
 * sameContainer). size: tamanho do arquivo (confere o BMP).
 */
export function detectType(buf, size = null) {
  if (!buf || buf.length < 4) return null;
  if (looksLikeText(buf)) return ascii(buf, 0, 5) === '%PDF-' ? { category: 'document', format: 'PDF' } : null;
  // Vídeo, áudio e imagens no contêiner MPEG-4 (a marca diz qual)
  if (ascii(buf, 4, 8) === 'ftyp') {
    const brand = ascii(buf, 8, 12);
    if (FTYP_AUDIO.has(brand)) return { category: 'audio', format: 'Áudio MPEG-4 (M4A)', container: 'mp4' };
    if (FTYP_IMAGE.has(brand)) return { category: 'image', format: brand === 'crx ' ? 'Foto RAW da Canon (CR3)' : 'Imagem HEIF/HEIC/AVIF', container: 'mp4' };
    if (brand.startsWith('3g')) return { category: 'video', format: 'Vídeo 3GP', container: 'mp4' };
    if (brand === 'qt  ') return { category: 'video', format: 'Vídeo QuickTime (MOV)', container: 'mp4' };
    if (FTYP_VIDEO.has(brand)) return { category: 'video', format: 'Vídeo MPEG-4 (MP4)', container: 'mp4' };
    return { category: null, format: `MPEG-4 (marca "${brand.trim()}")`, container: 'mp4' };
  }
  // QuickTime antigo, sem "ftyp": o primeiro átomo já é o filme, os dados ou o espaço reservado
  if (['moov', 'mdat', 'wide', 'pnot'].includes(ascii(buf, 4, 8))) return { category: 'video', format: 'Vídeo QuickTime (MOV)', container: 'mp4' };
  if (startsWith(buf, [0x1a, 0x45, 0xdf, 0xa3])) return { category: 'video', format: 'Vídeo Matroska/WebM', container: 'matroska' };
  if (ascii(buf, 0, 4) === 'RIFF') {
    const kind = ascii(buf, 8, 12);
    if (kind === 'AVI ') return { category: 'video', format: 'Vídeo AVI' };
    if (kind === 'WAVE') return { category: 'audio', format: 'Áudio WAV' };
    if (kind === 'WEBP') return { category: 'image', format: 'Imagem WebP' };
  }
  if (startsWith(buf, [0x30, 0x26, 0xb2, 0x75, 0x8e, 0x66, 0xcf, 0x11])) return { category: 'video', format: 'Windows Media (WMV/WMA)', container: 'asf' };
  if (ascii(buf, 0, 3) === 'FLV' && buf[3] === 0x01 && buf.length >= 9 && buf.readUInt32BE(5) === 9) return { category: 'video', format: 'Vídeo Flash (FLV)' };
  if (startsWith(buf, [0x00, 0x00, 0x01, 0xba]) || startsWith(buf, [0x00, 0x00, 0x01, 0xb3])) return { category: 'video', format: 'Vídeo MPEG' };
  if (id3(buf)) return { category: 'audio', format: 'Áudio MP3' };
  if (ascii(buf, 0, 4) === 'fLaC') return { category: 'audio', format: 'Áudio FLAC' };
  if (ascii(buf, 0, 4) === 'OggS' && buf[4] === 0) return { category: 'audio', format: 'Áudio/vídeo Ogg', container: 'ogg' };
  if (ascii(buf, 0, 4) === 'MThd' && buf.length >= 8 && buf.readUInt32BE(4) === 6) return { category: 'audio', format: 'MIDI' };
  if (ascii(buf, 0, 6) === '#!AMR') return { category: 'audio', format: 'Áudio AMR' };
  // Imagens
  if (startsWith(buf, [0xff, 0xd8, 0xff])) return { category: 'image', format: 'Imagem JPEG' };
  if (startsWith(buf, [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])) return { category: 'image', format: 'Imagem PNG' };
  if (ascii(buf, 0, 6) === 'GIF87a' || ascii(buf, 0, 6) === 'GIF89a') return { category: 'image', format: 'Imagem GIF' };
  if (ascii(buf, 0, 2) === 'BM' && buf.length >= 6 && size !== null && buf.readUInt32LE(2) === size) return { category: 'image', format: 'Imagem BMP' };
  if (startsWith(buf, [0x49, 0x49, 0x2a, 0x00]) || startsWith(buf, [0x4d, 0x4d, 0x00, 0x2a])) return { category: 'image', format: 'Imagem TIFF (ou foto RAW)' };
  if (ascii(buf, 0, 4) === '8BPS' && (buf[5] === 1 || buf[5] === 2) && buf[4] === 0) return { category: 'image', format: 'Photoshop (PSD)' };
  // Executáveis
  if (windowsExecutable(buf)) return { category: 'executable', format: 'Executável do Windows' };
  if (startsWith(buf, [0x7f, 0x45, 0x4c, 0x46]) && [1, 2].includes(buf[4]) && [1, 2].includes(buf[5]) && buf[6] === 1) return { category: 'executable', format: 'Executável Linux (ELF)' };
  // Compactados (o ZIP pode ser um documento do Office, um Java ou um Android)
  if (startsWith(buf, [0x50, 0x4b, 0x03, 0x04])) return zipKind(buf);
  if (ascii(buf, 0, 6) === 'Rar!\x1a\x07') return { category: 'archive', format: 'RAR' };
  if (startsWith(buf, [0x37, 0x7a, 0xbc, 0xaf, 0x27, 0x1c])) return { category: 'archive', format: '7-Zip' };
  if (startsWith(buf, [0x1f, 0x8b, 0x08])) return { category: 'archive', format: 'GZIP', container: 'compressed' };
  if (ascii(buf, 0, 3) === 'BZh' && buf[3] >= 0x31 && buf[3] <= 0x39 && (startsWith(buf, [0x31, 0x41, 0x59, 0x26, 0x53, 0x59], 4) || startsWith(buf, [0x17, 0x72, 0x45, 0x38, 0x50, 0x90], 4))) {
    return { category: 'archive', format: 'BZIP2', container: 'compressed' };
  }
  if (startsWith(buf, [0xfd, 0x37, 0x7a, 0x58, 0x5a, 0x00])) return { category: 'archive', format: 'XZ', container: 'compressed' };
  if (ascii(buf, 0, 4) === 'MSCF' && buf.length >= 8 && buf.readUInt32LE(4) === 0) return { category: 'archive', format: 'CAB' };
  if (startsWith(buf, [0x28, 0xb5, 0x2f, 0xfd])) return { category: 'archive', format: 'Zstandard', container: 'compressed' };
  if (ascii(buf, 257, 262) === 'ustar') return { category: 'archive', format: 'TAR', container: 'tar' };
  // Imagens de disco e máquinas virtuais
  if (ascii(buf, 0, 8) === 'vhdxfile') return { category: 'disk', format: 'Disco virtual (VHDX)' };
  if (ascii(buf, 0, 8) === 'conectix') return { category: 'disk', format: 'Disco virtual (VHD)' };
  if (ascii(buf, 0, 4) === 'KDMV') return { category: 'disk', format: 'Disco virtual (VMDK)' };
  if (ascii(buf, 0, 4) === 'QFI\xfb') return { category: 'disk', format: 'Disco virtual (QCOW)' };
  if (ascii(buf, 0, 72).includes('VirtualBox Disk Image')) return { category: 'disk', format: 'Disco virtual (VDI)' };
  // Documentos, e-mails e bancos de dados com assinatura própria
  if (ascii(buf, 0, 5) === '%PDF-') return { category: 'document', format: 'PDF' };
  if (ascii(buf, 0, 4) === '!BDN') return { category: 'email', format: 'Arquivo de dados do Outlook (PST/OST)' };
  if (ascii(buf, 0, 16) === 'SQLite format 3\x00') return { category: 'database', format: 'SQLite' };
  if (/^Standard (Jet|ACE) DB/.test(ascii(buf, 4, 19))) return { category: 'database', format: 'Microsoft Access' };
  // Assinaturas fracas, conferidas por último (várias repetições seguidas)
  if (transportStream(buf)) return { category: 'video', format: 'Vídeo MPEG-TS' };
  if (mp3Frames(buf)) return { category: 'audio', format: 'Áudio MP3' };
  return null;
}

// Contêineres de áudio e vídeo: o mesmo formato serve a vídeos, músicas e imagens (um .m4a é MP4,
// um .mka é Matroska, um .heic ou um .cr3 é MPEG-4, um .wma é ASF).
const AV_CONTAINERS = new Set(['mp4', 'matroska', 'ogg', 'asf']);
const MEDIA = new Set(['audio', 'video', 'image']);

/**
 * Extensões de outros tipos que usam o mesmo formato por dentro: não são arquivos renomeados (um
 * .docx ou um .jar é ZIP; um .doc ou um .xls com conteúdo ZIP foi só renomeado entre versões do
 * Office; um .ova é TAR; um .dmg costuma ser compactado).
 */
const SAME_CONTAINER = {
  zip: [
    ...['.zip', '.docx', '.docm', '.dotx', '.dotm', '.xlsx', '.xlsm', '.xlsb', '.xltx', '.xltm', '.pptx', '.pptm', '.ppsx', '.ppsm', '.potx', '.potm', '.vsdx', '.xps', '.oxps'],
    ...['.doc', '.dot', '.xls', '.xlt', '.ppt', '.pot', '.pps'],
    ...['.odt', '.ods', '.odp', '.odg', '.ott', '.ots', '.otp', '.epub', '.pages', '.numbers', '.key'],
    ...['.jar', '.war', '.ear', '.apk', '.aab', '.appx', '.appxbundle', '.msix', '.msixbundle', '.xpi', '.vsix', '.nupkg', '.kmz', '.3mf'],
  ],
  tar: ['.ova'],
  compressed: ['.dmg'],
};

/**
 * O conteúdo reconhecido é o formato próprio da extensão (e não um arquivo renomeado)?
 * extensionCategory: a categoria da extensão no catálogo (ou null).
 */
export function sameContainer(detected, extension, extensionCategory = null) {
  const container = detected?.container;
  if (!container) return false;
  if (AV_CONTAINERS.has(container)) return MEDIA.has(extensionCategory);
  // Um aplicativo Java ou Android dentro de um .docx foi renomeado; dentro de um .zip, não.
  if (container === 'zip' && detected.category === 'executable') return extension === '.zip' || extensionCategory === 'executable';
  return Boolean(SAME_CONTAINER[container]?.includes(extension));
}
