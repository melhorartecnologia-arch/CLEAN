// PDF: texto por página (pdf.js via unpdf) e metadados (autor, aplicativo, datas).
import { getDocumentProxy } from 'unpdf';
import { compact } from './ooxml.js';

export function isPdf(buf) {
  return buf.length > 5 && buf.toString('latin1', 0, 1024).includes('%PDF-');
}

/** Converte datas PDF ("D:20240301103000-03'00'") para ISO. */
export function pdfDate(value) {
  const m = /^D?:?(\d{4})(\d{2})?(\d{2})?(\d{2})?(\d{2})?(\d{2})?([Zz]|[+-]\d{2}'?\d{2}'?)?/.exec(String(value || ''));
  if (!m) return null;
  const [, y, mo = '01', d = '01', h = '00', mi = '00', s = '00', tz] = m;
  let offset = 'Z';
  if (tz && tz.toUpperCase() !== 'Z') {
    const digits = tz.replace(/'/g, '');
    offset = `${digits.slice(0, 3)}:${digits.slice(3, 5) || '00'}`;
  }
  const date = new Date(`${y}-${mo}-${d}T${h}:${mi}:${s}${offset}`);
  return Number.isNaN(date.getTime()) ? null : date.toISOString();
}

/**
 * Extrai o texto página a página, parando ao atingir maxChars.
 * Retorna { segments, metadata, truncated, encrypted }.
 */
export async function pdfExtract(buf, { maxChars = 20_000_000, withText = true } = {}) {
  let pdf;
  try {
    // Cópia do buffer: o pdf.js pode transferir (e invalidar) o ArrayBuffer recebido.
    pdf = await getDocumentProxy(new Uint8Array(buf), {
      verbosity: 0,
      isEvalSupported: false,
      disableFontFace: true,
      useSystemFonts: false,
      stopAtErrors: false,
    });
  } catch (err) {
    if (err?.name === 'PasswordException') return { segments: [], metadata: {}, encrypted: true };
    throw err;
  }
  try {
    const segments = [];
    let total = 0;
    let truncated = false;
    if (withText) {
      for (let p = 1; p <= pdf.numPages; p++) {
        if (total >= maxChars) {
          truncated = true;
          break;
        }
        const page = await pdf.getPage(p);
        const content = await page.getTextContent();
        let text = '';
        for (const item of content.items) {
          if (typeof item.str !== 'string') continue;
          text += item.str;
          if (item.hasEOL) text += '\n';
        }
        page.cleanup();
        total += text.length;
        if (text.trim()) segments.push({ label: `Página ${p}`, text });
      }
    }
    let metadata = {};
    try {
      const { info } = await pdf.getMetadata();
      metadata = compact({
        author: clean(info?.Author),
        title: clean(info?.Title),
        application: clean(info?.Creator) || clean(info?.Producer),
        created: pdfDate(info?.CreationDate),
        modified: pdfDate(info?.ModDate),
      });
    } catch {
      // metadados são opcionais
    }
    return { segments, metadata, truncated, pages: pdf.numPages };
  } finally {
    await pdf.loadingTask?.destroy?.().catch(() => {});
  }
}

function clean(value) {
  return typeof value === 'string' ? value.replace(/\0/g, '').trim() || null : null;
}
