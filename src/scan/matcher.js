// Busca de termos da lista de referência em nomes e conteúdos de arquivos.
//
// - Termos de texto: busca sem diferenciar maiúsculas/minúsculas nem acentos ("salario" encontra
//   "SALÁRIO"), com espaços flexíveis ("João  da\nSilva" encontra "joão da silva"). Usa o
//   algoritmo Aho-Corasick, então o custo não cresce com o tamanho da lista.
// - Termos regex: expressão regular JavaScript (flags "gi") aplicada ao texto original, com
//   validador opcional (CPF, CNPJ, cartão...) para descartar falsos positivos.
import { VALIDATORS, PRESETS } from './presets.js';

// Expressões dos modelos prontos: sabidamente rápidas, dispensam o tempo limite.
const SAFE_PATTERNS = new Set(PRESETS.map((p) => p.value));

const FOLD = new Uint16Array(65536); // caractere -> caractere minúsculo sem acento (1 para 1)
const WORD = new Uint8Array(65536); // 1 se o caractere é letra ou dígito

(function buildTables() {
  const wordRe = /[\p{L}\p{N}]/u;
  const spaceRe = /\s/u;
  for (let c = 0; c < 65536; c++) {
    if (c >= 0xd800 && c <= 0xdfff) {
      FOLD[c] = c;
      WORD[c] = 1;
      continue;
    }
    const ch = String.fromCharCode(c);
    if (spaceRe.test(ch)) {
      FOLD[c] = 32;
      continue;
    }
    // Remove acentos apenas de alfabetos latino, grego e cirílico (a decomposição de outros
    // alfabetos, como o coreano, mudaria o sentido do texto).
    const base = c < 0x0590 || (c >= 0x1e00 && c <= 0x1fff) ? ch.normalize('NFD').charAt(0) : ch;
    const lower = base.toLowerCase();
    FOLD[c] = (lower.length === 1 ? lower : base).charCodeAt(0);
    WORD[c] = wordRe.test(ch) ? 1 : 0;
  }
})();

/** Texto em minúsculas, sem acentos e com espaços simples (para comparações e filtros). */
export function foldText(value) {
  const s = String(value ?? '').normalize('NFC');
  let out = '';
  let prevSpace = false;
  for (let i = 0; i < s.length; i++) {
    const c = FOLD[s.charCodeAt(i)];
    if (c === 32) {
      if (prevSpace) continue;
      prevSpace = true;
    } else {
      prevSpace = false;
    }
    out += String.fromCharCode(c);
  }
  return out.trim();
}

/** Autômato Aho-Corasick sobre os códigos de caractere já normalizados. */
class TextAutomaton {
  constructor() {
    this.next = [new Map()];
    this.fail = [0];
    this.out = [null];
    this.link = [0]; // ligação para o próximo estado (pela cadeia de falhas) que possui saída
    this.maxLen = 0;
  }

  add(key, entry) {
    let s = 0;
    for (let i = 0; i < key.length; i++) {
      const c = key.charCodeAt(i);
      let t = this.next[s].get(c);
      if (t === undefined) {
        t = this.next.length;
        this.next.push(new Map());
        this.fail.push(0);
        this.out.push(null);
        this.link.push(0);
        this.next[s].set(c, t);
      }
      s = t;
    }
    (this.out[s] ||= []).push(entry);
    this.maxLen = Math.max(this.maxLen, key.length);
  }

  build() {
    const queue = [...this.next[0].values()];
    for (let qi = 0; qi < queue.length; qi++) {
      const u = queue[qi];
      for (const [c, v] of this.next[u]) {
        let f = this.fail[u];
        while (f !== 0 && !this.next[f].has(c)) f = this.fail[f];
        const t = this.next[f].get(c);
        this.fail[v] = t !== undefined && t !== v ? t : 0;
        this.link[v] = this.out[this.fail[v]] ? this.fail[v] : this.link[this.fail[v]];
        queue.push(v);
      }
    }
  }

  /** Percorre o texto chamando onMatch(entry, inicio, fim) com posições do texto original. */
  search(text, onMatch) {
    if (this.maxLen === 0) return;
    const { next, fail, out, link } = this;
    const size = this.maxLen + 1;
    const hist = new Int32Array(size); // posição original dos últimos caracteres processados
    let s = 0;
    let k = 0;
    let prevSpace = false;
    for (let i = 0; i < text.length; i++) {
      const c = FOLD[text.charCodeAt(i)];
      if (c === 32) {
        if (prevSpace) continue; // sequências de espaços contam como um espaço só
        prevSpace = true;
      } else {
        prevSpace = false;
      }
      hist[k % size] = i;
      let t = next[s].get(c);
      while (t === undefined && s !== 0) {
        s = fail[s];
        t = next[s].get(c);
      }
      s = t === undefined ? 0 : t;
      for (let o = out[s] ? s : link[s]; o !== 0; o = link[o]) {
        for (const entry of out[o]) {
          if (onMatch(entry, hist[(k - entry.len + 1) % size], i + 1) === false) return;
        }
      }
      k++;
    }
  }
}

/** Valida um termo antes de salvar. Retorna uma mensagem de erro ou null. */
export function validateTerm(term) {
  if (!term || typeof term.value !== 'string' || !term.value.trim()) return 'O termo não pode ser vazio.';
  if (term.type === 'regex') {
    let re;
    try {
      re = new RegExp(term.value, 'gi');
    } catch (err) {
      return `Expressão regular inválida: ${err.message}`;
    }
    if (re.exec('')?.[0] === '') return 'A expressão regular não pode corresponder a um texto vazio.';
    if (term.validator && !VALIDATORS[term.validator]) return `Validador desconhecido: ${term.validator}`;
  } else if (term.type !== 'text') {
    return `Tipo de termo desconhecido: ${term.type}`;
  } else if (!foldText(term.value)) {
    return 'O termo não pode ser vazio.';
  }
  return null;
}

/** Nome exibido do termo nos relatórios. */
export function termDisplay(term) {
  if (term.type === 'regex' && term.label) return term.label;
  return term.value;
}

const DEFAULTS = { maxSamples: 3, maxValues: 10, contextChars: 60, maxCount: 100000 };

/**
 * Compila os termos (de uma ou mais listas) em um buscador reutilizável.
 * Cada termo: { id, value, type: 'text'|'regex', wholeWord?, validator?, label?, listName? }
 */
export class Matcher {
  constructor(terms, options = {}) {
    this.options = { ...DEFAULTS, ...options };
    this.automaton = new TextAutomaton();
    this.regexes = [];
    this.invalid = [];
    this.size = 0;
    for (const term of terms) {
      const error = validateTerm(term);
      if (error) {
        this.invalid.push({ term, error });
        continue;
      }
      const info = {
        id: term.id,
        term: termDisplay(term),
        list: term.listName || '',
        kind: term.type,
      };
      if (term.type === 'regex') {
        const validator = term.validator ? VALIDATORS[term.validator].fn : null;
        this.regexes.push({ info, re: new RegExp(term.value, 'gi'), validator, safe: SAFE_PATTERNS.has(term.value) });
      } else {
        const key = foldText(term.value);
        this.automaton.add(key, {
          info,
          len: key.length,
          wholeWord: Boolean(term.wholeWord),
          firstIsWord: WORD[key.charCodeAt(0)] === 1,
          lastIsWord: WORD[key.charCodeAt(key.length - 1)] === 1,
        });
      }
      this.size++;
    }
    this.automaton.build();
    // guard(fn): executa fn com tempo limite (ver scanner.js). Só é usado se houver expressões
    // regulares escritas pelo usuário, que podem ter retrocesso excessivo e travar a análise.
    this.guard = options.guard && this.regexes.some((r) => !r.safe) ? options.guard : null;
    this.timedOut = false;
  }

  /**
   * Procura os termos em um conjunto de trechos de texto.
   * segments: [{ text, label?, lines? }] — label identifica a parte (ex.: "Página 3") e lines
   * indica que o número da linha deve ser informado nos exemplos.
   * Retorna [{ termId, term, list, kind, location, count, values, samples, truncated }].
   */
  match(segments, location) {
    return this.matchGroups([{ segments, location }])[0];
  }

  /**
   * Procura em vários grupos de trechos (ex.: nome e conteúdo do mesmo arquivo) de uma vez: as
   * expressões regulares de todos os grupos rodam numa única execução protegida por tempo limite.
   * Retorna uma lista de ocorrências por grupo.
   */
  matchGroups(groups) {
    const results = [];
    const prepared = [];
    for (const { segments, location } of groups) {
      const hits = new Map();
      results.push(hits);
      for (const segment of segments || []) {
        if (!segment || !segment.text) continue;
        const text = segment.text.normalize('NFC');
        const record = this.#recorder(text, segment, location, hits);
        prepared.push({ text, record });
        this.#matchTexts(text, record);
      }
    }
    this.timedOut = false;
    if (this.regexes.length && prepared.length) {
      const run = () => {
        for (const { text, record } of prepared) this.#matchRegexes(text, record);
      };
      if (this.guard) {
        try {
          this.guard(run);
        } catch (err) {
          if (err?.code !== 'ERR_SCRIPT_EXECUTION_TIMEOUT') throw err;
          this.timedOut = true; // mantém o que já foi encontrado
        }
      } else {
        run();
      }
    }
    return results.map((hits) => [...hits.values()].map((hit) => ({ ...hit, values: [...hit.values] })));
  }

  /** Função que registra uma ocorrência (contagem, valores distintos e exemplos com contexto). */
  #recorder(text, segment, location, hits) {
    const { maxCount } = this.options;
    return (info, start, end) => {
      let hit = hits.get(info.id);
      if (!hit) {
        hit = {
          termId: info.id,
          term: info.term,
          list: info.list,
          kind: info.kind,
          location,
          count: 0,
          values: new Set(),
          samples: [],
          truncated: false,
        };
        hits.set(info.id, hit);
      }
      if (hit.count >= maxCount) {
        hit.truncated = true;
        return false;
      }
      hit.count++;
      if (hit.values.size < this.options.maxValues) hit.values.add(clean(text.slice(start, end)).slice(0, 200));
      if (hit.samples.length < this.options.maxSamples) hit.samples.push(this.#sample(text, start, end, segment));
      return true;
    };
  }

  #matchTexts(text, record) {
    this.automaton.search(text, (entry, start, end) => {
      if (entry.wholeWord) {
        if (entry.firstIsWord && start > 0 && WORD[text.charCodeAt(start - 1)]) return true;
        if (entry.lastIsWord && end < text.length && WORD[text.charCodeAt(end)]) return true;
      }
      record(entry.info, start, end);
      return true;
    });
  }

  #matchRegexes(text, record) {
    for (const { info, re, validator } of this.regexes) {
      re.lastIndex = 0;
      let m;
      while ((m = re.exec(text)) !== null) {
        if (m[0].length === 0) {
          re.lastIndex = m.index + 1;
          continue;
        }
        if (validator && !validator(m[0])) continue;
        if (record(info, m.index, m.index + m[0].length) === false) break;
      }
    }
  }

  #sample(text, start, end, segment) {
    const c = this.options.contextChars;
    const from = Math.max(0, start - c);
    const to = Math.min(text.length, end + c);
    const where = [];
    if (segment.label) where.push(segment.label);
    if (segment.lines) where.push(`linha ${lineNumber(text, start)}`);
    return {
      before: (from > 0 ? '…' : '') + clean(text.slice(from, start)).trimStart(),
      match: clean(text.slice(start, end)).slice(0, 300),
      after: clean(text.slice(end, to)).trimEnd() + (to < text.length ? '…' : ''),
      where: where.join(', '),
    };
  }
}

function lineNumber(text, index) {
  let line = 1;
  for (let i = text.indexOf('\n'); i !== -1 && i < index; i = text.indexOf('\n', i + 1)) line++;
  return line;
}

/** Remove caracteres de controle e agrupa espaços para exibição. */
function clean(value) {
  // eslint-disable-next-line no-control-regex
  return value.replace(/[\u0000-\u001f\u007f-\u009f�]+/g, ' ').replace(/\s+/g, ' ');
}
