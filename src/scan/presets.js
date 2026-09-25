// Validadores e modelos prontos de termos (expressões regulares) para dados sensíveis comuns.

/** Mantém apenas os dígitos de um texto. */
function onlyDigits(value) {
  return String(value).replace(/\D/g, '');
}

/** CPF: 11 dígitos, rejeita sequências repetidas e confere os dois dígitos verificadores. */
export function isValidCpf(value) {
  const d = onlyDigits(value);
  if (d.length !== 11 || /^(\d)\1{10}$/.test(d)) return false;
  for (const len of [9, 10]) {
    let sum = 0;
    for (let i = 0; i < len; i++) sum += Number(d[i]) * (len + 1 - i);
    const dv = ((sum * 10) % 11) % 10;
    if (dv !== Number(d[len])) return false;
  }
  return true;
}

/**
 * CNPJ numérico ou alfanumérico (IN RFB 2.229/2024, vigente desde julho/2026).
 * Cada caractere vale (código ASCII - 48): '0'..'9' => 0..9, 'A' => 17 ... 'Z' => 42.
 */
export function isValidCnpj(value) {
  const raw = String(value);
  if (/[a-z]/.test(raw)) return false; // o CNPJ alfanumérico usa apenas letras maiúsculas
  const s = raw.replace(/[.\-/\s]/g, '');
  if (!/^[0-9A-Z]{12}\d{2}$/.test(s) || /^(.)\1{13}$/.test(s)) return false;
  const values = [...s].map((ch) => ch.charCodeAt(0) - 48);
  const calc = (len) => {
    const weights = len === 12 ? [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2] : [6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2];
    let sum = 0;
    for (let i = 0; i < len; i++) sum += values[i] * weights[i];
    const rest = sum % 11;
    return rest < 2 ? 0 : 11 - rest;
  };
  return calc(12) === values[12] && calc(13) === values[13];
}

/** Cartão de pagamento: 13 a 19 dígitos, prefixo 3-6 e algoritmo de Luhn. */
export function isValidCard(value) {
  const d = onlyDigits(value);
  if (d.length < 13 || d.length > 19 || !/^[3-6]/.test(d) || /^(\d)\1+$/.test(d)) return false;
  let sum = 0;
  for (let i = 0; i < d.length; i++) {
    let n = Number(d[d.length - 1 - i]);
    if (i % 2 === 1) {
      n *= 2;
      if (n > 9) n -= 9;
    }
    sum += n;
  }
  return sum % 10 === 0;
}

/** PIS/PASEP/NIT: 11 dígitos com dígito verificador (pesos 3,2,9,8,7,6,5,4,3,2). */
export function isValidPis(value) {
  const d = onlyDigits(value);
  if (d.length !== 11 || /^(\d)\1{10}$/.test(d)) return false;
  const weights = [3, 2, 9, 8, 7, 6, 5, 4, 3, 2];
  let sum = 0;
  for (let i = 0; i < 10; i++) sum += Number(d[i]) * weights[i];
  const rest = 11 - (sum % 11);
  const dv = rest >= 10 ? 0 : rest;
  return dv === Number(d[10]);
}

export const VALIDATORS = {
  cpf: { label: 'CPF (dígitos verificadores)', fn: isValidCpf },
  cnpj: { label: 'CNPJ numérico ou alfanumérico (dígitos verificadores)', fn: isValidCnpj },
  cartao: { label: 'Cartão de pagamento (algoritmo de Luhn)', fn: isValidCard },
  pis: { label: 'PIS/PASEP/NIT (dígito verificador)', fn: isValidPis },
};

/** Modelos oferecidos na tela de listas de referência. */
export const PRESETS = [
  {
    id: 'cpf',
    label: 'CPF',
    type: 'regex',
    value: String.raw`\b\d{3}\.?\d{3}\.?\d{3}-?\d{2}\b`,
    validator: 'cpf',
    description: 'Números de CPF válidos, com ou sem pontuação.',
  },
  {
    id: 'cnpj',
    label: 'CNPJ',
    type: 'regex',
    value: String.raw`\b(?:\d{2}\.?\d{3}\.?\d{3}\/?\d{4}-?\d{2}|[0-9A-Z]{2}\.[0-9A-Z]{3}\.[0-9A-Z]{3}\/[0-9A-Z]{4}-\d{2})\b`,
    validator: 'cnpj',
    description: 'CNPJ válido (numérico ou alfanumérico formatado).',
  },
  {
    id: 'pis',
    label: 'PIS/PASEP/NIT',
    type: 'regex',
    value: String.raw`\b\d{3}\.?\d{5}\.?\d{2}-?\d\b`,
    validator: 'pis',
    description: 'Números de PIS/PASEP/NIT válidos.',
  },
  {
    id: 'email',
    label: 'E-mail',
    type: 'regex',
    value: String.raw`\b[A-Za-z0-9._%+-]+@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)*\.[A-Za-z]{2,}\b`,
    validator: null,
    description: 'Endereços de e-mail.',
  },
  {
    id: 'telefone',
    label: 'Telefone (BR)',
    type: 'regex',
    value: String.raw`(?:\+55\s?)?(?:\(\d{2}\)\s?|\b\d{2}[\s-])9?\d{4}[\s-]\d{4}\b`,
    validator: null,
    description: 'Telefones com DDD, ex.: (24) 99999-9999.',
  },
  {
    id: 'cartao',
    label: 'Cartão de crédito',
    type: 'regex',
    value: String.raw`\b(?:\d[ .-]?){12,18}\d\b`,
    validator: 'cartao',
    description: 'Números de cartão válidos pelo algoritmo de Luhn.',
  },
  {
    id: 'senha',
    label: 'Senha em texto',
    type: 'regex',
    value: String.raw`\b(?:senha|password|passwd|pwd)\s*[:=]\s*\S+`,
    validator: null,
    description: 'Trechos como "senha: 123" ou "password=abc".',
  },
];
