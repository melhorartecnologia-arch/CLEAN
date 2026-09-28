// Certificado para a autenticação de um aplicativo no Microsoft Entra ID (no lugar do segredo do
// cliente): o CLEAN gera um par de chaves RSA e um certificado autoassinado (X.509) — o arquivo do
// certificado (só a parte pública) é enviado ao registro do aplicativo, e a chave privada fica
// gravada cifrada no CLEAN. Também aceita um certificado existente em PEM (certificado e chave).
import crypto from 'node:crypto';
import { promisify } from 'node:util';

const generateKeyPair = promisify(crypto.generateKeyPair);

// Codificação DER (ASN.1) mínima para montar o certificado.
const der = {
  length(n) {
    if (n < 0x80) return Buffer.from([n]);
    const bytes = [];
    for (let v = n; v > 0; v = Math.floor(v / 256)) bytes.unshift(v % 256);
    return Buffer.from([0x80 | bytes.length, ...bytes]);
  },
  tlv: (tag, content) => Buffer.concat([Buffer.from([tag]), der.length(content.length), content]),
  seq: (...items) => der.tlv(0x30, Buffer.concat(items)),
  set: (...items) => der.tlv(0x31, Buffer.concat(items)),
  /** Inteiro positivo a partir dos bytes (sem zeros à esquerda; com um zero se o 1º bit estiver ligado). */
  int(bytes) {
    let b = Buffer.from(bytes);
    while (b.length > 1 && b[0] === 0 && !(b[1] & 0x80)) b = b.subarray(1);
    if (b[0] & 0x80) b = Buffer.concat([Buffer.from([0]), b]);
    return der.tlv(0x02, b);
  },
  oid(text) {
    const parts = text.split('.').map(Number);
    const bytes = [40 * parts[0] + parts[1]];
    for (const part of parts.slice(2)) {
      const chunk = [part & 0x7f];
      for (let v = part >>> 7; v > 0; v >>>= 7) chunk.unshift((v & 0x7f) | 0x80);
      bytes.push(...chunk);
    }
    return der.tlv(0x06, Buffer.from(bytes));
  },
  null: () => Buffer.from([0x05, 0x00]),
  bool: (value) => der.tlv(0x01, Buffer.from([value ? 0xff : 0x00])),
  utf8: (text) => der.tlv(0x0c, Buffer.from(text, 'utf8')),
  octets: (bytes) => der.tlv(0x04, bytes),
  bits: (bytes, unused = 0) => der.tlv(0x03, Buffer.concat([Buffer.from([unused]), bytes])),
  explicit: (n, content) => der.tlv(0xa0 + n, content),
  /** UTCTime até 2049 e GeneralizedTime a partir de 2050 (RFC 5280). */
  time(date) {
    const text = date.toISOString().replace(/[-:T]/g, '').slice(0, 14);
    return date.getUTCFullYear() < 2050 ? der.tlv(0x17, Buffer.from(`${text.slice(2)}Z`)) : der.tlv(0x18, Buffer.from(`${text}Z`));
  },
};

const SHA256_WITH_RSA = () => der.seq(der.oid('1.2.840.113549.1.1.11'), der.null());
const pem = (label, bytes) => `-----BEGIN ${label}-----\n${bytes.toString('base64').replace(/.{1,64}/g, '$&\n')}-----END ${label}-----\n`;

/** Dados públicos do certificado: impressões digitais (a SHA-1 é a mostrada no Entra ID), assunto e validade. */
export function certificateInfo(certificatePem) {
  const cert = new crypto.X509Certificate(certificatePem);
  return {
    thumbprint: cert.fingerprint.replace(/:/g, '').toUpperCase(),
    thumbprint256: Buffer.from(cert.fingerprint256.replace(/:/g, ''), 'hex').toString('base64url'),
    subject: cert.subject.replace(/\n/g, ', '),
    notBefore: new Date(cert.validFrom).toISOString(),
    notAfter: new Date(cert.validTo).toISOString(),
    pem: cert.toString(),
  };
}

/**
 * Gera uma chave RSA de 2048 bits e um certificado autoassinado para ela (válido por `years` anos a
 * partir de uma hora atrás, para tolerar relógios um pouco adiantados). Retorna
 * { privateKeyPem, certificate: certificateInfo }.
 */
export async function createCertificate({ commonName = 'CLEAN', years = 2, now = new Date() } = {}) {
  const { publicKey, privateKey } = await generateKeyPair('rsa', { modulusLength: 2048 });
  const spki = publicKey.export({ type: 'spki', format: 'der' });
  const notBefore = new Date(now.getTime() - 3600 * 1000);
  const notAfter = new Date(now);
  notAfter.setUTCFullYear(notAfter.getUTCFullYear() + years);
  const serial = crypto.randomBytes(16);
  serial[0] = (serial[0] & 0x7f) || 0x01; // positivo e sem zero à esquerda
  const name = der.seq(der.set(der.seq(der.oid('2.5.4.3'), der.utf8(String(commonName).slice(0, 64) || 'CLEAN'))));
  // Identificador da chave: SHA-1 da chave pública (os bits do SubjectPublicKeyInfo).
  const keyId = crypto.createHash('sha1').update(crypto.createPublicKey({ key: spki, format: 'der', type: 'spki' }).export({ type: 'pkcs1', format: 'der' })).digest();
  const extensions = der.explicit(
    3,
    der.seq(
      der.seq(der.oid('2.5.29.19'), der.bool(true), der.octets(der.seq())), // basicConstraints: não é autoridade certificadora
      der.seq(der.oid('2.5.29.15'), der.bool(true), der.octets(der.bits(Buffer.from([0x80]), 7))), // keyUsage: assinatura digital
      der.seq(der.oid('2.5.29.14'), der.octets(der.octets(keyId))), // subjectKeyIdentifier
    ),
  );
  const tbs = der.seq(
    der.explicit(0, der.int([2])), // versão 3
    der.int(serial),
    SHA256_WITH_RSA(),
    name,
    der.seq(der.time(notBefore), der.time(notAfter)),
    name,
    spki,
    extensions,
  );
  const signature = crypto.sign('sha256', tbs, privateKey);
  const certificate = der.seq(tbs, SHA256_WITH_RSA(), der.bits(signature));
  return {
    privateKeyPem: privateKey.export({ type: 'pkcs8', format: 'pem' }),
    certificate: certificateInfo(pem('CERTIFICATE', certificate)),
  };
}

/**
 * Lê um certificado existente em PEM (o certificado e a chave privada sem senha, no mesmo texto).
 * Confere que a chave é a do certificado. Retorna { privateKeyPem, certificate: certificateInfo }.
 */
export function importCertificate(text) {
  const value = String(text || '');
  const certBlock = /-----BEGIN CERTIFICATE-----[\s\S]+?-----END CERTIFICATE-----/.exec(value)?.[0];
  if (!certBlock) throw new Error('O arquivo não tem um certificado em PEM ("-----BEGIN CERTIFICATE-----"). Para um arquivo .pfx, converta antes: openssl pkcs12 -in certificado.pfx -out certificado.pem -nodes');
  if (/-----BEGIN ENCRYPTED PRIVATE KEY-----/.test(value)) {
    throw new Error('A chave privada está protegida por senha: remova a senha antes (openssl pkey -in chave.pem -out chave-aberta.pem) e envie o certificado e a chave aberta no mesmo arquivo.');
  }
  const keyBlock = /-----BEGIN (RSA |EC )?PRIVATE KEY-----[\s\S]+?-----END \1?PRIVATE KEY-----/.exec(value)?.[0];
  if (!keyBlock) throw new Error('O arquivo não tem a chave privada ("-----BEGIN PRIVATE KEY-----"): envie o certificado e a chave privada no mesmo arquivo PEM.');
  let cert;
  let key;
  try {
    cert = new crypto.X509Certificate(certBlock);
  } catch {
    throw new Error('O certificado do arquivo é inválido.');
  }
  try {
    key = crypto.createPrivateKey(keyBlock);
  } catch {
    throw new Error('A chave privada do arquivo é inválida.');
  }
  if (key.asymmetricKeyType !== 'rsa') throw new Error('Use um certificado com chave RSA (o Microsoft Entra ID não aceita outros tipos para a autenticação de aplicativos).');
  if ((key.asymmetricKeyDetails?.modulusLength || 0) < 2048) throw new Error('A chave RSA do certificado precisa ter ao menos 2048 bits.');
  if (!cert.checkPrivateKey(key)) throw new Error('A chave privada não é a do certificado informado.');
  if (Date.parse(cert.validTo) < Date.now()) throw new Error(`O certificado venceu em ${new Date(cert.validTo).toLocaleDateString('pt-BR')}.`);
  return { privateKeyPem: key.export({ type: 'pkcs8', format: 'pem' }), certificate: certificateInfo(certBlock) };
}
