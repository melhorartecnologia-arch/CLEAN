// Segredos das conexões de e-mail (senhas, segredos de aplicativo, chaves privadas) gravados cifrados
// com AES-256-GCM. A chave fica no arquivo data/chave-segredos.key (criado na primeira execução) ou
// na variável de ambiente CLEAN_SECRET_KEY (32 bytes em base64). Sem a chave, os segredos não podem
// ser lidos: guarde uma cópia dela junto com o backup da pasta de dados.
import crypto from 'node:crypto';
import fs from 'node:fs/promises';
import path from 'node:path';

const PREFIX = 'enc:v1:';
export const KEY_FILE = 'chave-segredos.key';

function parseKey(text, origin) {
  const key = Buffer.from(String(text).trim(), 'base64');
  if (key.length !== 32) throw new Error(`Chave de segredos inválida em ${origin} (esperados 32 bytes em base64).`);
  return key;
}

export class SecretBox {
  constructor(key) {
    this.key = key;
  }

  /** Abre (ou cria) a chave da pasta de dados. */
  static async open(dataDir, envKey = process.env.CLEAN_SECRET_KEY) {
    if (envKey) return new SecretBox(parseKey(envKey, 'CLEAN_SECRET_KEY'));
    const file = path.join(dataDir, KEY_FILE);
    try {
      return new SecretBox(parseKey(await fs.readFile(file, 'utf8'), file));
    } catch (err) {
      if (err.code !== 'ENOENT') throw err;
    }
    try {
      await fs.writeFile(file, `${crypto.randomBytes(32).toString('base64')}\n`, { encoding: 'utf8', mode: 0o600, flag: 'wx' });
    } catch (err) {
      if (err.code !== 'EEXIST') throw err; // outro processo criou a chave ao mesmo tempo
    }
    return new SecretBox(parseKey(await fs.readFile(file, 'utf8'), file));
  }

  seal(text) {
    const iv = crypto.randomBytes(12);
    const cipher = crypto.createCipheriv('aes-256-gcm', this.key, iv);
    const data = Buffer.concat([cipher.update(String(text), 'utf8'), cipher.final()]);
    return `${PREFIX}${iv.toString('base64')}:${cipher.getAuthTag().toString('base64')}:${data.toString('base64')}`;
  }

  open(sealed) {
    if (!sealed) return '';
    if (typeof sealed !== 'string' || !sealed.startsWith(PREFIX)) throw new Error('Segredo gravado em formato desconhecido.');
    const [iv, tag, data] = sealed.slice(PREFIX.length).split(':').map((part) => Buffer.from(part || '', 'base64'));
    try {
      const decipher = crypto.createDecipheriv('aes-256-gcm', this.key, iv);
      decipher.setAuthTag(tag);
      return Buffer.concat([decipher.update(data), decipher.final()]).toString('utf8');
    } catch {
      throw new Error(`Não foi possível decifrar um segredo salvo: a chave (${KEY_FILE} ou CLEAN_SECRET_KEY) mudou. Informe as senhas e chaves da conexão novamente.`);
    }
  }
}
