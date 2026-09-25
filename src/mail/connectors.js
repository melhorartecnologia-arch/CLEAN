// Escolhe o conector de cada tipo de conexão de e-mail.
import { GraphConnector } from './graph.js';
import { GmailConnector } from './gmail.js';
import { ImapConnector } from './imap.js';

export const MAIL_TYPES = {
  graph: 'Microsoft 365 (Exchange Online)',
  gmail: 'Google Workspace (Gmail)',
  imap: 'IMAP',
};

/**
 * source: conexão com os segredos já decifrados (source.secrets).
 * options: { signal, log(level, message), endpoints } — endpoints troca os endereços das APIs (testes).
 */
export function createConnector(source, options = {}) {
  switch (source.type) {
    case 'graph':
      return new GraphConnector(source, options);
    case 'gmail':
      return new GmailConnector(source, options);
    case 'imap':
      return new ImapConnector(source, options);
    default:
      throw new Error(`Tipo de conexão de e-mail desconhecido: ${source.type}`);
  }
}
