// Mensagens amigáveis para erros de sistema de arquivos.

const ERROR_MESSAGES = {
  EACCES: 'Acesso negado',
  EPERM: 'Acesso negado (permissão)',
  ENOENT: 'Não encontrado (pode ter sido removido durante a análise)',
  EBUSY: 'Arquivo em uso ou bloqueado',
  ENAMETOOLONG: 'Caminho muito longo',
  ELOOP: 'Laço de atalhos',
  EIO: 'Erro de leitura (E/S)',
  ETIMEDOUT: 'Tempo esgotado ao acessar a rede',
  EHOSTUNREACH: 'Servidor inacessível',
  ENOTDIR: 'O caminho não é uma pasta',
};

export function friendlyError(err) {
  const base = ERROR_MESSAGES[err?.code];
  return base ? `${base} (${err.code})` : String(err?.message || err);
}

export class TimeoutError extends Error {
  constructor(message) {
    super(message);
    this.code = 'ETIMEOUT';
  }
}

/** Rejeita com TimeoutError se a promessa não terminar em `ms` milissegundos. */
export function withTimeout(promise, ms, message = 'Tempo esgotado.') {
  let timer;
  const timeout = new Promise((_, reject) => {
    timer = setTimeout(() => reject(new TimeoutError(message)), ms);
  });
  return Promise.race([promise, timeout]).finally(() => clearTimeout(timer));
}
