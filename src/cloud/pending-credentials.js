// Credenciais da Microsoft que ainda não foram salvas numa conexão: entradas de contas pelo código de
// dispositivo (em andamento ou concluídas) e certificados gerados ou importados. Ficam só na memória
// do servidor, por tempo limitado, até a conexão ser salva: os tokens e as chaves privadas nunca vão
// para o navegador, que recebe apenas um identificador aleatório.
import crypto from 'node:crypto';

const MAX_ITEMS = 200;

export class PendingCredentials {
  constructor({ now = () => Date.now() } = {}) {
    this.items = new Map();
    this.now = now;
  }

  #sweep() {
    const now = this.now();
    for (const [id, item] of this.items) if (item.expiresAt <= now) this.items.delete(id);
    // Limite de itens na memória: os mais antigos saem primeiro.
    while (this.items.size > MAX_ITEMS) this.items.delete(this.items.keys().next().value);
  }

  /** Guarda um item (kind: 'device' ou 'certificate') por ttlMs; devolve o item, com o identificador. */
  add(kind, data, ttlMs) {
    this.#sweep();
    const item = { ...data, id: crypto.randomUUID(), kind, expiresAt: this.now() + ttlMs };
    this.items.set(item.id, item);
    return item;
  }

  /** O item ainda válido com este identificador e deste tipo, ou null. */
  get(id, kind) {
    this.#sweep();
    const item = typeof id === 'string' ? this.items.get(id) : null;
    return item && item.kind === kind ? item : null;
  }

  delete(id) {
    if (typeof id === 'string') this.items.delete(id);
  }
}
