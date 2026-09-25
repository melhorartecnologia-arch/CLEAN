// Rotas /api/repositories: pastas (locais ou compartilhamentos de rede) a analisar.
import fs from 'node:fs/promises';
import { Router } from 'express';
import { HttpError, parseRepository, normalizeRepoPath } from './validate.js';
import { friendlyError } from '../scan/errors.js';

export function repositoriesRouter({ store, manager = null }) {
  const router = Router();

  router.get('/', (req, res) => {
    res.json(store.listRepositories());
  });

  router.post('/', (req, res) => {
    const data = parseRepository(req.body);
    res.status(201).json(store.createRepository(data));
  });

  router.put('/:id', (req, res) => {
    const existing = store.getRepository(req.params.id);
    if (!existing) throw new HttpError(404, 'Repositório não encontrado.');
    const before = { allowDelete: existing.allowDelete, path: existing.path };
    const updated = store.updateRepository(existing.id, parseRepository(req.body));
    // Análises em andamento deixam de excluir se a exclusão foi desligada ou o caminho mudou.
    if (before.allowDelete && (!updated.allowDelete || updated.path !== before.path)) {
      manager?.revokeDeletion('repository', existing.id, updated.allowDelete ? 'o caminho do repositório foi alterado' : 'a opção "Permitir exclusão" foi desligada');
    }
    res.json(updated);
  });

  router.delete('/:id', (req, res) => {
    const existing = store.getRepository(req.params.id);
    if (!existing) throw new HttpError(404, 'Repositório não encontrado.');
    store.deleteRepository(existing.id);
    if (existing.allowDelete) manager?.revokeDeletion('repository', existing.id, 'o repositório foi removido do cadastro');
    res.status(204).end();
  });

  // Verifica se a pasta existe e pode ser lida pela conta que executa o CLEAN.
  router.post('/test', async (req, res) => {
    const target = normalizeRepoPath(req.body?.path);
    const started = Date.now();
    try {
      const st = await fs.stat(target);
      if (!st.isDirectory()) return res.json({ ok: false, message: 'O caminho existe, mas não é uma pasta.' });
      const dir = await fs.opendir(target);
      let files = 0;
      let folders = 0;
      const sample = [];
      for await (const entry of dir) {
        if (entry.isDirectory()) folders++;
        else files++;
        if (sample.length < 8) sample.push(entry.name);
        if (files + folders >= 500) break;
      }
      const more = files + folders >= 500 ? ' (ou mais)' : '';
      res.json({
        ok: true,
        message: `Pasta acessível: ${folders} pasta(s) e ${files} arquivo(s) no primeiro nível${more}.`,
        sample,
        elapsedMs: Date.now() - started,
      });
    } catch (err) {
      res.json({ ok: false, message: friendlyError(err) });
    }
  });

  return router;
}
