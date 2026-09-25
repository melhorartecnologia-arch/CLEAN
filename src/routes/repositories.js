// Rotas /api/repositories: pastas (locais ou compartilhamentos de rede) a analisar.
import fs from 'node:fs/promises';
import { Router } from 'express';
import { HttpError, parseRepository, normalizeRepoPath } from './validate.js';
import { friendlyError } from '../scan/scanner.js';

export function repositoriesRouter({ store }) {
  const router = Router();

  router.get('/', (req, res) => {
    res.json(store.listRepositories());
  });

  router.post('/', (req, res) => {
    const data = parseRepository(req.body);
    res.status(201).json(store.createRepository(data));
  });

  router.put('/:id', (req, res) => {
    if (!store.getRepository(req.params.id)) throw new HttpError(404, 'Repositório não encontrado.');
    res.json(store.updateRepository(req.params.id, parseRepository(req.body)));
  });

  router.delete('/:id', (req, res) => {
    if (!store.deleteRepository(req.params.id)) throw new HttpError(404, 'Repositório não encontrado.');
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
