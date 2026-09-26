// Rotas /api/lists: listas de referência (termos procurados no nome e no conteúdo dos arquivos).
import { Router } from 'express';
import { HttpError, parseList, parseTerms, assertUnused } from './validate.js';
import { checkDeleteSchedules } from '../schedule/scheduler.js';
import { Worker } from 'node:worker_threads';

/** Roda o teste em uma worker thread com tempo limite. */
function testTerms(terms, text, timeoutMs = 3000) {
  return new Promise((resolve, reject) => {
    const worker = new Worker(new URL('../scan/sample-worker.js', import.meta.url), { workerData: { terms, text } });
    const timer = setTimeout(() => {
      worker.terminate();
      reject(new HttpError(422, 'O teste demorou demais. Revise as expressões regulares (podem ter retrocesso excessivo).'));
    }, timeoutMs);
    worker.once('message', (matches) => {
      clearTimeout(timer);
      worker.terminate();
      resolve(matches);
    });
    worker.once('error', (err) => {
      clearTimeout(timer);
      reject(err);
    });
  });
}

const summary = (list) => ({
  id: list.id,
  name: list.name,
  description: list.description,
  termCount: (list.terms || []).length,
  createdAt: list.createdAt,
  updatedAt: list.updatedAt,
});

export function listsRouter({ store }) {
  const router = Router();

  router.get('/', (req, res) => {
    res.json(store.listLists().map(summary));
  });

  router.get('/:id', (req, res) => {
    const list = store.getList(req.params.id);
    if (!list) throw new HttpError(404, 'Lista não encontrada.');
    res.json(list);
  });

  router.post('/', (req, res) => {
    res.status(201).json(store.createList(parseList(req.body)));
  });

  router.put('/:id', (req, res) => {
    if (!store.getList(req.params.id)) throw new HttpError(404, 'Lista não encontrada.');
    const data = parseList(req.body);
    // Termos novos ou alterados suspendem a exclusão automática dos agendamentos que usam a lista.
    const { result, warning } = checkDeleteSchedules(store, () => store.updateList(req.params.id, data));
    res.json({ ...result, ...(warning ? { scheduleWarning: warning } : {}) });
  });

  router.delete('/:id', (req, res) => {
    if (!store.getList(req.params.id)) throw new HttpError(404, 'Lista não encontrada.');
    assertUnused(store, 'list', req.params.id, 'A lista');
    store.deleteList(req.params.id);
    res.status(204).end();
  });

  // Testa os termos (ainda não salvos) contra um texto de exemplo.
  router.post('/test', async (req, res) => {
    const terms = parseTerms(req.body?.terms || []);
    const text = String(req.body?.text || '').slice(0, 200000);
    res.json({ matches: await testTerms(terms, text) });
  });

  return router;
}
