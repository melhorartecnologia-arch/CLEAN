// Rotas /api/lists: listas de referência (termos procurados no nome e no conteúdo dos arquivos).
import { Router } from 'express';
import { HttpError, parseList, parseTerms } from './validate.js';

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
    res.json(store.updateList(req.params.id, parseList(req.body)));
  });

  router.delete('/:id', (req, res) => {
    if (!store.deleteList(req.params.id)) throw new HttpError(404, 'Lista não encontrada.');
    res.status(204).end();
  });

  // Valida termos sem salvar (usado pela tela de edição).
  router.post('/validate', (req, res) => {
    res.json({ terms: parseTerms(req.body?.terms || []) });
  });

  return router;
}
