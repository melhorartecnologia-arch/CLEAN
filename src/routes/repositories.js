// Rotas /api/repositories: pastas (locais ou compartilhamentos de rede) e bibliotecas do OneDrive e
// do SharePoint a analisar. O segredo do cliente dos repositórios na nuvem é gravado cifrado e nunca
// volta para o navegador: a API informa apenas se ele está salvo.
import fs from 'node:fs/promises';
import { Router } from 'express';
import { HttpError, parseRepository, normalizeRepoPath } from './validate.js';
import { friendlyError } from '../scan/errors.js';
import { DrivesConnector } from '../cloud/drives.js';
import { isCloudRepo, deletionScope } from '../scan/delete.js';

/** Repositório sem os segredos (apenas indica se o segredo do cliente está salvo). */
export function publicRepository(repo) {
  const { secrets, ...rest } = repo;
  const out = { ...rest, type: repo.type || 'local' };
  if (repo.graph) out.graph = { ...repo.graph, hasClientSecret: Boolean(secrets?.clientSecret) };
  return out;
}

export function repositoriesRouter({ store, manager = null, endpoints = {} }) {
  const router = Router();
  const mailSourceOf = (body) => (typeof body?.credentialsFrom === 'string' && body.credentialsFrom ? store.getMailSource(body.credentialsFrom) : null);

  router.get('/', (req, res) => {
    res.json(store.listRepositories().map(publicRepository));
  });

  router.post('/', (req, res) => {
    const data = parseRepository(req.body, { box: store.secrets, mailSource: mailSourceOf(req.body) });
    res.status(201).json(publicRepository(store.createRepository(data)));
  });

  router.put('/:id', (req, res) => {
    const existing = store.getRepository(req.params.id);
    if (!existing) throw new HttpError(404, 'Repositório não encontrado.');
    const before = { allowDelete: existing.allowDelete, scope: deletionScope(existing), deleteMode: existing.deleteMode };
    const updated = store.updateRepository(existing.id, parseRepository(req.body, { existing, box: store.secrets, mailSource: mailSourceOf(req.body) }));
    // Análises em andamento deixam de excluir se a exclusão foi desligada ou o alcance/forma mudou.
    if (before.allowDelete && (!updated.allowDelete || deletionScope(updated) !== before.scope || (updated.deleteMode || null) !== (before.deleteMode || null))) {
      const reason = !updated.allowDelete
        ? 'a opção "Permitir exclusão" foi desligada'
        : deletionScope(updated) !== before.scope
          ? isCloudRepo(updated)
            ? 'as contas, os sites ou as credenciais do repositório foram alterados'
            : 'o caminho do repositório foi alterado'
          : 'a forma de exclusão foi alterada';
      manager?.revokeDeletion('repository', existing.id, reason);
    }
    res.json(publicRepository(updated));
  });

  router.delete('/:id', (req, res) => {
    const existing = store.getRepository(req.params.id);
    if (!existing) throw new HttpError(404, 'Repositório não encontrado.');
    store.deleteRepository(existing.id);
    if (existing.allowDelete) manager?.revokeDeletion('repository', existing.id, 'o repositório foi removido do cadastro');
    res.status(204).end();
  });

  // Verifica o acesso: a pasta (existe e pode ser lida pela conta que executa o CLEAN) ou a conexão
  // com o OneDrive/SharePoint (com os dados do formulário; segredo vazio usa o salvo).
  router.post('/test', async (req, res) => {
    if (req.body?.type === 'onedrive' || req.body?.type === 'sharepoint') return testCloud(req, res);
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

  async function testCloud(req, res) {
    const id = typeof req.body?.id === 'string' ? req.body.id : '';
    const existing = id ? store.getRepository(id) : null;
    if (id && !existing) throw new HttpError(404, 'Repositório não encontrado.');
    const data = parseRepository(req.body, { existing, box: store.secrets, mailSource: mailSourceOf(req.body), forTest: true });
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(new Error('timeout')), 60000);
    try {
      const connector = new DrivesConnector({ ...data, secrets: store.openRepositorySecrets(data) }, { signal: controller.signal, endpoints });
      res.json(await connector.test());
    } catch (err) {
      const message = controller.signal.aborted ? 'Tempo esgotado ao testar a conexão (60 s).' : friendlyError(err);
      res.json({ ok: false, message, details: [] });
    } finally {
      clearTimeout(timer);
    }
  }

  return router;
}
