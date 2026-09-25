// Rotas /api/scans: análises, resultados (com filtros), resumo e exportações.
import { Router } from 'express';
import { HttpError } from './validate.js';
import { ScanError } from '../scan/manager.js';
import {
  filterRecords,
  publicRecord,
  summarize,
  FILTER_KEYS,
  filterMailRecords,
  summarizeMail,
  MAIL_FILTER_KEYS,
  applyDeletions,
  deletionTotals,
  DELETION_LABELS,
  MAIL_DELETION_LABELS,
} from '../report/model.js';
import { deleteFile, deletionEvent, cleanPaths, keptPaths, isCloudRepo } from '../scan/delete.js';
import { DrivesConnector, keptCloud, keptCloudTarget, coveredByRepo, cloudTarget } from '../cloud/drives.js';
import { compileExclusions, DEFAULT_EXCLUDES } from '../scan/walker.js';
import { createConnector } from '../mail/connectors.js';
import { normalizeAddress } from '../mail/common.js';
import { friendlyError } from '../scan/errors.js';
import { PROJECT_ROOT } from '../config.js';
import { exportXlsx, exportCsv, exportHtml, exportJson } from '../report/exports.js';
import { exportMailXlsx, exportMailCsv, exportMailHtml } from '../report/mail-exports.js';

const byName = (a, b) => a.localeCompare(b, 'pt-BR');

/** Filtros, resumos e exportações de cada tipo de análise (arquivos ou e-mail). */
const MODELS = {
  files: {
    keys: FILTER_KEYS,
    filter: filterRecords,
    summarize,
    options: (records) => {
      const all = summarize(records);
      return {
        terms: [...new Set(all.byTerm.map((t) => t.term))].sort(byName),
        users: all.byUser.filter((u) => u.identified).map((u) => u.user).sort(byName),
        extensions: [...new Set(records.map((r) => r.extension).filter(Boolean))].sort(),
      };
    },
    xlsx: exportXlsx,
    csv: exportCsv,
    html: exportHtml,
  },
  mail: {
    keys: MAIL_FILTER_KEYS,
    filter: filterMailRecords,
    summarize: summarizeMail,
    options: (records) => {
      const all = summarizeMail(records);
      return {
        terms: [...new Set(all.byTerm.map((t) => t.term))].sort(byName),
        mailboxes: all.byMailbox.map((m) => m.mailbox).sort(byName),
        // Os remetentes mais frequentes (a lista completa pode ter milhares de endereços).
        senders: all.bySender.filter((s) => s.sender).slice(0, 500).map((s) => ({ value: s.sender, label: s.label })).sort((a, b) => byName(a.label, b.label)),
        sources: all.bySource.map((s) => ({ value: s.sourceId, label: s.source })),
      };
    },
    xlsx: exportMailXlsx,
    csv: exportMailCsv,
    html: exportMailHtml,
  },
};

const modelOf = (scan) => (scan.kind === 'mail' ? MODELS.mail : MODELS.files);

const listFields = (scan) => {
  const { log, ...rest } = scan;
  return rest;
};

/** Nome de arquivo seguro para download (sem acentos nem caracteres especiais). */
function downloadName(scan, ext) {
  const base = scan.name
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^\w.-]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .slice(0, 80);
  return `relatorio-${base || scan.id}.${ext}`;
}

/** Apenas os filtros conhecidos, como texto (parâmetros repetidos usam o primeiro valor). */
function readFilters(query, keys = FILTER_KEYS) {
  const filters = {};
  for (const key of keys) {
    const raw = Array.isArray(query[key]) ? query[key][0] : query[key];
    if (typeof raw === 'string' && raw) filters[key] = raw.slice(0, 500);
  }
  return filters;
}

/**
 * Memoriza filtros e resumos por análise: enquanto os resultados não mudam, paginar ou alternar
 * entre resultados e resumo não refaz a filtragem e a ordenação de todos os registros.
 */
class Memo {
  constructor(size = 24) {
    this.size = size;
    this.map = new Map();
  }

  get(key, compute) {
    if (this.map.has(key)) {
      const value = this.map.get(key);
      this.map.delete(key);
      this.map.set(key, value);
      return value;
    }
    const value = compute();
    this.map.set(key, value);
    while (this.map.size > this.size) this.map.delete(this.map.keys().next().value);
    return value;
  }

  forget(id) {
    for (const key of [...this.map.keys()]) if (key.startsWith(`${id}|`)) this.map.delete(key);
  }
}

async function stream(res, filename, type, write) {
  res.setHeader('Content-Type', type);
  res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
  try {
    await write(res);
    res.end();
  } catch (err) {
    if (!res.headersSent) throw err;
    console.error('[CLEAN] Falha ao gerar a exportação:', err);
    res.destroy(err);
  }
}

/**
 * Quem fez a ação (exclusão manual ou início de uma análise com exclusão): o usuário da
 * autenticação, se houver, e o endereço de acesso (atrás de um proxy na mesma máquina, o do
 * navegador, informado pelo proxy).
 */
export function actor(req) {
  const ip = String(req.ip || req.socket?.remoteAddress || '').replace(/^::ffff:/, '');
  const where = ip === '127.0.0.1' || ip === '::1' ? 'acesso local' : `acesso de ${ip}`;
  return req.cleanUser ? `${req.cleanUser} (${where})` : where;
}

/** A caixa da mensagem como está no cadastro (no IMAP, com o login configurado para ela). */
function mailboxFor(source, record) {
  const address = normalizeAddress(record.mailbox);
  const box = (source.mailboxes || []).find((m) => normalizeAddress(m.address) === address);
  return { address: record.mailbox, name: record.mailboxName || '', ...(box?.login ? { login: box.login } : {}) };
}

const METHOD_TEXT = { permanent: 'exclusão definitiva', trash: 'mover para a lixeira', file: 'exclusão definitiva' };
const EXCLUDED_NOW = (repo) => `O arquivo está numa pasta (ou tem um nome) que o repositório "${repo.name}" passou a ignorar: ele não é excluído pelo relatório.`;

/**
 * O arquivo está numa pasta (ou tem um nome) que o repositório passou a ignorar depois da análise:
 * ele não é excluído pelo relatório. No OneDrive/SharePoint, vale também o nome da biblioteca.
 */
function excludedNow(repo, record) {
  const isExcluded = compileExclusions([...DEFAULT_EXCLUDES, ...(repo.exclude || [])]);
  const parts = String(record.relativePath || record.name || '')
    .split(/[\\/]+/)
    .filter(Boolean);
  const check = (prefix) => parts.some((name, i) => isExcluded(name, [...prefix, ...parts.slice(0, i + 1)].join('/')));
  if (!record.cloud) return check([]);
  const library = record.cloud.library || '';
  return isExcluded(library, library) || check([]) || check([library]);
}

/**
 * O repositório mudou depois da análise: de tipo (pasta ↔ OneDrive/SharePoint), de locatário ou de
 * alcance (a conta ou o site do arquivo saiu da lista ou passou a ser ignorado).
 */
function cloudChanged(record, repo) {
  if (Boolean(record.cloud) !== isCloudRepo(repo)) return true;
  if (!record.cloud) return false;
  if (record.cloud.kind !== repo.type) return true;
  if (String(record.cloud.tenant || '').toLowerCase() !== String(repo.graph?.tenantId || '').toLowerCase()) return true;
  return !coveredByRepo(repo, record.cloud);
}

export function scansRouter({ store, manager, endpoints = {} }) {
  const router = Router();
  const memo = new Memo();
  // Itens com exclusão manual em andamento ("análise:item"): um segundo pedido é recusado.
  const deleting = new Set();
  const deletingIn = (scanId) => [...deleting].some((key) => key.startsWith(`${scanId}:`));

  /**
   * Como cada item pode ser excluído ({ method }) ou por que não pode ({ blocked: 'removed' quando o
   * repositório/conexão saiu do cadastro, 'not-allowed' sem "Permitir exclusão" }).
   */
  const deletionTarget = (scan, record) => {
    const target = scan.kind === 'mail' ? store.getMailSource(record.sourceId) : store.getRepository(record.repositoryId);
    if (!target) return { blocked: 'removed' };
    if (scan.kind !== 'mail' && cloudChanged(record, target)) return { blocked: 'changed' };
    if (scan.kind !== 'mail' && excludedNow(target, record)) return { blocked: 'excluded' };
    if (!target.allowDelete) return { blocked: 'not-allowed' };
    if (scan.kind === 'mail') return { method: target.deleteMode === 'trash' ? 'trash' : 'permanent' };
    if (record.cloud) return { method: target.deleteMode === 'permanent' ? 'permanent' : 'trash' };
    return { method: 'file' };
  };

  const getScan = (req) => {
    const scan = store.getScan(req.params.id);
    if (!scan) throw new HttpError(404, 'Análise não encontrada.');
    return scan;
  };

  /** Registros filtrados e ordenados (a chave inclui a quantidade de registros lidos). */
  const filtered = async (scan, filters) => {
    const [records, deletions] = await Promise.all([store.readResults(scan.id), store.readDeletions(scan.id)]);
    applyDeletions(records, deletions);
    const { page, pageSize, ...criteria } = filters;
    const key = `${scan.id}|${records.length}|${deletions.length}|f|${JSON.stringify(criteria)}`;
    return { records, deletions, list: memo.get(key, () => modelOf(scan).filter(records, criteria)) };
  };
  const filtersOf = (scan, query) => readFilters(query, modelOf(scan).keys);

  router.get('/', (req, res) => {
    const kind = req.query.kind === 'mail' || req.query.kind === 'files' ? req.query.kind : '';
    const scans = store
      .listScans()
      .filter((scan) => !kind || (scan.kind || 'files') === kind)
      .map(listFields)
      .sort((a, b) => String(b.createdAt).localeCompare(String(a.createdAt)));
    res.json(scans);
  });

  router.post('/', async (req, res) => {
    try {
      const scan = await manager.start(req.body || {}, { by: actor(req) });
      res.status(201).json(scan);
    } catch (err) {
      if (err instanceof ScanError) throw new HttpError(err.status, err.message);
      throw err;
    }
  });

  router.get('/:id', (req, res) => {
    res.json(getScan(req));
  });

  router.post('/:id/cancel', (req, res) => {
    const scan = getScan(req);
    if (!manager.cancel(scan.id)) throw new HttpError(409, 'A análise não está em andamento.');
    res.json(store.getScan(scan.id));
  });

  router.delete('/:id', async (req, res) => {
    const scan = getScan(req);
    if (manager.isActive(scan.id)) throw new HttpError(409, 'Cancele a análise antes de excluí-la.');
    if (deletingIn(scan.id)) throw new HttpError(409, 'Há uma exclusão de item em andamento neste relatório. Tente de novo em instantes.');
    await store.deleteScan(scan.id);
    memo.forget(scan.id);
    res.status(204).end();
  });

  router.get('/:id/results', async (req, res) => {
    const scan = getScan(req);
    const filters = filtersOf(scan, req.query);
    const { records, list } = await filtered(scan, filters);
    const pageSize = Math.min(Math.max(Number.parseInt(req.query.pageSize, 10) || 50, 1), 500);
    const pages = Math.max(1, Math.ceil(list.length / pageSize));
    const page = Math.min(Math.max(Number.parseInt(filters.page, 10) || 1, 1), pages);
    res.json({
      total: list.length,
      totalAll: records.length,
      page,
      pages,
      pageSize,
      items: list.slice((page - 1) * pageSize, page * pageSize).map((r) => {
        const target = deletionTarget(scan, r);
        const gone = r.deletion?.status === 'deleted' || r.deletion?.status === 'missing';
        const inProgress = deleting.has(`${scan.id}:${r.id}`);
        return {
          ...publicRecord(r),
          canDelete: Boolean(target.method) && !gone && !inProgress && !manager.isActive(scan.id),
          deleting: inProgress,
          deleteMethod: target.method || null,
          deleteBlocked: target.blocked || null,
        };
      }),
    });
  });

  // Exclusão manual de um item do relatório (arquivo ou mensagem), com registro de quem excluiu.
  router.post('/:id/results/:rid/delete', async (req, res) => {
    const scan = getScan(req);
    if (manager.isActive(scan.id)) throw new HttpError(409, 'Aguarde o fim da análise para excluir itens pelo relatório.');
    if (req.body?.confirm !== true) throw new HttpError(400, 'Confirme a exclusão.');
    const rid = Number(req.params.rid);
    const key = `${scan.id}:${rid}`;
    if (deleting.has(key)) throw new HttpError(409, 'A exclusão deste item já está em andamento.');
    deleting.add(key);
    try {
      await deleteItem(req, res, scan, rid);
    } finally {
      deleting.delete(key);
    }
  });

  async function deleteItem(req, res, scan, rid) {
    const [records, deletions] = await Promise.all([store.readResults(scan.id), store.readDeletions(scan.id)]);
    applyDeletions(records, deletions);
    const record = records.find((r) => r.id === rid);
    if (!record) throw new HttpError(404, 'Item não encontrado nesta análise.');
    if (record.deletion?.status === 'deleted') throw new HttpError(409, 'Este item já foi excluído.');
    const mail = scan.kind === 'mail';
    // O modo mostrado na confirmação precisa ser o que vai ser usado (o cadastro pode ter mudado).
    const checkMethod = (method) => {
      const shown = req.body?.method;
      if (shown !== undefined && shown !== method) {
        throw new HttpError(409, `A forma de exclusão mudou no cadastro (agora: ${METHOD_TEXT[method]}). Confira e confirme de novo.`, 'method-changed');
      }
    };
    const by = actor(req);
    let method;
    let result;
    let label;
    let item;
    if (mail) {
      const source = store.getMailSource(record.sourceId);
      if (!source) throw new HttpError(409, 'A conexão de e-mail desta mensagem foi excluída do cadastro.');
      if (!source.allowDelete) throw new HttpError(403, `A exclusão não está permitida na conexão "${source.name}". Ative "Permitir exclusão" em Caixas de e-mail.`);
      method = source.deleteMode === 'trash' ? 'trash' : 'permanent';
      checkMethod(method);
      label = `da mensagem "${record.subject || '(sem assunto)'}" da caixa ${record.mailbox}`;
      item = `${record.mailbox} › ${record.folder} › ${record.subject || '(sem assunto)'}`;
      let secrets;
      try {
        secrets = store.openMailSecrets(source);
      } catch (err) {
        throw new HttpError(409, err.message);
      }
      const connector = createConnector({ ...source, secrets }, { signal: AbortSignal.timeout(120000), endpoints });
      try {
        const map = await connector.deleteMessages(mailboxFor(source, record), [{ id: record.messageId, messageId: record.internetMessageId }], method);
        const r = map.get(record.messageId) || { ok: false, error: 'O servidor não confirmou a exclusão.' };
        result = { status: r.ok ? 'deleted' : r.missing ? 'missing' : 'failed', error: r.ok ? null : r.error, note: r.note };
      } catch (err) {
        result = { status: 'failed', error: friendlyError(err) };
      } finally {
        await connector.close?.();
      }
    } else if (record.cloud) {
      const repo = store.getRepository(record.repositoryId);
      if (!repo) throw new HttpError(409, 'O repositório deste arquivo foi excluído do cadastro.');
      if (cloudChanged(record, repo)) {
        throw new HttpError(409, `O cadastro do repositório "${repo.name}" mudou depois da análise (tipo, locatário, contas ou sites): faça uma nova análise para excluir.`);
      }
      if (!repo.allowDelete) throw new HttpError(403, `A exclusão não está permitida no repositório "${repo.name}". Ative "Permitir exclusão" em Repositórios.`);
      if (excludedNow(repo, record)) throw new HttpError(409, EXCLUDED_NOW(repo));
      method = repo.deleteMode === 'permanent' ? 'permanent' : 'trash';
      checkMethod(method);
      label = `do arquivo ${record.path}`;
      item = record.path;
      let secrets;
      try {
        secrets = store.openRepositorySecrets(repo);
      } catch (err) {
        throw new HttpError(409, err.message);
      }
      const connector = new DrivesConnector({ ...repo, secrets }, { signal: AbortSignal.timeout(120000), endpoints });
      const kept = keptCloudTarget(record.cloud, await connector.resolveKept(keptCloud(repo, store.listRepositories())));
      result = kept ? { status: 'failed', error: kept.error } : await connector.deleteItem(cloudTarget(record), method, { force: req.body?.force === true });
      if (result.status === 'changed') return res.status(409).json({ error: `${result.error} Confirme para excluir mesmo assim.`, code: 'changed' });
    } else {
      const repo = store.getRepository(record.repositoryId);
      if (!repo) throw new HttpError(409, 'O repositório deste arquivo foi excluído do cadastro.');
      if (cloudChanged(record, repo)) throw new HttpError(409, `O cadastro do repositório "${repo.name}" mudou depois da análise (agora é ${repo.type === 'sharepoint' ? 'SharePoint' : 'OneDrive'}): faça uma nova análise para excluir.`);
      if (!repo.allowDelete) throw new HttpError(403, `A exclusão não está permitida no repositório "${repo.name}". Ative "Permitir exclusão" em Repositórios.`);
      if (excludedNow(repo, record)) throw new HttpError(409, EXCLUDED_NOW(repo));
      method = 'file';
      checkMethod(method);
      label = `do arquivo ${record.path}`;
      item = record.path;
      result = await deleteFile(record.path, {
        root: repo.path,
        expected: { size: record.size, modified: record.modified },
        force: req.body?.force === true,
        protect: [...cleanPaths({ dataDir: store.dataDir, appDir: PROJECT_ROOT }), ...keptPaths(repo, store.listRepositories())],
      });
      if (result.status === 'changed') return res.status(409).json({ error: `${result.error} Confirme para excluir mesmo assim.`, code: 'changed' });
    }
    const event = deletionEvent(record.id, result, { mode: 'manual', method, by, item });
    const labels = mail ? MAIL_DELETION_LABELS : DELETION_LABELS;
    const outcome = result.status === 'failed' ? `falhou: ${result.error}` : `${labels[result.status].toLowerCase()}${result.note ? ` (${result.note})` : ''}`;
    memo.forget(scan.id);
    try {
      await store.appendDeletions(scan.id, [event]);
    } catch (err) {
      console.error('[CLEAN] Falha ao gravar o registro de exclusões:', err.message);
      store.appendLog(scan.id, { level: 'error', message: `Exclusão manual ${label} por ${by}: ${outcome}. Falha ao gravar o registro da exclusão: ${err.message}` });
      return res.status(500).json({ error: `Resultado: ${outcome}. Mas houve uma falha ao gravar o registro da exclusão (${err.message}); a ação ficou anotada no Registro da análise.`, deletion: event });
    }
    store.appendLog(scan.id, { level: result.status === 'failed' ? 'warn' : 'info', message: `Exclusão manual ${label} por ${by}: ${outcome}.` });
    res.json({ deletion: event });
  }

  // Resumo do recorte filtrado (gráficos) e opções de filtro calculadas sobre todos os resultados.
  router.get('/:id/summary', async (req, res) => {
    const scan = getScan(req);
    const model = modelOf(scan);
    const filters = filtersOf(scan, req.query);
    const { records, deletions, list } = await filtered(scan, filters);
    const { page, pageSize, sort, dir, ...criteria } = filters;
    const summary = memo.get(`${scan.id}|${records.length}|${deletions.length}|s|${JSON.stringify(criteria)}`, () => model.summarize(list));
    const options = memo.get(`${scan.id}|${records.length}|o`, () => model.options(records));
    res.json({ ...summary, options, deletions: deletionTotals(records) });
  });

  router.get('/:id/errors', async (req, res) => {
    const scan = getScan(req);
    const errors = await store.readErrors(scan.id);
    const limit = Math.min(Math.max(Number.parseInt(req.query.limit, 10) || 200, 1), 5000);
    res.json({ total: errors.length, items: errors.slice(0, limit) });
  });

  router.get('/:id/export.xlsx', async (req, res) => {
    const scan = getScan(req);
    const { records, deletions, list } = await filtered(scan, filtersOf(scan, req.query));
    const errors = await store.readErrors(scan.id);
    await stream(res, downloadName(scan, 'xlsx'), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', (out) =>
      modelOf(scan).xlsx(scan, list, errors, out, { deletions, records }),
    );
  });

  router.get('/:id/export.csv', async (req, res) => {
    const scan = getScan(req);
    const { list } = await filtered(scan, filtersOf(scan, req.query));
    await stream(res, downloadName(scan, 'csv'), 'text/csv; charset=utf-8', (out) => modelOf(scan).csv(list, out));
  });

  router.get('/:id/export.html', async (req, res) => {
    const scan = getScan(req);
    const { list } = await filtered(scan, filtersOf(scan, req.query));
    await stream(res, downloadName(scan, 'html'), 'text/html; charset=utf-8', (out) => modelOf(scan).html(scan, list, out));
  });

  router.get('/:id/export.json', async (req, res) => {
    const scan = getScan(req);
    const { list } = await filtered(scan, filtersOf(scan, req.query));
    await stream(res, downloadName(scan, 'json'), 'application/json; charset=utf-8', (out) => exportJson(scan, list, out));
  });

  return router;
}
