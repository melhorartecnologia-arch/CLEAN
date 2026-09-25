// Rotas /api/scans: análises, resultados (com filtros), resumo e exportações.
import { Router } from 'express';
import { HttpError } from './validate.js';
import { ScanError } from '../scan/manager.js';
import { filterRecords, publicRecord, summarize, FILTER_KEYS, filterMailRecords, summarizeMail, MAIL_FILTER_KEYS, applyDeletions, deletionTotals, DELETION_LABELS } from '../report/model.js';
import { deleteFile, deletionEvent } from '../scan/delete.js';
import { createConnector } from '../mail/connectors.js';
import { friendlyError } from '../scan/errors.js';
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

/** Quem fez a exclusão manual: o usuário da autenticação ou o endereço de acesso. */
function actor(req) {
  if (req.cleanUser) return req.cleanUser;
  const ip = String(req.socket?.remoteAddress || '').replace(/^::ffff:/, '');
  return ip === '127.0.0.1' || ip === '::1' ? 'acesso local' : `acesso de ${ip}`;
}

export function scansRouter({ store, manager, endpoints = {} }) {
  const router = Router();
  const memo = new Memo();

  /** Onde cada item pode ser excluído (repositório/conexão com "Permitir exclusão"). */
  const deletionTarget = (scan, record) => {
    if (scan.kind === 'mail') {
      const source = store.getMailSource(record.sourceId);
      return source?.allowDelete ? { method: source.deleteMode === 'trash' ? 'trash' : 'permanent' } : null;
    }
    return store.getRepository(record.repositoryId)?.allowDelete ? { method: 'file' } : null;
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
      const scan = await manager.start(req.body || {});
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
        const deleted = r.deletion?.status === 'deleted' || r.deletion?.status === 'missing';
        return { ...publicRecord(r), canDelete: Boolean(target) && !deleted && !manager.isActive(scan.id), deleteMethod: target?.method || null };
      }),
    });
  });

  // Exclusão manual de um item do relatório (arquivo ou mensagem), com registro de quem excluiu.
  router.post('/:id/results/:rid/delete', async (req, res) => {
    const scan = getScan(req);
    if (manager.isActive(scan.id)) throw new HttpError(409, 'Aguarde o fim da análise para excluir itens pelo relatório.');
    if (req.body?.confirm !== true) throw new HttpError(400, 'Confirme a exclusão.');
    const rid = Number(req.params.rid);
    const [records, deletions] = await Promise.all([store.readResults(scan.id), store.readDeletions(scan.id)]);
    applyDeletions(records, deletions);
    const record = records.find((r) => r.id === rid);
    if (!record) throw new HttpError(404, 'Item não encontrado nesta análise.');
    if (record.deletion?.status === 'deleted') throw new HttpError(409, 'Este item já foi excluído.');
    const by = actor(req);
    let method;
    let result;
    let label;
    if (scan.kind === 'mail') {
      const source = store.getMailSource(record.sourceId);
      if (!source) throw new HttpError(409, 'A conexão de e-mail desta mensagem foi excluída do cadastro.');
      if (!source.allowDelete) throw new HttpError(403, `A exclusão não está permitida na conexão "${source.name}". Ative "Permitir exclusão" em Caixas de e-mail.`);
      method = source.deleteMode === 'trash' ? 'trash' : 'permanent';
      label = `mensagem "${record.subject || '(sem assunto)'}" da caixa ${record.mailbox}`;
      let secrets;
      try {
        secrets = store.openMailSecrets(source);
      } catch (err) {
        throw new HttpError(409, err.message);
      }
      const connector = createConnector({ ...source, secrets }, { signal: AbortSignal.timeout(120000), endpoints });
      try {
        const map = await connector.deleteMessages({ address: record.mailbox, name: record.mailboxName }, [record.messageId], method);
        const r = map.get(record.messageId) || { ok: false, error: 'O servidor não confirmou a exclusão.' };
        result = { status: r.ok ? 'deleted' : r.missing ? 'missing' : 'failed', error: r.ok ? null : r.error };
      } catch (err) {
        result = { status: 'failed', error: friendlyError(err) };
      } finally {
        await connector.close?.();
      }
    } else {
      const repo = store.getRepository(record.repositoryId);
      if (!repo) throw new HttpError(409, 'O repositório deste arquivo foi excluído do cadastro.');
      if (!repo.allowDelete) throw new HttpError(403, `A exclusão não está permitida no repositório "${repo.name}". Ative "Permitir exclusão" em Repositórios.`);
      method = 'file';
      label = `arquivo ${record.path}`;
      result = await deleteFile(record.path, { root: repo.path, expected: { size: record.size, modified: record.modified }, force: req.body?.force === true });
      if (result.status === 'changed') return res.status(409).json({ error: `${result.error} Confirme para excluir mesmo assim.`, code: 'changed' });
    }
    const event = deletionEvent(record.id, result, { mode: 'manual', method, by });
    await store.appendDeletions(scan.id, [event]);
    const outcome = result.status === 'failed' ? `falhou: ${result.error}` : DELETION_LABELS[result.status].toLowerCase();
    store.appendLog(scan.id, { level: result.status === 'failed' ? 'warn' : 'info', message: `Exclusão manual (${by}) do ${label}: ${outcome}.` });
    memo.forget(scan.id);
    res.json({ deletion: event });
  });

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
