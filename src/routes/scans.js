// Rotas /api/scans: análises, resultados (com filtros), resumo e exportações.
import { Router } from 'express';
import { HttpError } from './validate.js';
import { ScanError } from '../scan/manager.js';
import { filterRecords, publicRecord, summarize, FILTER_KEYS } from '../report/model.js';
import { exportXlsx, exportCsv, exportHtml, exportJson } from '../report/exports.js';

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
function readFilters(query) {
  const filters = {};
  for (const key of FILTER_KEYS) {
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

export function scansRouter({ store, manager }) {
  const router = Router();
  const memo = new Memo();

  const getScan = (req) => {
    const scan = store.getScan(req.params.id);
    if (!scan) throw new HttpError(404, 'Análise não encontrada.');
    return scan;
  };

  /** Registros filtrados e ordenados (a chave inclui a quantidade de registros lidos). */
  const filtered = async (scan, filters) => {
    const records = await store.readResults(scan.id);
    const { page, pageSize, ...criteria } = filters;
    const key = `${scan.id}|${records.length}|f|${JSON.stringify(criteria)}`;
    return { records, list: memo.get(key, () => filterRecords(records, criteria)) };
  };

  router.get('/', (req, res) => {
    const scans = store
      .listScans()
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
    const filters = readFilters(req.query);
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
      items: list.slice((page - 1) * pageSize, page * pageSize).map(publicRecord),
    });
  });

  // Resumo do recorte filtrado (gráficos) e opções de filtro calculadas sobre todos os resultados.
  router.get('/:id/summary', async (req, res) => {
    const scan = getScan(req);
    const filters = readFilters(req.query);
    const { records, list } = await filtered(scan, filters);
    const { page, pageSize, sort, dir, ...criteria } = filters;
    const summary = memo.get(`${scan.id}|${records.length}|s|${JSON.stringify(criteria)}`, () => summarize(list));
    const options = memo.get(`${scan.id}|${records.length}|o`, () => {
      const all = summarize(records);
      return {
        terms: [...new Set(all.byTerm.map((t) => t.term))].sort((a, b) => a.localeCompare(b, 'pt-BR')),
        users: all.byUser.filter((u) => u.identified).map((u) => u.user).sort((a, b) => a.localeCompare(b, 'pt-BR')),
        extensions: [...new Set(records.map((r) => r.extension).filter(Boolean))].sort(),
      };
    });
    res.json({ ...summary, options });
  });

  router.get('/:id/errors', async (req, res) => {
    const scan = getScan(req);
    const errors = await store.readErrors(scan.id);
    const limit = Math.min(Math.max(Number.parseInt(req.query.limit, 10) || 200, 1), 5000);
    res.json({ total: errors.length, items: errors.slice(0, limit) });
  });

  router.get('/:id/export.xlsx', async (req, res) => {
    const scan = getScan(req);
    const { list } = await filtered(scan, readFilters(req.query));
    const errors = await store.readErrors(scan.id);
    await stream(res, downloadName(scan, 'xlsx'), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', (out) =>
      exportXlsx(scan, list, errors, out),
    );
  });

  router.get('/:id/export.csv', async (req, res) => {
    const scan = getScan(req);
    const { list } = await filtered(scan, readFilters(req.query));
    await stream(res, downloadName(scan, 'csv'), 'text/csv; charset=utf-8', (out) => exportCsv(list, out));
  });

  router.get('/:id/export.html', async (req, res) => {
    const scan = getScan(req);
    const { list } = await filtered(scan, readFilters(req.query));
    await stream(res, downloadName(scan, 'html'), 'text/html; charset=utf-8', (out) => exportHtml(scan, list, out));
  });

  router.get('/:id/export.json', async (req, res) => {
    const scan = getScan(req);
    const { list } = await filtered(scan, readFilters(req.query));
    await stream(res, downloadName(scan, 'json'), 'application/json; charset=utf-8', (out) => exportJson(scan, list, out));
  });

  return router;
}
