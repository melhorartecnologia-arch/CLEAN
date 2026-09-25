// Rotas /api/scans: análises, resultados (com filtros), resumo e exportações.
import { Router } from 'express';
import { HttpError } from './validate.js';
import { ScanError } from '../scan/manager.js';
import { filterRecords, publicRecord, summarize } from '../report/model.js';
import { exportXlsx, exportCsv, exportHtml } from '../report/exports.js';

const listFields = (scan) => {
  const { log, ...rest } = scan;
  return rest;
};

/** Nome de arquivo seguro para download (sem acentos nem caracteres especiais). */
function downloadName(scan, ext) {
  const base = scan.name
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .replace(/[^\w.-]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .slice(0, 80);
  return `relatorio-${base || scan.id}.${ext}`;
}

function attachment(res, filename, type) {
  res.setHeader('Content-Type', type);
  res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
}

export function scansRouter({ store, manager }) {
  const router = Router();

  const getScan = (req) => {
    const scan = store.getScan(req.params.id);
    if (!scan) throw new HttpError(404, 'Análise não encontrada.');
    return scan;
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
    res.status(204).end();
  });

  router.get('/:id/results', async (req, res) => {
    const scan = getScan(req);
    const records = await store.readResults(scan.id);
    const filtered = filterRecords(records, req.query);
    const pageSize = Math.min(Math.max(Number.parseInt(req.query.pageSize, 10) || 50, 1), 500);
    const pages = Math.max(1, Math.ceil(filtered.length / pageSize));
    const page = Math.min(Math.max(Number.parseInt(req.query.page, 10) || 1, 1), pages);
    res.json({
      total: filtered.length,
      totalAll: records.length,
      page,
      pages,
      pageSize,
      items: filtered.slice((page - 1) * pageSize, page * pageSize).map(publicRecord),
    });
  });

  router.get('/:id/summary', async (req, res) => {
    const scan = getScan(req);
    const records = await store.readResults(scan.id);
    const summary = summarize(records);
    const extensions = [...new Set(records.map((r) => r.extension))].filter(Boolean).sort();
    res.json({ ...summary, extensions });
  });

  router.get('/:id/errors', async (req, res) => {
    const scan = getScan(req);
    const errors = await store.readErrors(scan.id);
    const limit = Math.min(Math.max(Number.parseInt(req.query.limit, 10) || 200, 1), 5000);
    res.json({ total: errors.length, items: errors.slice(0, limit) });
  });

  router.get('/:id/export.xlsx', async (req, res) => {
    const scan = getScan(req);
    const [records, errors] = await Promise.all([store.readResults(scan.id), store.readErrors(scan.id)]);
    const data = exportXlsx(scan, filterRecords(records, req.query), errors);
    attachment(res, downloadName(scan, 'xlsx'), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(data);
  });

  router.get('/:id/export.csv', async (req, res) => {
    const scan = getScan(req);
    const records = filterRecords(await store.readResults(scan.id), req.query);
    attachment(res, downloadName(scan, 'csv'), 'text/csv; charset=utf-8');
    res.send(exportCsv(records));
  });

  router.get('/:id/export.html', async (req, res) => {
    const scan = getScan(req);
    const records = filterRecords(await store.readResults(scan.id), req.query);
    attachment(res, downloadName(scan, 'html'), 'text/html; charset=utf-8');
    res.send(exportHtml(scan, records));
  });

  router.get('/:id/export.json', async (req, res) => {
    const scan = getScan(req);
    const records = filterRecords(await store.readResults(scan.id), req.query);
    attachment(res, downloadName(scan, 'json'), 'application/json; charset=utf-8');
    res.send(JSON.stringify({ scan, results: records.map(publicRecord) }, null, 2));
  });

  return router;
}
