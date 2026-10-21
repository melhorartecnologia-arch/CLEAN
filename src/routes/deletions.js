// Rotas /api/deletions: registro geral de exclusões (todas as exclusões feitas pelo CLEAN, mesmo
// depois que o relatório da análise é excluído), com filtros, paginação, resumo e exportações.
import { Router } from 'express';
import { DeletionLog, parseDeletionFilters, publicEvent } from '../report/deletion-log.js';
import { exportDeletionsXlsx, exportDeletionsCsv, exportDeletionsHtml, exportFileName, HTML_LIMIT } from '../report/deletion-exports.js';

const MAX_PAGE_SIZE = 500;

const positiveInt = (value, fallback) => {
  const n = Number.parseInt(Array.isArray(value) ? value[0] : value, 10);
  return Number.isInteger(n) && n > 0 ? n : fallback;
};

async function stream(res, filename, type, write) {
  res.setHeader('Content-Type', type);
  res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
  try {
    await write(res);
    res.end();
  } catch (err) {
    if (!res.headersSent) throw err;
    console.error('[CLEAN] Falha ao gerar a exportação do registro de exclusões:', err);
    res.destroy(err);
  }
}

const EXPORTS = {
  xlsx: { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', write: exportDeletionsXlsx },
  csv: { type: 'text/csv; charset=utf-8', write: exportDeletionsCsv },
  html: { type: 'text/html; charset=utf-8', write: exportDeletionsHtml, limit: HTML_LIMIT },
};

export function deletionsRouter({ store }) {
  const router = Router();
  const log = new DeletionLog(store);

  // Eventos do recorte, das exclusões mais recentes para as mais antigas.
  router.get('/', async (req, res) => {
    const filters = parseDeletionFilters(req.query);
    const result = await log.query(filters);
    const pageSize = Math.min(positiveInt(req.query.pageSize, 50), MAX_PAGE_SIZE);
    const pages = Math.max(1, Math.ceil(result.total / pageSize));
    const page = Math.min(positiveInt(req.query.page, 1), pages);
    const items = await log.page(result, page, pageSize);
    res.json({ total: result.total, page, pages, pageSize, logBytes: result.size, items: items.map(publicEvent) });
  });

  // Totais e agrupamentos do recorte, e as opções dos filtros (pelo registro inteiro).
  router.get('/summary', async (req, res) => {
    const result = await log.query(parseDeletionFilters(req.query));
    res.json({ ...result.summary, total: result.total, logBytes: result.size, options: result.options });
  });

  for (const [ext, { type, write, limit }] of Object.entries(EXPORTS)) {
    router.get(`/export.${ext}`, async (req, res) => {
      const filters = parseDeletionFilters(req.query);
      const result = await log.query(filters);
      const data = { filters, summary: result.summary, options: result.options, total: result.total, events: log.events(result, { limit }) };
      await stream(res, exportFileName(filters, result.summary, ext), type, (out) => write(data, out));
    });
  }

  return router;
}
