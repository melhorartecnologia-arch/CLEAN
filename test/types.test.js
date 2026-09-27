// Busca por tipo de arquivo: catálogo, tipo real pelo conteúdo, motor (com exclusão automática e
// limite), API (relatório, filtros, exportações e agendamentos) e exclusão em lote pelo relatório.
import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { sanitizeFileTypes, typeMatcher, normalizeExtension, describeFileTypes, FileTypesError } from '../src/types/catalog.js';
import { detectType, sameContainer } from '../src/types/signature.js';
import { Scanner } from '../src/scan/scanner.js';
import { Store } from '../src/store.js';
import { ScanManager } from '../src/scan/manager.js';
import { Scheduler } from '../src/schedule/scheduler.js';
import { createApp } from '../src/app.js';

let root;
const stores = [];

before(() => {
  root = fs.mkdtempSync(path.join(os.tmpdir(), 'clean-types-'));
});

after(async () => {
  await Promise.all(stores.map((s) => s.close()));
  fs.rmSync(root, { recursive: true, force: true });
});

const MP4 = Buffer.concat([Buffer.from('000000186674797069736f6d', 'hex'), Buffer.alloc(30 * 1024)]);
const JPEG = Buffer.concat([Buffer.from('ffd8ffe000104a464946', 'hex'), Buffer.alloc(4096)]);

/** Pasta com arquivos de vários tipos (inclusive uma foto renomeada para .pdf). */
function tree() {
  const dir = fs.mkdtempSync(path.join(root, 'repo-'));
  const put = (rel, content) => {
    const target = path.join(dir, rel);
    fs.mkdirSync(path.dirname(target), { recursive: true });
    fs.writeFileSync(target, content);
  };
  put('Filmes/ferias.mp4', MP4);
  put('Musicas/samba.MP3', Buffer.from(`ID3\x03${'x'.repeat(200)}`, 'latin1'));
  put('RH/foto-renomeada.pdf', JPEG);
  put('RH/relatorio.pdf', '%PDF-1.7 relatório');
  put('Instaladores/setup.exe', Buffer.from(`MZ${'x'.repeat(100)}`, 'latin1'));
  put('backup.tar.gz', Buffer.from('1f8b0800', 'hex'));
  put('notas.txt', 'texto');
  return dir;
}

test('catálogo: validação, extensões compostas, maiúsculas e descrição', () => {
  const fails = (input, pattern) => assert.throws(() => sanitizeFileTypes(input), (err) => err instanceof FileTypesError && pattern.test(err.message));
  fails({}, /Escolha ao menos um tipo/);
  fails({ categories: ['filmes'] }, /Tipo de arquivo desconhecido: filmes/);
  fails({ extensions: 'mp3, ../x' }, /Extensão inválida: \.\.\/x/);
  fails({ categories: ['video'], minSizeMB: -1 }, /tamanho mínimo/);
  fails({ categories: ['video'], maxDeletions: 1.5 }, /limite de exclusões/);
  assert.equal(normalizeExtension('*.MP3'), '.mp3');
  assert.equal(normalizeExtension('tar.gz'), '.tar.gz');
  assert.equal(normalizeExtension('.'), null);
  const t = sanitizeFileTypes({ categories: ['audio', 'video', 'audio'], extensions: 'XYZ; .tar.gz', minSizeMB: '1,5' });
  assert.deepEqual(t, { categories: ['video', 'audio'], extensions: ['.tar.gz', '.xyz'], checkContent: false, minSizeMB: 1.5, maxDeletions: 1000 });
  const match = typeMatcher(t);
  assert.deepEqual(match('Férias.MP4'), { category: 'video', extension: '.mp4' });
  assert.deepEqual(match('backup.TAR.GZ'), { category: 'custom', extension: '.tar.gz' });
  assert.equal(match('arquivo.gz'), null, '.gz sozinho não foi escolhido');
  assert.equal(match('.mp3'), null, 'arquivo oculto sem nome');
  assert.equal(match('sem-extensao'), null);
  assert.equal(describeFileTypes(t), 'Vídeos, Músicas e áudio, .tar.gz e .xyz (a partir de 1,5 MB)');
});

test('tipo real pelo conteúdo: assinaturas e casos ambíguos', () => {
  assert.equal(detectType(MP4).category, 'video');
  assert.equal(detectType(JPEG).format, 'Imagem JPEG');
  assert.equal(detectType(Buffer.from('%PDF-1.4')).category, 'document');
  assert.equal(detectType(Buffer.from('MZ\x90\x00', 'latin1')).category, 'executable');
  const zip = (first) => {
    const b = Buffer.alloc(128);
    b.write('PK\x03\x04', 0, 'latin1');
    b.writeUInt16LE(first.length, 26);
    b.write(first, 30, 'latin1');
    return b;
  };
  assert.equal(detectType(zip('fotos/01.jpg')).category, 'archive');
  assert.equal(detectType(zip('[Content_Types].xml')).category, null, 'documento do Office não é "compactado"');
  assert.equal(detectType(zip('META-INF/MANIFEST.MF')).category, 'executable');
  assert.equal(detectType(Buffer.from('Relatório de vendas', 'utf8')), null);
  const bmp = Buffer.alloc(64);
  bmp.write('BM', 0, 'latin1');
  bmp.writeUInt32LE(64, 2);
  assert.equal(detectType(bmp, 64).category, 'image');
  assert.equal(detectType(bmp, 999), null, 'BMP só com o tamanho certo no cabeçalho');
  // Textos não são confundidos com assinaturas curtas (nem o UTF-16, que começa com FF FE).
  assert.equal(detectType(Buffer.from('\ufeffNome;Valor\r\n', 'utf16le')), null, 'texto UTF-16 não é MP3');
  assert.equal(detectType(Buffer.from('MZ-2024: vendas do trimestre')), null, 'texto começando com MZ');
  assert.equal(detectType(Buffer.from('ID3 tags da coleção')), null, 'texto começando com ID3');
  // O formato próprio de outras extensões não é um arquivo renomeado.
  const odd = detectType(zip('customXml/item1.xml'));
  assert.equal(odd.category, 'archive');
  assert.equal(sameContainer(odd, '.xlsx'), true);
  assert.equal(sameContainer(odd, '.pdf'), false);
  assert.equal(sameContainer(detectType(MP4), '.m4a'), true, 'um .m4a é MP4 por dentro');
});

async function runTypes(dir, fileTypes, { deleteMatches = false, keep = [] } = {}) {
  const messages = [];
  const stats = await new Scanner(
    {
      repositories: [{ id: 'r', name: 'Dados', path: dir, exclude: [], allowDelete: true, keep }],
      terms: [],
      options: { deleteMatches, resolveOwner: false, checkName: false, checkContent: false, concurrency: 1 },
      fileTypes: sanitizeFileTypes(fileTypes),
    },
    (m) => messages.push(m),
  ).run();
  return {
    stats,
    records: messages.filter((m) => m.type === 'results').flatMap((m) => m.records),
    events: messages.filter((m) => m.type === 'deletions').flatMap((m) => m.items),
    logs: messages.filter((m) => m.type === 'log').map((m) => m.message),
  };
}

test('motor: pela extensão, pelo tipo real, tamanho mínimo e exclusão automática com limite', async () => {
  let dir = tree();
  let run = await runTypes(dir, { categories: ['video', 'audio', 'image', 'executable'], extensions: ['.tar.gz'] });
  assert.deepEqual(run.records.map((r) => r.relativePath.replaceAll('\\', '/')).sort(), ['Filmes/ferias.mp4', 'Instaladores/setup.exe', 'Musicas/samba.MP3', 'backup.tar.gz']);
  assert.equal(run.stats.filesSeen, 7);
  assert.ok(run.records.every((r) => r.typeMatch.by === 'extension' && r.terms.length === 0));
  assert.match(run.logs[0], /Busca por tipo iniciada em 1 repositório\(s\): Vídeos, Músicas e áudio, Imagens e fotos, Executáveis e instaladores e \.tar\.gz\./);

  // Tipo real: a foto renomeada para .pdf aparece (o PDF de verdade, não).
  run = await runTypes(dir, { categories: ['image'], checkContent: true });
  assert.deepEqual(run.records.map((r) => [r.name, r.typeMatch.by, r.typeMatch.format]), [['foto-renomeada.pdf', 'content', 'Imagem JPEG']]);
  assert.equal(run.stats.typesByContent, 1);

  // Tipo real só nos arquivos sem extensão ou com a extensão de outro tipo conhecido: um .dll (fora do
  // catálogo) é o que diz ser, e o formato próprio da extensão (.xlsx é ZIP, .m4a é MP4) não conta.
  const other = fs.mkdtempSync(path.join(root, 'repo-'));
  const oddZip = Buffer.concat([Buffer.from('504b0304', 'hex'), Buffer.alloc(200)]);
  fs.writeFileSync(path.join(other, 'planilha.xlsx'), oddZip);
  fs.writeFileSync(path.join(other, 'fotos.pdf'), oddZip);
  fs.writeFileSync(path.join(other, 'biblioteca.dll'), Buffer.concat([Buffer.from('MZ\x90\x00', 'latin1'), Buffer.alloc(100)]));
  fs.writeFileSync(path.join(other, 'musica.m4a'), MP4);
  fs.writeFileSync(path.join(other, 'video-sem-extensao'), MP4);
  run = await runTypes(other, { categories: ['video', 'archive', 'executable'], checkContent: true });
  assert.deepEqual(run.records.map((r) => [r.name, r.typeMatch.category, r.typeMatch.by]).sort(), [
    ['fotos.pdf', 'archive', 'content'],
    ['video-sem-extensao', 'video', 'content'],
  ]);

  // Tamanho mínimo: só o vídeo (30 KB) passa de 0,02 MB (cerca de 21 KB).
  run = await runTypes(dir, { categories: ['video', 'audio'], minSizeMB: 0.02 });
  assert.deepEqual(run.records.map((r) => r.name), ['ferias.mp4']);
  assert.equal(run.stats.filesSkippedBySize, 1);

  // Exclusão automática com limite 1: um excluído, o outro só listado (e o aviso).
  run = await runTypes(dir, { categories: ['video', 'executable'], maxDeletions: 1 }, { deleteMatches: true });
  assert.equal(run.stats.filesMatched, 2);
  assert.equal(run.stats.deleted, 1);
  assert.equal(run.stats.deleteSkipped, 1);
  assert.ok(run.events.every((e) => e.mode === 'auto' && e.method === 'file'));
  assert.ok(run.logs.some((l) => /Limite de 1 exclusão desta execução atingido: 1 arquivo\(s\) encontrado\(s\) só foram listados.*aumente o limite nas opções da busca/.test(l)));

  // Local protegido (repositório sem exclusão dentro do analisado): não é tentado.
  dir = tree();
  run = await runTypes(dir, { categories: ['video', 'audio'] }, { deleteMatches: true, keep: [{ path: path.join(dir, 'Filmes'), error: 'Protegido.' }] });
  assert.ok(fs.existsSync(path.join(dir, 'Filmes/ferias.mp4')));
  assert.ok(!fs.existsSync(path.join(dir, 'Musicas/samba.MP3')));
  assert.equal(run.stats.deleteProtected, 1);
  assert.equal(run.events.filter((e) => e.status === 'failed').length, 0);
});

test('API: busca por tipo, relatório, exportações, exclusão em lote e agendamento', async () => {
  const store = await new Store(path.join(root, `data-${Math.random().toString(36).slice(2)}`)).init();
  stores.push(store);
  const manager = new ScanManager(store);
  const scheduler = new Scheduler({ store, manager });
  const app = createApp({ store, manager, scheduler, config: { authUser: '', authPassword: '' } });
  const server = await new Promise((resolve) => {
    const s = app.listen(0, '127.0.0.1', () => resolve(s));
  });
  const base = `http://127.0.0.1:${server.address().port}`;
  const api = async (method, url, body) => {
    const res = await fetch(base + url, { method, headers: { 'Content-Type': 'application/json', 'X-CLEAN': '1' }, body: body === undefined ? undefined : JSON.stringify(body) });
    return { status: res.status, data: res.status === 204 ? null : await res.json() };
  };
  const waitScan = async (id) => {
    for (let i = 0; i < 200; i++) {
      const { data } = await api('GET', `/api/scans/${id}`);
      if (!['queued', 'running'].includes(data.status)) return data;
      await new Promise((r) => setTimeout(r, 50));
    }
    throw new Error('não terminou');
  };
  const waitBulk = async (id) => {
    for (let i = 0; i < 200; i++) {
      const { data } = await api('GET', `/api/scans/${id}/bulk-delete`);
      if (data && !data.running) return data;
      await new Promise((r) => setTimeout(r, 50));
    }
    throw new Error('o lote não terminou');
  };
  try {
    const dir = tree();
    const repo = (await api('POST', '/api/repositories', { name: 'Dados', path: dir })).data;
    const body = { repositoryIds: [repo.id], fileTypes: { categories: ['video', 'audio', 'image', 'executable'], checkContent: true }, options: { resolveOwner: false } };
    // Com exclusão automática: exige "Permitir exclusão" e EXCLUIR, como nas análises.
    let res = await api('POST', '/api/scans', { ...body, options: { deleteMatches: true } });
    assert.match(res.data.error, /A exclusão não está permitida em "Dados"/);
    res = await api('POST', '/api/scans', body);
    assert.equal(res.status, 201, JSON.stringify(res.data));
    assert.match(res.data.name, /^Busca por tipo de /);
    let scan = await waitScan(res.data.id);
    assert.equal(scan.stats.filesMatched, 4);
    assert.deepEqual(scan.fileTypes.categories, ['video', 'audio', 'image', 'executable']);

    // Resumo por tipo, filtros e ordenação padrão (os maiores primeiro).
    const summary = (await api('GET', `/api/scans/${scan.id}/summary`)).data;
    assert.deepEqual(summary.types.byType.map((g) => [g.key, g.count]), [['video', 1], ['audio', 1], ['image', 1], ['executable', 1]]);
    assert.equal(summary.types.byContent, 1);
    const items = (await api('GET', `/api/scans/${scan.id}/results`)).data.items;
    assert.equal(items[0].name, 'ferias.mp4', 'o maior primeiro');
    assert.equal((await api('GET', `/api/scans/${scan.id}/results?type=image`)).data.total, 1);
    assert.equal((await api('GET', `/api/scans/${scan.id}/results?found=content`)).data.total, 1);
    const csv = (await (await fetch(`${base}/api/scans/${scan.id}/export.csv`)).text()).trim().split('\r\n');
    assert.equal(csv.length, 5);
    assert.match(csv[0], /Extensão;Tipo;Encontrado por;Formato real;Tamanho \(KB\)/);
    assert.ok(csv.some((l) => /foto-renomeada\.pdf;\.pdf;Imagens e fotos;Conteúdo \(tipo real\);Imagem JPEG/.test(l)));
    const html = await (await fetch(`${base}/api/scans/${scan.id}/export.html`)).text();
    assert.match(html, /Relatório CLEAN – busca por tipo/);
    assert.match(html, /Somente procurar \(revisão no relatório\)/);
    assert.equal((await fetch(`${base}/api/scans/${scan.id}/export.xlsx`)).status, 200);

    // Exclusão em lote: sem "Permitir exclusão", a prévia mostra que nada pode ser excluído.
    let preview = (await api('POST', `/api/scans/${scan.id}/bulk-delete/preview`, { all: true, filters: {} })).data;
    assert.deepEqual([preview.total, preview.ready, preview.blocked['not-allowed']], [4, 0, 4]);
    await api('PUT', `/api/repositories/${repo.id}`, { name: 'Dados', path: dir, allowDelete: true });
    preview = (await api('POST', `/api/scans/${scan.id}/bulk-delete/preview`, { all: true, filters: { type: 'video' } })).data;
    assert.deepEqual([preview.total, preview.ready], [1, 1]);
    assert.deepEqual(preview.repositories.map((g) => [g.name, g.method, g.count]), [['Dados', 'file', 1]]);
    assert.match((await api('POST', `/api/scans/${scan.id}/bulk-delete`, { all: true, filters: { type: 'video' } })).data.error, /Digite EXCLUIR/);
    assert.equal((await api('POST', `/api/scans/${scan.id}/bulk-delete`, { all: true, filters: { type: 'video' }, confirmDelete: 'EXCLUIR', methods: { [repo.id]: 'trash' } })).data.code, 'method-changed');
    res = await api('POST', `/api/scans/${scan.id}/bulk-delete`, { all: true, filters: { type: 'video' }, confirmDelete: 'excluir', methods: { [repo.id]: 'file' } });
    assert.equal(res.status, 202, JSON.stringify(res.data));
    let job = await waitBulk(scan.id);
    assert.deepEqual([job.total, job.deleted, job.failed], [1, 1, 0]);
    assert.ok(!fs.existsSync(path.join(dir, 'Filmes/ferias.mp4')));

    // Selecionados (ids), um deles alterado depois da busca: mantido.
    const pick = (await api('GET', `/api/scans/${scan.id}/results?sort=path`)).data.items.filter((r) => ['samba.MP3', 'setup.exe'].includes(r.name));
    fs.appendFileSync(path.join(dir, 'Instaladores/setup.exe'), 'alterado');
    res = await api('POST', `/api/scans/${scan.id}/bulk-delete`, { ids: pick.map((r) => r.id), confirmDelete: 'EXCLUIR' });
    assert.equal(res.status, 202);
    job = await waitBulk(scan.id);
    assert.deepEqual([job.deleted, job.changed], [1, 1]);
    assert.ok(fs.existsSync(path.join(dir, 'Instaladores/setup.exe')));
    const events = fs.readFileSync(path.join(store.dataDir, 'exclusoes.ndjson'), 'utf8').trim().split('\n').map((l) => JSON.parse(l));
    assert.ok(events.every((e) => e.mode === 'manual' && /\(exclusão em lote\)$/.test(e.by)));
    scan = (await api('GET', `/api/scans/${scan.id}`)).data;
    assert.ok(scan.log.some((l) => /Exclusão em lote por acesso local concluída: 1 excluído\(s\), 1 mantido\(s\)/.test(l.message)));
    // Já excluídos não entram de novo.
    preview = (await api('POST', `/api/scans/${scan.id}/bulk-delete/preview`, { all: true, filters: {} })).data;
    assert.equal(preview.total, 2);

    // Falhas seguidas (aqui, cada arquivo virou uma pasta): o lote para em 20, sem tentar o resto.
    const many = fs.mkdtempSync(path.join(root, 'repo-'));
    for (let i = 1; i <= 25; i++) fs.writeFileSync(path.join(many, `video-${String(i).padStart(2, '0')}.mp4`), MP4);
    const manyRepo = (await api('POST', '/api/repositories', { name: 'Vídeos', path: many, allowDelete: true })).data;
    res = await api('POST', '/api/scans', { repositoryIds: [manyRepo.id], fileTypes: { categories: ['video'] }, options: { resolveOwner: false } });
    const manyScan = await waitScan(res.data.id);
    assert.equal(manyScan.stats.filesMatched, 25);
    for (const name of fs.readdirSync(many)) {
      fs.rmSync(path.join(many, name));
      fs.mkdirSync(path.join(many, name));
    }
    res = await api('POST', `/api/scans/${manyScan.id}/bulk-delete`, { all: true, filters: {}, confirmDelete: 'EXCLUIR' });
    assert.equal(res.status, 202, JSON.stringify(res.data));
    job = await waitBulk(manyScan.id);
    assert.deepEqual([job.total, job.done, job.failed, job.deleted], [25, 20, 20, 0]);
    assert.match(job.halted, /^20 falhas seguidas \(a última: O caminho não é um arquivo\.\)/);
    assert.equal(job.methods, undefined, 'detalhes internos do lote não saem na API');
    const manyLog = (await api('GET', `/api/scans/${manyScan.id}`)).data.log;
    assert.ok(manyLog.some((l) => l.level === 'warn' && /Exclusão em lote .* interrompida depois de 20 de 25 por 20 falhas seguidas/.test(l.message)));

    // Agendamento de busca por tipo (aparece em Agendamentos) e troca para termos.
    const schedule = {
      purpose: 'types',
      kind: 'files',
      name: 'Mídia semanal',
      repositoryIds: [repo.id],
      fileTypes: { categories: ['video', 'audio'] },
      options: { resolveOwner: false },
      rule: { frequency: 'weekly', startDate: '2026-09-01', time: '03:00', weekdays: [0] },
      period: { type: 'all' },
    };
    const created = await api('POST', '/api/schedules', schedule);
    assert.equal(created.status, 201, JSON.stringify(created.data));
    assert.equal(created.data.typesText, 'Vídeos e Músicas e áudio');
    assert.equal((await api('GET', '/api/schedules')).data.length, 1);
    assert.equal((await api('GET', '/api/schedules?purpose=retention')).data.length, 0);
    const run = await api('POST', `/api/schedules/${created.data.id}/run`, { confirm: true });
    assert.equal(run.status, 201, JSON.stringify(run.data));
    const fromSchedule = await waitScan(run.data.scan.id);
    assert.deepEqual(fromSchedule.fileTypes.categories, ['video', 'audio']);
    assert.equal(fromSchedule.stats.filesMatched, 0, 'o vídeo e a música já foram excluídos');
    assert.match((await api('PUT', `/api/schedules/${created.data.id}`, { ...schedule, purpose: 'retention' })).data.error, /política de retenção/);
    assert.match((await api('POST', '/api/schedules', { ...schedule, kind: 'mail' })).data.error, /vale para arquivos/);
  } finally {
    server.close();
    await scheduler.stop();
  }
});
