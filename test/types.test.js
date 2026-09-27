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
import { guardFor, guardMatcher } from '../src/scan/delete.js';

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
/** Início de um arquivo MPEG-4 com a marca dada (ex.: "crx " das fotos CR3 da Canon). */
const ftyp = (brand) => Buffer.concat([Buffer.from('00000018', 'hex'), Buffer.from(`ftyp${brand}`, 'latin1'), Buffer.alloc(2048)]);
/** Executável do Windows: "MZ" e o cabeçalho PE no endereço indicado em 0x3C. */
function peFile() {
  const b = Buffer.alloc(1024);
  b.write('MZ', 0, 'latin1');
  b.writeUInt32LE(0x80, 0x3c);
  b.write('PE\0\0', 0x80, 'latin1');
  return b;
}

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
  assert.equal(detectType(peFile()).category, 'executable');
  assert.equal(detectType(Buffer.concat([Buffer.from('MZ\x90\x00', 'latin1'), Buffer.alloc(200)])), null, 'só "MZ", sem o cabeçalho PE');
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
  // Texto UTF-16 sem BOM com "G" a cada 188 bytes não é vídeo MPEG-TS.
  const utf16 = Buffer.from('G'.padEnd(94, 'x').repeat(12), 'utf16le');
  assert.equal(detectType(utf16), null);
  // O formato próprio de outras extensões não é um arquivo renomeado.
  const odd = detectType(zip('conteudo/relatorio.bin'));
  assert.equal(odd.category, 'archive');
  assert.equal(sameContainer(odd, '.xlsx'), true);
  assert.equal(sameContainer(odd, '.doc', 'document'), true, 'um .doc com conteúdo ZIP é um .docx renomeado');
  assert.equal(sameContainer(odd, '.pdf', 'document'), false);
  const jar = detectType(zip('META-INF/MANIFEST.MF'));
  assert.equal(sameContainer(jar, '.zip', 'archive'), true, 'um .zip assinado (META-INF) continua um .zip');
  assert.equal(sameContainer(jar, '.docx', 'document'), false, 'um aplicativo Java renomeado para .docx');
  assert.equal(sameContainer(detectType(MP4), '.m4a', 'audio'), true, 'um .m4a é MP4 por dentro');
  assert.equal(sameContainer(detectType(MP4), '.mp3', 'audio'), true, 'áudio MP4 com a extensão errada de áudio');
  assert.equal(sameContainer(detectType(MP4), '.pdf', 'document'), false);

  // Marcas do MPEG-4: fotos CR3 e HEIF são imagens; marcas desconhecidas não têm categoria.
  assert.deepEqual([detectType(ftyp('crx ')).category, detectType(ftyp('crx ')).format], ['image', 'Foto RAW da Canon (CR3)']);
  assert.equal(detectType(ftyp('mif2')).category, 'image');
  assert.equal(detectType(ftyp('M4A ')).category, 'audio');
  assert.equal(detectType(ftyp('qt  ')).category, 'video');
  assert.equal(detectType(ftyp('abcd')).category, null, 'marca desconhecida');
  assert.equal(sameContainer(detectType(ftyp('crx ')), '.cr3', 'image'), true);

  // MP3 sem ID3: dois quadros seguidos (MPEG-1, camada III, 128 kbit/s, 44,1 kHz: 417 bytes).
  const frame = Buffer.concat([Buffer.from('fffb9064', 'hex'), Buffer.alloc(413)]);
  assert.equal(detectType(Buffer.concat([frame, frame, frame])).format, 'Áudio MP3');
  assert.equal(detectType(Buffer.concat([frame, Buffer.alloc(600, 0x55)])), null, 'um cabeçalho só é coincidência');
  // GIF com 0x47 ("G") nos bytes 188 e 376 continua GIF.
  const gif = Buffer.alloc(800, 0x01);
  gif.write('GIF89a', 0, 'latin1');
  gif[188] = 0x47;
  gif[376] = 0x47;
  assert.equal(detectType(gif).format, 'Imagem GIF');
});

test('exclusão em lote: a conferência rápida das pastas protegidas dá o mesmo resultado', () => {
  const base = path.join(root, 'protegidas');
  const guards = [
    { path: path.join(base, 'Diretoria'), error: 'Diretoria' },
    { path: path.join(base, 'CLEAN'), except: [path.join(base, 'CLEAN', 'demo')], error: 'CLEAN' },
  ];
  const match = guardMatcher(guards);
  const cases = [
    path.join(base, 'Diretoria', 'ata.docx'),
    path.join(base, 'Diretoria'),
    path.join(base, 'Diretoria2', 'ata.docx'), // mesmo começo de nome, outra pasta
    path.join(base, 'CLEAN', 'data', 'db.json'),
    path.join(base, 'CLEAN', 'demo', 'exemplo.pdf'),
    path.join(base, 'Publico', '..', 'Diretoria', 'x.txt'),
    path.join(base, 'Publico', 'video.mp4'),
  ];
  for (const target of cases) assert.equal(match(target)?.error ?? null, guardFor(guards, target)?.error ?? null, target);
  assert.equal(match(cases[2]), null);
  assert.equal(match(cases[5])?.error, 'Diretoria');
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
  assert.equal(run.records.find((r) => r.name === 'backup.tar.gz').extension, '.tar.gz', 'a extensão procurada, composta');
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
  fs.writeFileSync(path.join(other, 'musica-convertida.mp3'), ftyp('mp42')); // áudio MP4 com extensão de MP3
  fs.writeFileSync(path.join(other, 'IMG_0001.CR3'), ftyp('crx ')); // foto RAW da Canon (MPEG-4 por dentro)
  fs.writeFileSync(path.join(other, 'IMG_0002.HEIC'), ftyp('mif2'));
  fs.writeFileSync(path.join(other, 'video-sem-extensao'), MP4);
  run = await runTypes(other, { categories: ['video', 'archive', 'executable'], checkContent: true });
  assert.deepEqual(run.records.map((r) => [r.name, r.typeMatch.category, r.typeMatch.by]).sort(), [
    ['fotos.pdf', 'archive', 'content'],
    ['video-sem-extensao', 'video', 'content'],
  ]);

  // Exclusão automática: os encontrados só pelo tipo real ficam para a revisão (não são excluídos).
  run = await runTypes(other, { categories: ['video', 'archive'], checkContent: true }, { deleteMatches: true });
  assert.equal(run.stats.deleted, 0);
  assert.equal(run.stats.deleteReview, 2);
  assert.ok(fs.existsSync(path.join(other, 'video-sem-extensao')) && fs.existsSync(path.join(other, 'fotos.pdf')));
  assert.ok(run.logs.some((l) => /2 arquivo\(s\) encontrado\(s\) pelo tipo real \(conteúdo\) não foram excluídos automaticamente/.test(l)));

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
  // Exclusão em lote como a tela faz: prévia e confirmação com o token dela.
  const bulk = async (scanId, target, extra = {}) => {
    const shown = (await api('POST', `/api/scans/${scanId}/bulk-delete/preview`, target)).data;
    return api('POST', `/api/scans/${scanId}/bulk-delete`, { ...target, confirmDelete: 'EXCLUIR', token: shown.token, ...extra });
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
    assert.match((await api('POST', `/api/scans/${scan.id}/bulk-delete`, { all: true, filters: { type: 'video' }, confirmDelete: 'EXCLUIR' })).data.error, /Faça a prévia/);
    assert.equal((await bulk(scan.id, { all: true, filters: { type: 'video' } }, { methods: { [repo.id]: 'trash' } })).data.code, 'method-changed');
    res = await bulk(scan.id, { all: true, filters: { type: 'video' } }, { confirmDelete: 'excluir', methods: { [repo.id]: 'file' } });
    assert.equal(res.status, 202, JSON.stringify(res.data));
    let job = await waitBulk(scan.id);
    assert.deepEqual([job.total, job.deleted, job.failed], [1, 1, 0]);
    assert.ok(!fs.existsSync(path.join(dir, 'Filmes/ferias.mp4')));

    // Selecionados (ids), um deles alterado depois da busca: mantido.
    const pick = (await api('GET', `/api/scans/${scan.id}/results?sort=path`)).data.items.filter((r) => ['samba.MP3', 'setup.exe'].includes(r.name));
    fs.appendFileSync(path.join(dir, 'Instaladores/setup.exe'), 'alterado');
    res = await bulk(scan.id, { ids: pick.map((r) => r.id) });
    assert.equal(res.status, 202);
    job = await waitBulk(scan.id);
    assert.deepEqual([job.deleted, job.changed], [1, 1]);
    assert.ok(fs.existsSync(path.join(dir, 'Instaladores/setup.exe')));
    const events = fs.readFileSync(path.join(store.dataDir, 'exclusoes.ndjson'), 'utf8').trim().split('\n').map((l) => JSON.parse(l));
    assert.ok(events.every((e) => e.mode === 'manual' && /\(exclusão em lote\)$/.test(e.by)));
    scan = (await api('GET', `/api/scans/${scan.id}`)).data;
    assert.ok(scan.log.some((l) => /Exclusão em lote por acesso local concluída: 1 excluído\(s\), 1 mantido\(s\)/.test(l.message)));
    // Já excluídos não entram de novo (e a prévia diz quantos ficaram de fora por isso).
    preview = (await api('POST', `/api/scans/${scan.id}/bulk-delete/preview`, { all: true, filters: {} })).data;
    assert.deepEqual([preview.total, preview.gone], [2, 2]);
    assert.equal((await api('GET', `/api/scans/${scan.id}/results`)).data.bulkCandidates, 2, 'os que ainda podem ser excluídos no recorte');
    // Pedidos inválidos: filtro desconhecido (ampliaria o alvo), nenhum arquivo, mais do que a prévia mostrou.
    assert.match((await api('POST', `/api/scans/${scan.id}/bulk-delete/preview`, { all: true, filters: { tipo: 'video' } })).data.error, /Filtro inválido: tipo/);
    assert.match((await api('POST', `/api/scans/${scan.id}/bulk-delete/preview`, { ids: [] })).data.error, /Escolha ao menos um arquivo/);
    res = await bulk(scan.id, { all: true, filters: {} }, { expected: 0 });
    assert.deepEqual([res.status, res.data.code], [409, 'preview-changed']);

    // Dois pedidos ao mesmo tempo (duas abas): só um lote começa.
    const twice = fs.mkdtempSync(path.join(root, 'repo-'));
    for (let i = 1; i <= 3; i++) fs.writeFileSync(path.join(twice, `clipe-${i}.mp4`), MP4);
    const twiceRepo = (await api('POST', '/api/repositories', { name: 'Clipes', path: twice, allowDelete: true })).data;
    res = await api('POST', '/api/scans', { repositoryIds: [twiceRepo.id], fileTypes: { categories: ['video'] }, options: { resolveOwner: false } });
    const twiceScan = await waitScan(res.data.id);
    const twicePreview = (await api('POST', `/api/scans/${twiceScan.id}/bulk-delete/preview`, { all: true, filters: {} })).data;
    const both = await Promise.all([1, 2].map(() => api('POST', `/api/scans/${twiceScan.id}/bulk-delete`, { all: true, filters: {}, confirmDelete: 'EXCLUIR', token: twicePreview.token })));
    assert.deepEqual(both.map((r) => r.status).sort(), [202, 409]);
    job = await waitBulk(twiceScan.id);
    assert.equal(job.deleted, 3);

    // Falha ao gravar o registro de exclusões (ex.: disco cheio): o lote para no primeiro arquivo.
    const noLog = fs.mkdtempSync(path.join(root, 'repo-'));
    for (let i = 1; i <= 3; i++) fs.writeFileSync(path.join(noLog, `filme-${i}.mp4`), MP4);
    const noLogRepo = (await api('POST', '/api/repositories', { name: 'Filmes', path: noLog, allowDelete: true })).data;
    res = await api('POST', '/api/scans', { repositoryIds: [noLogRepo.id], fileTypes: { categories: ['video'] }, options: { resolveOwner: false } });
    const noLogScan = await waitScan(res.data.id);
    const append = store.appendDeletions;
    store.appendDeletions = async () => {
      throw new Error('disco cheio');
    };
    try {
      res = await bulk(noLogScan.id, { all: true, filters: {} });
      assert.equal(res.status, 202);
      job = await waitBulk(noLogScan.id);
    } finally {
      store.appendDeletions = append;
    }
    assert.deepEqual([job.done, job.deleted], [1, 1]);
    assert.match(job.halted, /falha ao gravar o registro de exclusões/);
    assert.equal(fs.readdirSync(noLog).length, 2, 'os demais não foram excluídos');
    assert.ok((await api('GET', `/api/scans/${noLogScan.id}`)).data.log.some((l) => l.level === 'error' && /falha ao gravar o registro da exclusão de .*filme-\d\.mp4 \(deleted\): disco cheio/.test(l.message)));

    // Um lote que não terminou porque o CLEAN foi encerrado fica anotado ao iniciar de novo.
    store.updateScan(noLogScan.id, { bulkDeletion: { by: 'acesso local', startedAt: new Date().toISOString(), total: 3 } });
    createApp({ store, manager, scheduler, config: { authUser: '', authPassword: '' } });
    assert.equal(store.getScan(noLogScan.id).bulkDeletion, null);
    assert.ok(store.getScan(noLogScan.id).log.some((l) => l.level === 'warn' && /Exclusão em lote por acesso local \(3 arquivo\(s\), iniciada em .*\) interrompida: o CLEAN foi encerrado antes do fim/.test(l.message)));

    // Falhas seguidas numa pasta (aqui, cada arquivo virou uma pasta): depois de 5, os demais arquivos
    // dela não são tentados neste lote; os que falharam ficam por último na próxima tentativa.
    const many = fs.mkdtempSync(path.join(root, 'repo-'));
    fs.mkdirSync(path.join(many, 'A'));
    for (let i = 1; i <= 8; i++) fs.writeFileSync(path.join(many, 'A', `video-${String(i).padStart(2, '0')}.mp4`), MP4);
    const manyRepo = (await api('POST', '/api/repositories', { name: 'Vídeos', path: many, allowDelete: true })).data;
    res = await api('POST', '/api/scans', { repositoryIds: [manyRepo.id], fileTypes: { categories: ['video'] }, options: { resolveOwner: false } });
    const manyScan = await waitScan(res.data.id);
    assert.equal(manyScan.stats.filesMatched, 8);
    const byName = Object.fromEntries((await api('GET', `/api/scans/${manyScan.id}/results?sort=path&pageSize=50`)).data.items.map((r) => [r.name, r.id]));
    for (let i = 1; i <= 5; i++) {
      const file = path.join(many, 'A', `video-${String(i).padStart(2, '0')}.mp4`);
      fs.rmSync(file);
      fs.mkdirSync(file);
    }
    // Os que falham primeiro na ordem do pedido.
    const order = ['video-01.mp4', 'video-02.mp4', 'video-03.mp4', 'video-04.mp4', 'video-05.mp4', 'video-06.mp4', 'video-07.mp4', 'video-08.mp4'].map((n) => byName[n]);
    res = await bulk(manyScan.id, { ids: order });
    assert.equal(res.status, 202, JSON.stringify(res.data));
    job = await waitBulk(manyScan.id);
    assert.deepEqual([job.total, job.failed, job.notTried, job.deleted, job.halted], [8, 5, 3, 0, null]);
    assert.equal(job.methods, undefined, 'detalhes internos do lote não saem na API');
    let manyLog = (await api('GET', `/api/scans/${manyScan.id}`)).data.log;
    assert.ok(manyLog.some((l) => l.level === 'warn' && /concluída: 0 excluído\(s\), 5 falha\(s\), 3 não tentado\(s\) em pastas com 5 falhas seguidas/.test(l.message)));
    // Repetir o lote avança: os que falharam antes ficam por último.
    res = await bulk(manyScan.id, { ids: order });
    job = await waitBulk(manyScan.id);
    assert.deepEqual([job.deleted, job.failed], [3, 5]);
    assert.deepEqual(fs.readdirSync(path.join(many, 'A')).sort(), ['video-01.mp4', 'video-02.mp4', 'video-03.mp4', 'video-04.mp4', 'video-05.mp4']);

    // Falhas seguidas em várias pastas (ex.: a conta do CLEAN sem permissão): o lote para em 20.
    const spread = fs.mkdtempSync(path.join(root, 'repo-'));
    for (let i = 1; i <= 25; i++) {
      fs.mkdirSync(path.join(spread, `P${String(i).padStart(2, '0')}`));
      fs.writeFileSync(path.join(spread, `P${String(i).padStart(2, '0')}`, 'clipe.mp4'), MP4);
    }
    const spreadRepo = (await api('POST', '/api/repositories', { name: 'Espalhados', path: spread, allowDelete: true })).data;
    res = await api('POST', '/api/scans', { repositoryIds: [spreadRepo.id], fileTypes: { categories: ['video'] }, options: { resolveOwner: false } });
    const spreadScan = await waitScan(res.data.id);
    for (let i = 1; i <= 25; i++) {
      const file = path.join(spread, `P${String(i).padStart(2, '0')}`, 'clipe.mp4');
      fs.rmSync(file);
      fs.mkdirSync(file);
    }
    res = await bulk(spreadScan.id, { all: true, filters: {} });
    job = await waitBulk(spreadScan.id);
    assert.deepEqual([job.total, job.done, job.failed], [25, 20, 20]);
    assert.match(job.halted, /^20 falhas seguidas em várias pastas \(a última: O caminho não é um arquivo\.\)/);
    manyLog = (await api('GET', `/api/scans/${spreadScan.id}`)).data.log;
    assert.ok(manyLog.some((l) => l.level === 'warn' && /Exclusão em lote .* interrompida depois de 20 de 25 por 20 falhas seguidas/.test(l.message)));

    // A confirmação vale para os arquivos da prévia: outros (ou outra prévia) pedem para começar de novo.
    const tokenDir = fs.mkdtempSync(path.join(root, 'repo-'));
    for (let i = 1; i <= 3; i++) fs.writeFileSync(path.join(tokenDir, `aula-${i}.mp4`), MP4);
    const tokenRepo = (await api('POST', '/api/repositories', { name: 'Aulas', path: tokenDir, allowDelete: true })).data;
    res = await api('POST', '/api/scans', { repositoryIds: [tokenRepo.id], fileTypes: { categories: ['video'] }, options: { resolveOwner: false } });
    const tokenScan = await waitScan(res.data.id);
    const firstId = (await api('GET', `/api/scans/${tokenScan.id}/results`)).data.items[0].id;
    preview = (await api('POST', `/api/scans/${tokenScan.id}/bulk-delete/preview`, { ids: [firstId] })).data;
    assert.ok(preview.token);
    res = await api('POST', `/api/scans/${tokenScan.id}/bulk-delete`, { all: true, filters: {}, confirmDelete: 'EXCLUIR', token: preview.token });
    assert.deepEqual([res.status, res.data.code], [409, 'preview-changed']);
    assert.match(res.data.error, /2 arquivo\(s\) passaram a poder ser excluídos depois da prévia/);
    res = await api('POST', `/api/scans/${tokenScan.id}/bulk-delete`, { ids: [firstId], confirmDelete: 'EXCLUIR', token: 'outra' });
    assert.deepEqual([res.status, res.data.code], [409, 'preview-changed']);
    res = await api('POST', `/api/scans/${tokenScan.id}/bulk-delete`, { ids: [firstId], confirmDelete: 'EXCLUIR', token: preview.token });
    assert.equal(res.status, 202, JSON.stringify(res.data));
    job = await waitBulk(tokenScan.id);
    assert.equal(job.deleted, 1);

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
