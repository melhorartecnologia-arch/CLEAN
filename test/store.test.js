import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { Store } from '../src/store.js';

test('leituras simultâneas dos resultados não duplicam nem perdem registros', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'clean-store-'));
  try {
    const store = await new Store(dir).init();
    const scan = store.createScan({ name: 'x', status: 'running' });
    const batch = (from, n) => Array.from({ length: n }, (_, i) => ({ id: from + i, name: `arquivo ${from + i}` }));
    await store.appendResults(scan.id, batch(1, 20));
    const [a, b] = await Promise.all([store.readResults(scan.id), store.readResults(scan.id)]);
    assert.equal(a.length, 20);
    assert.equal(b.length, 20);
    await store.appendResults(scan.id, batch(21, 10));
    const [c, d, e] = await Promise.all([store.readResults(scan.id), store.readResults(scan.id), store.readErrors(scan.id)]);
    assert.equal(c.length, 30);
    assert.equal(d.length, 30);
    assert.deepEqual(e, []);
    assert.deepEqual(
      c.map((r) => r.id),
      Array.from({ length: 30 }, (_, i) => i + 1),
    );
    await store.close();
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('dados persistem e análises interrompidas são marcadas ao reabrir', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'clean-store-'));
  try {
    const store = await new Store(dir).init();
    const repo = store.createRepository({ name: 'R', path: '/tmp' });
    const scan = store.createScan({ name: 'y', status: 'running' });
    store.appendLog(scan.id, { level: 'info', message: 'iniciada' });
    await store.close();
    const reopened = await new Store(dir).init();
    assert.equal(reopened.getRepository(repo.id).name, 'R');
    assert.equal(reopened.getScan(scan.id).status, 'interrupted');
    assert.equal(reopened.getScan(scan.id).log[0].message, 'iniciada');
    assert.throws(() => reopened.scanDir('../../etc'), /inválido/);
    await reopened.close();
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});
