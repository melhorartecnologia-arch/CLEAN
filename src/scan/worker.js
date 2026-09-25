// Ponto de entrada da thread de análise (worker_threads). Mantém o servidor web responsivo e
// permite interromper uma análise a qualquer momento.
import { parentPort, workerData } from 'node:worker_threads';
import { Scanner } from './scanner.js';

const scanner = new Scanner(workerData, (message) => parentPort.postMessage(message));

parentPort.on('message', (message) => {
  if (message?.type === 'cancel') scanner.cancel();
});

scanner
  .run()
  .catch((err) => parentPort.postMessage({ type: 'fatal', message: err?.stack || String(err) }))
  .finally(() => parentPort.unref());
