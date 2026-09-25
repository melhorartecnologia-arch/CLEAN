// Ponto de entrada da thread de análise (worker_threads). Mantém o servidor web responsivo e
// permite interromper uma análise a qualquer momento.
import { parentPort, workerData } from 'node:worker_threads';
import { Scanner } from './scanner.js';
import { MailScanner } from '../mail/scanner.js';

const emit = (message) => parentPort.postMessage(message);
const scanner = workerData.kind === 'mail' ? new MailScanner(workerData, emit) : new Scanner(workerData, emit);

// Bibliotecas de leitura (ex.: pdf.js com um PDF danificado) podem gerar erros fora da promessa
// aguardada. Sem estes tratadores a thread inteira seria encerrada e a análise falharia; com eles o
// erro vira um aviso e o arquivo em questão é encerrado pelo tempo limite.
const seen = new Set();
function report(kind, err) {
  const message = `${err?.name || 'Erro'}: ${err?.message || err}`;
  if (seen.has(message) || seen.size > 50) return;
  seen.add(message);
  const items = scanner.inFlight.size ? ` Em leitura: ${[...scanner.inFlight].join(' | ')}` : '';
  scanner.log('warn', `Erro interno ignorado (${kind}): ${message}.${items}`);
}
process.on('unhandledRejection', (reason) => report('promessa', reason));
process.on('uncaughtException', (err) => report('exceção', err));

parentPort.on('message', (message) => {
  if (message?.type === 'cancel') scanner.cancel();
});

scanner
  .run()
  .catch((err) => parentPort.postMessage({ type: 'fatal', message: err?.stack || String(err) }))
  .finally(() => parentPort.unref());
