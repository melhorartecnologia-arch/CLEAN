// Executa o teste de termos em uma thread separada: uma expressão regular mal escrita (retrocesso
// catastrófico) não trava o servidor, pois a thread é encerrada após o tempo limite.
import { parentPort, workerData } from 'node:worker_threads';
import { Matcher } from './matcher.js';

const matcher = new Matcher(workerData.terms, { maxSamples: 5 });
parentPort.postMessage(matcher.match([{ text: workerData.text }], 'content'));
