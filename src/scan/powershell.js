// Execução de scripts PowerShell (Windows PowerShell 5.1 ou PowerShell 7).
// O script é enviado com -EncodedCommand, que não depende da política de execução de scripts, e
// troca dados com o Node por arquivos temporários UTF-8 (evita problemas de página de código do console).
import { spawn } from 'node:child_process';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import crypto from 'node:crypto';

export function powershellPath() {
  return process.env.POWERSHELL_PATH || (process.platform === 'win32' ? 'powershell.exe' : 'pwsh');
}

export function encodeCommand(script) {
  return Buffer.from(script, 'utf16le').toString('base64');
}

/**
 * Executa o script passando `input` (linhas) no arquivo indicado por $env:CLEAN_IN e devolve as
 * linhas gravadas pelo script em $env:CLEAN_OUT. Variáveis extras podem ser passadas em `env`.
 */
export async function runPowerShell(script, { input = [], env = {}, timeoutMs = 10 * 60 * 1000, signal } = {}) {
  if (signal?.aborted) throw new Error('Operação cancelada.');
  const id = crypto.randomUUID();
  const inFile = path.join(os.tmpdir(), `clean-${id}-in.txt`);
  const outFile = path.join(os.tmpdir(), `clean-${id}-out.txt`);
  await fs.writeFile(inFile, input.join('\n'), 'utf8');
  try {
    await new Promise((resolve, reject) => {
      const child = spawn(powershellPath(), ['-NoProfile', '-NonInteractive', '-EncodedCommand', encodeCommand(script)], {
        env: { ...process.env, ...env, CLEAN_IN: inFile, CLEAN_OUT: outFile },
        windowsHide: true,
        stdio: ['ignore', 'ignore', 'pipe'],
      });
      let stderr = '';
      child.stderr.on('data', (chunk) => {
        if (stderr.length < 8000) stderr += chunk;
      });
      const timer = setTimeout(() => child.kill(), timeoutMs);
      const onAbort = () => child.kill();
      signal?.addEventListener('abort', onAbort, { once: true });
      child.on('error', (err) => {
        clearTimeout(timer);
        signal?.removeEventListener('abort', onAbort);
        reject(new Error(`Não foi possível executar o PowerShell (${powershellPath()}): ${err.message}`));
      });
      child.on('close', (code) => {
        clearTimeout(timer);
        signal?.removeEventListener('abort', onAbort);
        if (signal?.aborted) reject(new Error('Operação cancelada.'));
        else if (code === 0) resolve();
        else reject(new Error(`PowerShell terminou com código ${code}: ${cleanStderr(stderr)}`));
      });
    });
    const raw = await fs.readFile(outFile, 'utf8').catch(() => '');
    return raw
      .replace(/^\uFEFF/, '')
      .split(/\r?\n/)
      .filter((line) => line.trim());
  } finally {
    await Promise.all([fs.rm(inFile, { force: true }), fs.rm(outFile, { force: true })]);
  }
}

/**
 * Remove arquivos temporários de execuções anteriores que não terminaram normalmente (ex.: servidor
 * encerrado no meio de uma análise). Eles contêm caminhos de arquivos e nomes de usuários.
 */
export async function cleanupTempFiles(maxAgeMs = 60 * 60 * 1000) {
  const dir = os.tmpdir();
  const names = await fs.readdir(dir).catch(() => []);
  const now = Date.now();
  await Promise.all(
    names
      .filter((name) => /^clean-[0-9a-f-]{36}-(in|out)\.txt$/.test(name))
      .map(async (name) => {
        const file = path.join(dir, name);
        const st = await fs.stat(file).catch(() => null);
        if (st && now - st.mtimeMs > maxAgeMs) await fs.rm(file, { force: true }).catch(() => {});
      }),
  );
}

function cleanStderr(text) {
  // Mensagens de erro do PowerShell chegam em formato CLIXML quando a saída é redirecionada.
  const plain = text.replace(/#< CLIXML/g, '').replace(/<[^>]+>/g, ' ').replace(/_x000D__x000A_/g, ' ');
  return plain.replace(/\s+/g, ' ').trim().slice(0, 500) || 'sem detalhes';
}

/** Converte linhas JSON em objetos, ignorando linhas inválidas. */
export function parseJsonLines(lines) {
  const out = [];
  for (const line of lines) {
    try {
      out.push(JSON.parse(line));
    } catch {
      // linha inválida
    }
  }
  return out;
}
