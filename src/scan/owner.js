// Proprietário dos arquivos: no Windows usa Get-Acl (em lotes, via PowerShell); em outros
// sistemas usa o UID do arquivo (útil para testes e para compartilhamentos montados via Samba).
import fs from 'node:fs/promises';
import { runPowerShell, parseJsonLines } from './powershell.js';

// Compatível com Windows PowerShell 5.1 e com o modo de linguagem restrita (sem tipos .NET).
export const OWNER_SCRIPT = `
$ErrorActionPreference = 'Stop'
$paths = @(Get-Content -LiteralPath $env:CLEAN_IN -Encoding UTF8)
$out = foreach ($p in $paths) {
  if (-not $p) { continue }
  $owner = ''
  $err = ''
  try {
    $owner = (Get-Acl -LiteralPath $p).Owner
  } catch {
    $err = $_.Exception.Message
    if ($p.Length -ge 248 -and -not $p.StartsWith('\\\\?\\')) {
      if ($p.StartsWith('\\\\')) { $long = '\\\\?\\UNC\\' + $p.Substring(2) } else { $long = '\\\\?\\' + $p }
      try {
        $owner = (Get-Acl -LiteralPath $long).Owner
        $err = ''
      } catch { }
    }
  }
  ConvertTo-Json -Compress -InputObject @{ p = $p; o = [string]$owner; e = [string]$err }
}
if ($out) { Set-Content -LiteralPath $env:CLEAN_OUT -Encoding UTF8 -Value $out }
`;

let passwdCache = null;
async function posixUserName(uid) {
  if (!passwdCache) {
    passwdCache = new Map();
    const text = await fs.readFile('/etc/passwd', 'utf8').catch(() => '');
    for (const line of text.split('\n')) {
      const [name, , id] = line.split(':');
      if (name && id !== undefined) passwdCache.set(Number(id), name);
    }
  }
  return passwdCache.get(uid) || `uid:${uid}`;
}

export class OwnerResolver {
  constructor({ platform = process.platform, batchSize = 400 } = {}) {
    this.platform = platform;
    this.batchSize = batchSize;
    this.failure = null; // erro que impediu o uso do PowerShell (informado uma única vez)
  }

  /** Resolve os proprietários. Retorna Map(caminho -> { owner, error }). */
  async resolve(paths) {
    const result = new Map();
    if (paths.length === 0) return result;
    if (this.platform !== 'win32') {
      for (const p of paths) {
        try {
          const st = await fs.stat(p);
          result.set(p, { owner: await posixUserName(st.uid), error: null });
        } catch (err) {
          result.set(p, { owner: null, error: err.message });
        }
      }
      return result;
    }
    if (this.failure) {
      for (const p of paths) result.set(p, { owner: null, error: this.failure });
      return result;
    }
    for (let i = 0; i < paths.length; i += this.batchSize) {
      const batch = paths.slice(i, i + this.batchSize);
      try {
        const rows = parseJsonLines(await runPowerShell(OWNER_SCRIPT, { input: batch, timeoutMs: 5 * 60 * 1000 }));
        for (const row of rows) result.set(row.p, { owner: row.o || null, error: row.e || null });
      } catch (err) {
        this.failure = err.message;
        for (const p of batch) result.set(p, { owner: null, error: err.message });
      }
    }
    for (const p of paths) if (!result.has(p)) result.set(p, { owner: null, error: 'Sem resposta do PowerShell.' });
    return result;
  }
}
