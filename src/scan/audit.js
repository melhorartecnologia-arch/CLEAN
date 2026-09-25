// Log de auditoria de segurança do Windows: identifica quem acessou/alterou cada arquivo por último.
//
// Requer auditoria habilitada no servidor de arquivos:
//   - 4663 "Tentativa de acesso a um objeto" (Auditoria do Sistema de Arquivos + SACL na pasta);
//   - 5145 "Verificação de acesso a objeto de compartilhamento de rede" (Auditoria Detalhada de
//     Compartilhamento de Arquivos).
// A conta que executa o CLEAN precisa ler o log de Segurança (Administradores ou "Leitores de Log
// de Eventos") do computador consultado.
import path from 'node:path';
import { runPowerShell, parseJsonLines } from './powershell.js';

// Percorre os eventos do mais recente para o mais antigo e emite, por arquivo, o último acesso
// qualquer (a=1) e a última alteração (w=1). Usa Properties por posição (mais rápido que ToXml()).
export const AUDIT_SCRIPT = `
$ErrorActionPreference = 'Stop'
$filter = @{ LogName = 'Security'; Id = @(4663, 5145); StartTime = (Get-Date).AddDays(-[int]$env:CLEAN_DAYS) }
$params = @{ FilterHashtable = $filter; MaxEvents = [int]$env:CLEAN_MAX; ErrorAction = 'Stop' }
if ($env:CLEAN_COMPUTER) { $params['ComputerName'] = $env:CLEAN_COMPUTER }
$seenAny = @{}
$seenWrite = @{}
$writeBits = 0xD0116
try {
  $out = Get-WinEvent @params | ForEach-Object {
    $e = $_
    $v = $e.Properties
    $share = ''
    $shareLocal = ''
    if ($e.Id -eq 4663) {
      if ($v.Count -lt 10 -or [string]$v[5].Value -ne 'File') { return }
      $file = [string]$v[6].Value
      $mask = [string]$v[9].Value
    } else {
      if ($v.Count -lt 11) { return }
      $share = [string]$v[7].Value
      $shareLocal = ([string]$v[8].Value) -replace '^\\\\\\?\\?\\\\', ''
      $rel = ([string]$v[9].Value).Trim('\\')
      if (-not $rel) { return }
      $file = $shareLocal.TrimEnd('\\') + '\\' + $rel
      $mask = [string]$v[10].Value
    }
    $user = [string]$v[1].Value
    if (-not $file -or -not $user -or $user -eq '-' -or $user.EndsWith('$')) { return }
    $key = $file.ToLowerInvariant()
    $isWrite = $false
    try { $isWrite = ([Convert]::ToInt64($mask, 16) -band $writeBits) -ne 0 } catch { }
    $any = -not $seenAny.ContainsKey($key)
    $write = $isWrite -and -not $seenWrite.ContainsKey($key)
    if (-not ($any -or $write)) { return }
    if ($any) { $seenAny[$key] = $true }
    if ($write) { $seenWrite[$key] = $true }
    ConvertTo-Json -Compress -InputObject @{
      t = $e.TimeCreated.ToUniversalTime().ToString('o'); id = $e.Id; u = $user; d = [string]$v[2].Value
      f = $file; s = $share; sl = $shareLocal; m = $mask; a = [int]$any; w = [int]$write
    }
  }
} catch {
  if ($_.FullyQualifiedErrorId -like 'NoMatchingEventsFound*') { $out = @() } else { throw }
}
if ($out) { Set-Content -LiteralPath $env:CLEAN_OUT -Encoding UTF8 -Value $out }
`;

const ACCESS = [
  [0x10000, 'Exclusão'],
  [0x2 | 0x4, 'Gravação'],
  [0x40000 | 0x80000, 'Alteração de permissões'],
  [0x10 | 0x100, 'Alteração de atributos'],
  [0x1, 'Leitura'],
];

/** Descreve a máscara de acesso (ex.: "0x2" -> "Gravação"). */
export function describeAccess(mask) {
  const value = Number.parseInt(String(mask || '0'), 16);
  if (!Number.isFinite(value)) return 'Acesso';
  for (const [bits, label] of ACCESS) if (value & bits) return label;
  return 'Acesso';
}

/** Normaliza caminhos do Windows para comparação: minúsculas, barras invertidas, sem prefixos. */
export function normalizeWinPath(p) {
  return String(p || '')
    .replace(/\//g, '\\')
    .replace(/^\\\\\?\\UNC\\/i, '\\\\')
    .replace(/^\\(\?\?|\\\?)\\/, '')
    .replace(/\\+$/, '')
    .toLowerCase();
}

/** Consulta o log de Segurança. Retorna a lista de eventos já reduzida (um por arquivo e tipo). */
export async function queryAuditEvents({ computer = '', days = 30, maxEvents = 200000 } = {}) {
  if (process.platform !== 'win32') {
    throw new Error('A consulta ao log de auditoria só está disponível quando o CLEAN roda no Windows.');
  }
  const lines = await runPowerShell(AUDIT_SCRIPT, {
    env: { CLEAN_COMPUTER: computer, CLEAN_DAYS: String(days), CLEAN_MAX: String(maxEvents) },
    timeoutMs: 30 * 60 * 1000,
  });
  return parseJsonLines(lines);
}

/** Índice de eventos por caminho local no servidor. */
export class AuditIndex {
  constructor(events = []) {
    this.any = new Map();
    this.write = new Map();
    this.shares = new Map(); // nome do compartilhamento (minúsculo) -> caminho local no servidor
    this.count = events.length;
    for (const e of events) {
      const key = normalizeWinPath(e.f);
      const info = {
        user: e.d && e.d !== '-' ? `${e.d}\\${e.u}` : e.u,
        time: e.t,
        eventId: e.id,
        action: describeAccess(e.m),
      };
      if (e.a && !this.any.has(key)) this.any.set(key, info);
      if (e.w && !this.write.has(key)) this.write.set(key, info);
      if (e.s && e.sl) {
        const share = String(e.s).replace(/^\\\\[^\\]*\\/, '').toLowerCase();
        if (share && !this.shares.has(share)) this.shares.set(share, e.sl.replace(/\\+$/, ''));
      }
    }
  }

  /**
   * Caminho local (no servidor) correspondente a um arquivo do repositório.
   * repoPath: raiz configurada (ex.: \\servidor\Financeiro ou D:\Dados); localRoot: caminho local
   * informado manualmente (opcional); relativePath: caminho do arquivo relativo à raiz.
   */
  localPathFor(repoPath, relativePath, localRoot = '') {
    const rel = String(relativePath || '').replace(/\//g, '\\');
    if (localRoot) return path.win32.join(localRoot, rel);
    const root = String(repoPath || '').replace(/\//g, '\\');
    const unc = /^\\\\([^\\]+)\\([^\\]+)(?:\\(.*))?$/.exec(root);
    if (!unc) return path.win32.join(root, rel);
    const [, , share, sub = ''] = unc;
    let base = null;
    if (/^[a-z]\$$/i.test(share)) base = `${share[0]}:\\`; // compartilhamento administrativo (D$)
    else base = this.shares.get(share.toLowerCase()) || null;
    return base ? path.win32.join(base, sub, rel) : null;
  }

  lookup(repoPath, relativePath, localRoot = '') {
    const local = this.localPathFor(repoPath, relativePath, localRoot);
    if (!local) return null;
    const key = normalizeWinPath(local);
    const any = this.any.get(key);
    if (!any) return null;
    const write = this.write.get(key);
    return { ...any, lastWrite: write && write !== any ? write : null };
  }
}

/**
 * Escolhe o "último usuário" entre as fontes disponíveis, da mais precisa para a menos precisa:
 * log de auditoria > metadados do documento ("salvo por último por") > proprietário do arquivo.
 */
export function pickLastUser({ audit, metadata, owner }) {
  if (audit?.user) return { user: audit.user, source: 'audit' };
  if (metadata?.lastModifiedBy) return { user: metadata.lastModifiedBy, source: 'metadata' };
  if (owner) return { user: owner, source: 'owner' };
  return { user: null, source: null };
}
