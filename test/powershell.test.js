// Testa os scripts PowerShell com versões simuladas de Get-Acl e Get-WinEvent.
// Roda no Windows (powershell.exe) ou onde houver PowerShell 7 (pwsh); caso contrário é ignorado.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { OWNER_SCRIPT } from '../src/scan/owner.js';
import { AUDIT_SCRIPT, AuditIndex, describeAccess, normalizeWinPath, pickLastUser } from '../src/scan/audit.js';
import { runPowerShell, parseJsonLines, powershellPath } from '../src/scan/powershell.js';

const hasPowerShell = spawnSync(powershellPath(), ['-NoProfile', '-Command', 'exit 0'], { windowsHide: true }).status === 0;
const skip = hasPowerShell ? false : 'PowerShell não encontrado';

test('scripts PowerShell não têm erros de sintaxe', { skip }, () => {
  for (const script of [OWNER_SCRIPT, AUDIT_SCRIPT]) {
    const check = `$errors = $null; [void][System.Management.Automation.Language.Parser]::ParseInput($env:SCRIPT, [ref]$null, [ref]$errors); if ($errors.Count) { $errors | ForEach-Object { $_.Message }; exit 1 }`;
    const r = spawnSync(powershellPath(), ['-NoProfile', '-NonInteractive', '-Command', check], {
      env: { ...process.env, SCRIPT: script },
      encoding: 'utf8',
      windowsHide: true,
    });
    assert.equal(r.status, 0, r.stdout + r.stderr);
  }
});

const MOCK_ACL = `
function Get-Acl {
  param([string]$LiteralPath)
  if ($LiteralPath -like '*inexistente*') { throw "Caminho não encontrado: $LiteralPath" }
  if ($LiteralPath.Length -ge 248 -and -not $LiteralPath.StartsWith('\\\\?\\')) { throw 'Caminho muito longo' }
  if ($LiteralPath.StartsWith('\\\\?\\')) { return [pscustomobject]@{ Owner = 'EMPRESA\\longo' } }
  [pscustomobject]@{ Owner = 'EMPRESA\\joão.silva' }
}
`;

test('proprietários: acentos, colchetes, erros e caminhos longos', { skip }, async () => {
  const long = `D:\\Dados\\${'pasta\\'.repeat(45)}arquivo.docx`;
  const paths = ['D:\\Dados\\relatório [final] ação.xlsx', 'D:\\Dados\\inexistente.txt', long];
  const rows = parseJsonLines(await runPowerShell(MOCK_ACL + OWNER_SCRIPT, { input: paths }));
  assert.equal(rows.length, 3);
  assert.deepEqual(rows[0], { p: paths[0], o: 'EMPRESA\\joão.silva', e: '' });
  assert.equal(rows[1].o, '');
  assert.match(rows[1].e, /não encontrado/);
  assert.equal(rows[2].o, 'EMPRESA\\longo');
  assert.equal(rows[2].e, '');
});

const MOCK_EVENTS = `
function Get-WinEvent {
  [CmdletBinding()]
  param($FilterHashtable, $MaxEvents, $ComputerName)
  if ($env:CLEAN_COMPUTER -eq 'vazio') {
    $err = New-Object System.Management.Automation.ErrorRecord (New-Object System.Exception 'No events were found'), 'NoMatchingEventsFound', 'ObjectNotFound', $null
    $PSCmdlet.ThrowTerminatingError($err)
  }
  function ev($id, $t, [object[]]$vals) {
    [pscustomobject]@{ Id = $id; TimeCreated = [datetime]::Parse($t).ToUniversalTime(); Properties = @($vals | ForEach-Object { [pscustomobject]@{ Value = $_ } }) }
  }
  # Mais recentes primeiro, como o Get-WinEvent real
  ev 5145 '2026-09-20T15:00:00Z' @('S-1', 'ana', 'EMPRESA', '0x1', 'File', '10.0.0.5', '5000', '\\\\*\\Financeiro', '\\??\\E:\\Shares\\Financeiro', 'RH\\salarios.xlsx', '0x120089', '%%4416', '')
  ev 4663 '2026-09-20T14:00:00Z' @('S-2', 'bruno', 'EMPRESA', '0x2', 'Security', 'File', 'E:\\Shares\\Financeiro\\RH\\salarios.xlsx', '0x10', '%%4417', '0x2', '0x4', 'EXCEL.EXE', '')
  ev 4663 '2026-09-20T13:00:00Z' @('S-3', 'SERVIDOR$', 'EMPRESA', '0x3', 'Security', 'File', 'E:\\Shares\\Financeiro\\RH\\outro.docx', '0x10', '%%4417', '0x2', '0x4', 'x.exe', '')
  ev 4663 '2026-09-20T12:00:00Z' @('S-4', 'carla', 'EMPRESA', '0x4', 'Security', 'Key', '\\REGISTRY\\MACHINE\\X', '0x10', '%%4417', '0x2', '0x4', 'x.exe', '')
  ev 4663 '2026-09-19T12:00:00Z' @('S-5', 'daniel', 'EMPRESA', '0x5', 'Security', 'File', 'E:\\Shares\\Financeiro\\RH\\salarios.xlsx', '0x10', '%%1537', '0x10000', '0x4', 'explorer.exe', '')
  ev 4663 '2026-09-18T12:00:00Z' @('S-6', 'eva', 'EMPRESA', '0x6', 'Security', 'File', 'D:\\Local\\planilha.xlsx', '0x10', '%%4417', '0x2', '0x4', 'EXCEL.EXE', '')
}
`;

test('auditoria: último acesso, última alteração e mapeamento de compartilhamentos', { skip }, async () => {
  const events = parseJsonLines(
    await runPowerShell(MOCK_EVENTS + AUDIT_SCRIPT, { env: { CLEAN_COMPUTER: 'srv', CLEAN_DAYS: '30', CLEAN_MAX: '1000' } }),
  );
  // salarios.xlsx: acesso (ana, leitura) + alteração (bruno); conta de máquina e registro ignorados
  assert.equal(events.length, 3);
  const index = new AuditIndex(events);
  const hit = index.lookup('\\\\srv\\Financeiro', 'RH\\salarios.xlsx');
  assert.equal(hit.user, 'EMPRESA\\ana');
  assert.equal(hit.action, 'Leitura');
  assert.equal(hit.eventId, 5145);
  assert.equal(hit.lastWrite.user, 'EMPRESA\\bruno');
  assert.equal(hit.lastWrite.action, 'Gravação');
  // Caminho local e compartilhamento administrativo
  assert.equal(index.lookup('D:\\Local', 'planilha.xlsx').user, 'EMPRESA\\eva');
  assert.equal(index.lookup('\\\\srv\\d$\\Local', 'planilha.xlsx').user, 'EMPRESA\\eva');
  // Caminho local informado manualmente
  assert.equal(index.lookup('\\\\outro\\Fin', 'RH\\salarios.xlsx', 'E:\\Shares\\Financeiro').user, 'EMPRESA\\ana');
  assert.equal(index.lookup('\\\\srv\\Desconhecido', 'a.txt'), null);
});

test('auditoria: nenhum evento encontrado não é erro', { skip }, async () => {
  const lines = await runPowerShell(MOCK_EVENTS + AUDIT_SCRIPT, { env: { CLEAN_COMPUTER: 'vazio', CLEAN_DAYS: '1', CLEAN_MAX: '10' } });
  assert.deepEqual(lines, []);
});

test('funções auxiliares de auditoria', () => {
  assert.equal(describeAccess('0x10000'), 'Exclusão');
  assert.equal(describeAccess('0x6'), 'Gravação');
  assert.equal(describeAccess('0x120089'), 'Leitura');
  assert.equal(describeAccess('0x40000'), 'Alteração de permissões');
  assert.equal(normalizeWinPath('\\??\\E:\\Dados\\'), 'e:\\dados');
  assert.equal(normalizeWinPath('\\\\?\\UNC\\srv\\Share\\A.txt'), '\\\\srv\\share\\a.txt');
  assert.equal(normalizeWinPath('\\\\?\\C:\\X'), 'c:\\x');
  assert.deepEqual(pickLastUser({ audit: { user: 'A' }, metadata: { lastModifiedBy: 'B' }, owner: 'C' }), { user: 'A', source: 'audit' });
  assert.deepEqual(pickLastUser({ metadata: { lastModifiedBy: 'B' }, owner: 'C' }), { user: 'B', source: 'metadata' });
  assert.deepEqual(pickLastUser({ owner: 'C' }), { user: 'C', source: 'owner' });
  assert.deepEqual(pickLastUser({}), { user: null, source: null });
});
