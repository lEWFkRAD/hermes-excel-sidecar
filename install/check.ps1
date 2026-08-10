[CmdletBinding()]
param(
  [string] $ProfileName,
  [string] $HermesHome,
  [string] $OwnerReceiptPath
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$failures = New-Object System.Collections.Generic.List[string]
Import-Module (Join-Path $PSScriptRoot 'profile-ownership.psm1') -Force
[void](Resolve-HermesExcelProfileContext -ProfileName $ProfileName -HermesHome $HermesHome -OwnerReceiptPath $OwnerReceiptPath)

Get-ChildItem -LiteralPath $PSScriptRoot -File | Where-Object { $_.Extension -in @('.ps1', '.psm1') } | ForEach-Object {
  $tokens = $null
  $errors = $null
  [void][System.Management.Automation.Language.Parser]::ParseFile($_.FullName, [ref]$tokens, [ref]$errors)
  foreach ($error in $errors) {
    $failures.Add("$($_.Name): $($error.Message)")
  }
}

$manifest = [xml](Get-Content -LiteralPath (Join-Path $root 'manifest.xml') -Raw)
if ($manifest.OfficeApp.Id -ne '4fd4d435-7f9a-4d6d-9251-32f154f83a1f') {
  $failures.Add('manifest.xml has an unexpected or missing add-in ID.')
}

foreach ($relative in @(
  'broker\server.mjs',
  'taskpane.html',
  'taskpane.css',
  'taskpane.js',
  'assets\icon-16.png',
  'assets\icon-32.png',
  'assets\icon-64.png',
  'assets\icon-80.png',
  'service\bridge-service.cmd',
  'install\run-bridge.cmd.template',
  'install\run-bridge.vbs.template'
)) {
  if (-not (Test-Path -LiteralPath (Join-Path $root $relative))) {
    $failures.Add("Required payload file is missing: $relative")
  }
}

$taskScript = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'register-task.ps1') -Raw
if ($taskScript -notmatch 'shell\.Run\s+"""\$ServiceCmd""",\s*0,\s*True') {
  $failures.Add('Scheduled Task shim must wait for the supervisor (bWaitOnReturn=True).')
}

$installer = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'addin-install.ps1') -Raw
if ($installer -notmatch "service\s+-eq\s+'hermes-excel-bridge'") {
  $failures.Add('Installer health check must validate the bridge service identity.')
}
if ($installer -notmatch 'Ensure-IngestToken' -or $installer -notmatch 'hermes gateway restart') {
  $failures.Add('Installer must provision the adapter ingest token and restart Hermes gateway.')
}
if ($installer -notmatch 'PreviousTaskWasDisabled' -or $installer -notmatch 'Disable-ScheduledTask') {
  $failures.Add('Installer must preserve an operator-disabled bridge task across upgrades.')
}
foreach ($ownershipContract in @(
  'Enter-HermesExcelTransaction',
  'Assert-HermesExcelOwnership',
  'Write-HermesExcelOwnerReceipt',
  'AdoptLegacy',
  'OwnerFingerprint'
)) {
  if ($installer -notmatch [regex]::Escape($ownershipContract)) {
    $failures.Add("Installer is missing ownership contract '$ownershipContract'.")
  }
}
$ownershipModule = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'profile-ownership.psm1') -Raw
foreach ($failClosedContract in @(
  'FileShare]::None',
  'requires exactly one verifiable scheduled-task action',
  'AdoptLegacy is permitted only for installation',
  'File]::Replace'
)) {
  if ($ownershipModule -notmatch [regex]::Escape($failClosedContract)) {
    $failures.Add("Profile ownership module is missing fail-closed contract '$failClosedContract'.")
  }
}
$launcherTemplate = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'run-bridge.cmd.template') -Raw
if ($launcherTemplate -notmatch 'HERMES_EXCEL_INGEST_TOKEN=<"%HERMES_EXCEL_DATA_DIR%\\\.ingest-token"' -or
    $launcherTemplate -notmatch 'HERMES_EXCEL_ALLOW_RAW_FALLBACK=0') {
  $failures.Add('Installed launcher must share the ingest token and disable raw fallback.')
}
foreach ($profileVariable in @('HERMES_HOME', 'HERMES_CONFIG', 'HERMES_EXCEL_PROFILE_NAME', 'HERMES_EXCEL_OWNER_FINGERPRINT')) {
  if ($launcherTemplate -notmatch [regex]::Escape($profileVariable)) {
    $failures.Add("Installed launcher must bind '$profileVariable' to the selected owner profile.")
  }
}

$deployScript = Join-Path $PSScriptRoot 'deploy.ps1'
if (-not (Test-Path -LiteralPath $deployScript)) {
  $failures.Add('Canonical deploy script is missing.')
} else {
  $deployText = Get-Content -LiteralPath $deployScript -Raw
  if ($deployText -notmatch 'git.+ls-files' -or $deployText -notmatch 'npm\.cmd.+run.+verify') {
    $failures.Add('Deploy must use the tracked-file manifest and run the verification gate.')
  }
}

$smoke = Get-Content -LiteralPath (Join-Path $root 'broker\smoke.mjs') -Raw
if ($smoke -notmatch 'https://localhost:8788') {
  $failures.Add('Smoke test must default to the Excel bridge on port 8788.')
}

$supervisor = Get-Content -LiteralPath (Join-Path $root 'service\bridge-service.cmd') -Raw
if ($supervisor -notmatch 'bridge-supervisor\.lock' -or $supervisor -notmatch 'SUPERVISOR_PID') {
  $failures.Add('Bridge supervisor must enforce a single live restart loop.')
}
if ($deployText -notmatch 'Stop-Process.+ProcessId' -or $deployText -notmatch 'bridge-supervisor\.lock') {
  $failures.Add('Deploy restart must retire orphan bridge processes and the old singleton lock.')
}

$pane = Get-Content -LiteralPath (Join-Path $root 'taskpane.html') -Raw
if ($pane -match '<input[^>]+id="reviewToggle"[^>]+checked') {
  $failures.Add('Review-before-apply must remain opt-in in the initial HTML.')
}

if ($failures.Count) {
  $failures | ForEach-Object { Write-Error $_ }
  exit 1
}

Write-Host "Excel add-in install/package checks passed."
