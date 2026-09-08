<#
.SYNOPSIS
  Reverse a Hermes for Excel install -- stop and unregister everything.

.DESCRIPTION
  requires elevation? NO. Removes only per-user artifacts (HKCU + %LOCALAPPDATA%).

  Reverses, in order:
    1. Stop + unregister the "Hermes_Excel_Bridge" scheduled task.
    2. Stop any running bridge launched from the install dir.
    3. Remove the Startup-folder shortcut.
    4. Remove the WEF Developer registry value + catalog manifest (sideload).
    5. Delete %LOCALAPPDATA%\hermes\excel-addin (prompts unless -Force).

  Idempotent: every step tolerates already-absent artifacts.

.PARAMETER OfficeVersion
  Office registry version segment for the sideload cleanup. Default 16.0.

.PARAMETER Force
  Delete the install dir without prompting.
#>
[CmdletBinding()]
param(
  [string] $OfficeVersion = '16.0',
  [switch] $Force,
  [string] $ProfileName,
  [string] $HermesHome,
  [string] $OwnerReceiptPath
)

$ErrorActionPreference = 'Stop'

# ---- Config block ----------------------------------------------------------
$ScriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Path
Import-Module (Join-Path $ScriptDir 'profile-ownership.psm1') -Force
$ProfileContext = Resolve-HermesExcelProfileContext -ProfileName $ProfileName -HermesHome $HermesHome -OwnerReceiptPath $OwnerReceiptPath
$InstallDir = $ProfileContext.InstallPath
$TaskName   = 'Hermes_Excel_Bridge'
$TransactionLock = $null
$TransactionProof = $null
# ----------------------------------------------------------------------------

function Step($msg) { Write-Host "==> $msg" -ForegroundColor Cyan }

function Get-OwnedBridgeProcesses {
  $targets = @{
    'wscript.exe' = @((Join-Path $InstallDir 'run-bridge.vbs'), (Join-Path $InstallDir 'service\bridge-service.vbs'))
    'cscript.exe' = @((Join-Path $InstallDir 'run-bridge.vbs'), (Join-Path $InstallDir 'service\bridge-service.vbs'))
    'cmd.exe' = @((Join-Path $InstallDir 'run-bridge.cmd'), (Join-Path $InstallDir 'service\bridge-service.cmd'))
    'node.exe' = @((Join-Path $InstallDir 'broker\server.mjs'))
  }
  $owned = New-Object System.Collections.Generic.List[object]
  foreach ($process in @(Get-CimInstance Win32_Process -ErrorAction Stop)) {
    $name = ([string]$process.Name).ToLowerInvariant()
    if (-not $targets.ContainsKey($name) -or [string]::IsNullOrWhiteSpace([string]$process.CommandLine)) { continue }
    foreach ($target in $targets[$name]) {
      if ($process.CommandLine -match ('(?i)(?:^|["\s])' + [regex]::Escape($target) + '(?=["\s]|$)')) {
        $owned.Add($process)
        break
      }
    }
  }
  return $owned.ToArray()
}

function Stop-OwnedBridgeProcesses {
  for ($attempt = 0; $attempt -lt 20; $attempt++) {
    $owned = @(Get-OwnedBridgeProcesses)
    if ($owned.Count -eq 0) { return }
    # Stop relaunching parents first, then their exact cmd/node children.
    $ordered = @($owned | Sort-Object @{ Expression = {
      switch (([string]$_.Name).ToLowerInvariant()) {
        'wscript.exe' { 0 }
        'cscript.exe' { 0 }
        'cmd.exe' { 1 }
        default { 2 }
      }
    } })
    foreach ($process in $ordered) {
      Write-Host "    stopping owned $($process.Name) PID $($process.ProcessId)"
      try {
        Stop-Process -Id $process.ProcessId -Force -ErrorAction Stop
      } catch {
        if (Get-Process -Id $process.ProcessId -ErrorAction SilentlyContinue) { throw }
      }
    }
    Start-Sleep -Milliseconds 100
  }
  $remaining = @(Get-OwnedBridgeProcesses)
  $labels = @($remaining | ForEach-Object { "$($_.Name):$($_.ProcessId)" }) -join ', '
  throw "Owned bridge processes survived termination: $labels."
}

function Assert-OwnedBridgeQuiescent([int] $OwnedPort) {
  for ($attempt = 0; $attempt -lt 50; $attempt++) {
    $remaining = @(Get-OwnedBridgeProcesses)
    $listeners = @([Net.NetworkInformation.IPGlobalProperties]::GetIPGlobalProperties().GetActiveTcpListeners() |
      Where-Object { $_.Port -eq $OwnedPort })
    if ($remaining.Count -eq 0 -and $listeners.Count -eq 0) { return }
    Start-Sleep -Milliseconds 100
  }
  $remaining = @(Get-OwnedBridgeProcesses)
  if ($remaining.Count -ne 0) {
    $labels = @($remaining | ForEach-Object { "$($_.Name):$($_.ProcessId)" }) -join ', '
    throw "Owned bridge processes remain after teardown: $labels."
  }
  throw "Bridge port $OwnedPort is still listening after owned process teardown. Refusing to retire ownership."
}

try {
  $TransactionProof = [guid]::NewGuid().ToString('N')
  $TransactionLock = Enter-HermesExcelTransaction -Context $ProfileContext -Proof $TransactionProof
  $ownerReceipt = Assert-HermesExcelOwnership -Context $ProfileContext -Operation Mutate
  if (-not $ownerReceipt) {
    Write-Host 'No owned Hermes for Excel singleton is installed; nothing to roll back.'
    exit 0
  }
  $OwnedBridgePort = [int]$ownerReceipt.BridgePort
  $env:HERMES_EXCEL_TRANSACTION_PROOF = $TransactionProof
  $env:HERMES_EXCEL_TRANSACTION_PROFILE = $ProfileContext.ProfileName
  $env:HERMES_EXCEL_TRANSACTION_HOME = $ProfileContext.ProfileHome
  $env:HERMES_EXCEL_TRANSACTION_RECEIPT = $ProfileContext.ReceiptPath
  $env:HERMES_HOME = $ProfileContext.ProfileHome
  if (Test-Path -LiteralPath $InstallDir) {
    $proceed = $Force
    if (-not $proceed) {
      $answer = Read-Host "Delete '$InstallDir' and ALL its data (uploads/exports/logs/token)? [y/N]"
      $proceed = $answer -match '^(y|yes)$'
    }
    if ($proceed) {
      Write-Verbose 'Rollback payload deletion confirmed before singleton mutation.'
    } else {
      throw 'Rollback canceled before any owned singleton mutation; payload and ownership receipt were retained.'
    }
  }
  Write-Host ''
  Write-Host "Hermes for Excel rollback" -ForegroundColor White
  Write-Host "Install: $InstallDir"
  Write-Host ''

  # 1. Scheduled task (prefer the installed register-task.ps1; fall back inline).
  Step "1. Unregister scheduled task '$TaskName'"
  $regTask = Join-Path $InstallDir 'install\register-task.ps1'
  if (-not (Test-Path -LiteralPath $regTask)) { $regTask = Join-Path $ScriptDir 'register-task.ps1' }
  if (Test-Path -LiteralPath $regTask) {
    & powershell -NoProfile -ExecutionPolicy Bypass -File $regTask -InstallDir $InstallDir -Unregister `
      -ProfileName $ProfileContext.ProfileName -HermesHome $ProfileContext.ProfileHome `
      -OwnerReceiptPath $ProfileContext.ReceiptPath -TransactionProof $TransactionProof
    if ($LASTEXITCODE -ne 0) { throw "Scheduled task unregistration failed with exit code $LASTEXITCODE; receipt retained." }
  } else {
    $t = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    if ($t) {
      Stop-ScheduledTask -TaskName $TaskName -ErrorAction Stop
      Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
      Write-Host "    Removed task '$TaskName'."
    } else { Write-Host "    Task '$TaskName' not present." }
  }
  if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
    throw "Scheduled task '$TaskName' survived unregistration; ownership receipt retained."
  }

  # 2. Stop the exact wscript -> cmd -> node supervisor chain owned by this
  # profile. Enumeration/termination failures are fatal and retain the receipt.
  Step '2. Stop owned bridge supervisor chain'
  Stop-OwnedBridgeProcesses
  Assert-OwnedBridgeQuiescent -OwnedPort $OwnedBridgePort

  # 3. Startup shortcut.
  Step '3. Remove Startup shortcut'
  $lnk = Join-Path ([Environment]::GetFolderPath('Startup')) 'Hermes Excel Bridge.lnk'
  if (Test-Path -LiteralPath $lnk) {
    Remove-Item -LiteralPath $lnk -Force -ErrorAction Stop
    Write-Host "    Removed '$lnk'."
  } else { Write-Host '    No Startup shortcut present.' }
  $doclingLnk = Join-Path ([Environment]::GetFolderPath('Startup')) 'Hermes Docling Serve.lnk'
  if (Test-Path -LiteralPath $doclingLnk) {
    Remove-Item -LiteralPath $doclingLnk -Force -ErrorAction Stop
    Write-Host "    Removed '$doclingLnk'."
  }
  if ((Test-Path -LiteralPath $lnk) -or (Test-Path -LiteralPath $doclingLnk)) {
    throw 'One or more owned Startup shortcuts survived removal; ownership receipt retained.'
  }

  # 4. Sideload (WEF Developer value + catalog manifest).
  Step '4. Remove developer sideload'
  $regSide = Join-Path $InstallDir 'install\register-sideload.ps1'
  if (-not (Test-Path -LiteralPath $regSide)) { $regSide = Join-Path $ScriptDir 'register-sideload.ps1' }
  if (Test-Path -LiteralPath $regSide) {
    & powershell -NoProfile -ExecutionPolicy Bypass -File $regSide -InstallDir $InstallDir -Unregister `
      -ProfileName $ProfileContext.ProfileName -HermesHome $ProfileContext.ProfileHome `
      -OwnerReceiptPath $ProfileContext.ReceiptPath -TransactionProof $TransactionProof
    if ($LASTEXITCODE -ne 0) { throw "Office add-in unregistration failed with exit code $LASTEXITCODE; payload retained." }
  } else {
    # Inline fallback if the installed script is already gone.
    $key = "HKCU:\Software\Microsoft\Office\$OfficeVersion\WEF\Developer"
    if (Test-Path -LiteralPath $key) {
      $existingWef = Get-ItemProperty -LiteralPath $key -Name 'HermesExcelAddinCatalog' -ErrorAction SilentlyContinue
      if ($null -ne $existingWef) {
        Remove-ItemProperty -LiteralPath $key -Name 'HermesExcelAddinCatalog' -Force -ErrorAction Stop
      }
      if ($null -ne (Get-ItemProperty -LiteralPath $key -Name 'HermesExcelAddinCatalog' -ErrorAction SilentlyContinue)) {
        throw 'Inline Office add-in unregistration did not remove HermesExcelAddinCatalog.'
      }
      Write-Host '    Removed WEF Developer value (inline).'
    }
  }

  # 5. Delete install dir.
  Step "5. Delete install dir"
  if (Test-Path -LiteralPath $InstallDir) {
    $resolvedInstall = [IO.Path]::GetFullPath($InstallDir)
    $expectedInstall = [IO.Path]::GetFullPath((Join-Path $ProfileContext.ProfileHome 'excel-addin'))
    if ($resolvedInstall -ne $expectedInstall -or $resolvedInstall -eq [IO.Path]::GetPathRoot($resolvedInstall)) {
      throw "Refusing unsafe rollback target: $resolvedInstall"
    }
    Stop-OwnedBridgeProcesses
    Assert-OwnedBridgeQuiescent -OwnedPort $OwnedBridgePort
    Remove-Item -LiteralPath $InstallDir -Recurse -Force -ErrorAction Stop
    Write-Host "    Deleted '$InstallDir'."
  } else {
    Write-Host '    Install dir already absent.'
  }

  # Clean the User env var regardless.
  [Environment]::SetEnvironmentVariable('HERMES_EXCEL_BRIDGE_TOKEN', $null, 'User')
  [Environment]::SetEnvironmentVariable('HERMES_EXCEL_INGEST_TOKEN', $null, 'User')
  [Environment]::SetEnvironmentVariable('HERMES_EXCEL_ALLOW_ALL_USERS', $null, 'User')
  Remove-Item Env:HERMES_EXCEL_BRIDGE_TOKEN -ErrorAction SilentlyContinue
  Remove-Item Env:HERMES_EXCEL_INGEST_TOKEN -ErrorAction SilentlyContinue
  Remove-Item Env:HERMES_EXCEL_ALLOW_ALL_USERS -ErrorAction SilentlyContinue
  $hermesCommand = Get-Command hermes -ErrorAction SilentlyContinue
  if ($hermesCommand) {
    & hermes gateway restart
    if ($LASTEXITCODE -ne 0) { throw 'Hermes gateway restart failed; ownership receipt retained because the old adapter may remain active.' }
  } else {
    throw 'Hermes CLI not found; ownership receipt retained until the gateway can be restarted safely.'
  }

  # Re-check immediately before retiring ownership. A surviving orphaned
  # supervisor or listener must keep the receipt authoritative.
  Assert-OwnedBridgeQuiescent -OwnedPort $OwnedBridgePort

  # The shared receipt is the last artifact removed. Any earlier failure leaves
  # it authoritative so another profile cannot claim a partially-rolled-back singleton.
  Remove-HermesExcelOwnerReceipt -Context $ProfileContext

  Write-Host ''
  Write-Host 'Rollback complete.' -ForegroundColor Green
  exit 0
}
catch {
  Write-Error "ROLLBACK FAILED: $($_.Exception.Message)"
  exit 1
}
finally {
  Remove-Item Env:HERMES_EXCEL_TRANSACTION_PROOF -ErrorAction SilentlyContinue
  Remove-Item Env:HERMES_EXCEL_TRANSACTION_PROFILE -ErrorAction SilentlyContinue
  Remove-Item Env:HERMES_EXCEL_TRANSACTION_HOME -ErrorAction SilentlyContinue
  Remove-Item Env:HERMES_EXCEL_TRANSACTION_RECEIPT -ErrorAction SilentlyContinue
  Exit-HermesExcelTransaction -Lock $TransactionLock
}
