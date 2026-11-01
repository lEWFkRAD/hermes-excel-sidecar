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
  [switch] $Force
)

$ErrorActionPreference = 'Stop'

# ---- Config block ----------------------------------------------------------
$InstallDir = Join-Path $env:LOCALAPPDATA 'hermes\excel-addin'
$TaskName   = 'Hermes_Excel_Bridge'
$ScriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Path
# ----------------------------------------------------------------------------

function Step($msg) { Write-Host "==> $msg" -ForegroundColor Cyan }

function Get-OwnedBridgeProcesses([string] $Root) {
  $escapedRoot = [regex]::Escape([IO.Path]::GetFullPath($Root).TrimEnd('\'))
  # Exact executable-script paths only; a similarly named sibling install is
  # not ours. Include the cmd supervisor before its node child.
  $ownedScript = '(?i)(?:^|["\s])' + $escapedRoot + '[\\/](?:service[\\/]bridge-service\.(?:cmd|vbs)|run-bridge\.cmd|broker[\\/]server\.mjs)(?=["\s]|$)'
  @(Get-CimInstance Win32_Process -ErrorAction Stop | Where-Object {
    $_.Name -in @('wscript.exe', 'cmd.exe', 'node.exe') -and
    $_.CommandLine -and $_.CommandLine -match $ownedScript
  })
}

function Stop-OwnedBridgeProcesses([string] $Root) {
  # Scheduling must already be removed. Stop launchers first so a node killed
  # below cannot be replaced by a live restart loop.
  foreach ($pass in 1..3) {
    $owned = @(Get-OwnedBridgeProcesses $Root)
    if ($owned.Count -eq 0) { return }
    $ordered = $owned | Sort-Object @{ Expression = { if ($_.Name -eq 'node.exe') { 1 } else { 0 } } }
    foreach ($process in $ordered) {
      try { Stop-Process -Id $process.ProcessId -Force -ErrorAction Stop }
      catch {
        # A process exiting between enumeration and Stop-Process is harmless.
        if (Get-Process -Id $process.ProcessId -ErrorAction SilentlyContinue) { throw }
      }
    }
  }
  if (@(Get-OwnedBridgeProcesses $Root).Count -ne 0) {
    throw 'Bridge processes remain alive; rollback stopped and payload retained.'
  }
}


try {
  Write-Host ''
  Write-Host "Hermes for Excel rollback" -ForegroundColor White
  Write-Host "Install: $InstallDir"
  Write-Host ''

  # 1. Scheduled task (prefer the installed register-task.ps1; fall back inline).
  Step "1. Unregister scheduled task '$TaskName'"
  $regTask = Join-Path $InstallDir 'install\register-task.ps1'
  if (-not (Test-Path -LiteralPath $regTask)) { $regTask = Join-Path $ScriptDir 'register-task.ps1' }
  if (Test-Path -LiteralPath $regTask) {
    & powershell -NoProfile -ExecutionPolicy Bypass -File $regTask -InstallDir $InstallDir -Unregister
    if ($LASTEXITCODE -ne 0) { throw "Bridge task unregistration failed with exit code $LASTEXITCODE; payload retained." }
  } else {
    $t = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    if ($t) {
      Stop-ScheduledTask -TaskName $TaskName -ErrorAction Stop
      Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
      Write-Host "    Removed task '$TaskName'."
    } else { Write-Host "    Task '$TaskName' not present." }
  }

  # 2. Stop any running bridge from the install dir.
  Step '2. Stop running bridge'
  if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
    throw 'Bridge task is still registered; rollback stopped and payload retained.'
  }
  Stop-OwnedBridgeProcesses $InstallDir

  # 3. Startup shortcut.
  Step '3. Remove Startup shortcut'
  $lnk = Join-Path ([Environment]::GetFolderPath('Startup')) 'Hermes Excel Bridge.lnk'
  if (Test-Path -LiteralPath $lnk) {
    Remove-Item -LiteralPath $lnk -Force -ErrorAction SilentlyContinue
    Write-Host "    Removed '$lnk'."
  } else { Write-Host '    No Startup shortcut present.' }
  $doclingLnk = Join-Path ([Environment]::GetFolderPath('Startup')) 'Hermes Docling Serve.lnk'
  Remove-Item -LiteralPath $doclingLnk -Force -ErrorAction SilentlyContinue

  # 4. Sideload (WEF Developer value + catalog manifest).
  Step '4. Remove developer sideload'
  $regSide = Join-Path $InstallDir 'install\register-sideload.ps1'
  if (-not (Test-Path -LiteralPath $regSide)) { $regSide = Join-Path $ScriptDir 'register-sideload.ps1' }
  if (Test-Path -LiteralPath $regSide) {
    & powershell -NoProfile -ExecutionPolicy Bypass -File $regSide -InstallDir $InstallDir -Unregister
    if ($LASTEXITCODE -ne 0) { throw "Office add-in unregistration failed with exit code $LASTEXITCODE; payload retained." }
  } else {
    # Inline fallback if the installed script is already gone.
    $key = "HKCU:\Software\Microsoft\Office\$OfficeVersion\WEF\Developer"
    if (Test-Path -LiteralPath $key) {
      Remove-ItemProperty -LiteralPath $key -Name 'HermesExcelAddinCatalog' -Force -ErrorAction SilentlyContinue
      Write-Host '    Removed WEF Developer value (inline).'
    }
  }

  # 5. Delete install dir.
  Step "5. Delete install dir"
  if (Test-Path -LiteralPath $InstallDir) {
    $proceed = $Force
    if (-not $proceed) {
      $answer = Read-Host "    Delete '$InstallDir' and ALL its data (uploads/exports/logs/token)? [y/N]"
      $proceed = ($answer -match '^(y|yes)$')
    }
    if ($proceed) {
      $resolvedInstall = [IO.Path]::GetFullPath($InstallDir)
      $expectedInstall = [IO.Path]::GetFullPath((Join-Path $env:LOCALAPPDATA 'hermes\excel-addin'))
      if ($resolvedInstall -ne $expectedInstall -or $resolvedInstall -eq [IO.Path]::GetPathRoot($resolvedInstall)) {
        throw "Refusing unsafe rollback target: $resolvedInstall"
      }
      Stop-OwnedBridgeProcesses $resolvedInstall
      Remove-Item -LiteralPath $resolvedInstall -Recurse -Force -ErrorAction Stop
      Write-Host "    Deleted '$InstallDir'."
    } else {
      Write-Host '    Left install dir in place (not confirmed).'
    }
  } else {
    Write-Host '    Install dir already absent.'
  }

  # Clean the User env var regardless.
  [Environment]::SetEnvironmentVariable('HERMES_EXCEL_BRIDGE_TOKEN', $null, 'User')
  [Environment]::SetEnvironmentVariable('HERMES_EXCEL_INGEST_TOKEN', $null, 'User')
  Remove-Item Env:HERMES_EXCEL_BRIDGE_TOKEN -ErrorAction SilentlyContinue
  Remove-Item Env:HERMES_EXCEL_INGEST_TOKEN -ErrorAction SilentlyContinue
  $hermesCommand = Get-Command hermes -ErrorAction SilentlyContinue
  if ($hermesCommand) {
    & hermes gateway restart
    if ($LASTEXITCODE -ne 0) { Write-Warning 'Hermes gateway restart failed; the old adapter may remain active until the next restart.' }
  } else {
    Write-Warning 'Hermes CLI not found; restart the gateway manually to unload the Excel adapter.'
  }

  Write-Host ''
  Write-Host 'Rollback complete.' -ForegroundColor Green
  exit 0
}
catch {
  Write-Error "ROLLBACK FAILED: $($_.Exception.Message)"
  exit 1
}
