<#
.SYNOPSIS
  Fleet wrapper -- the one-line entry point for installing Hermes for Excel.

.DESCRIPTION
  requires elevation? NO.

  By default invokes addin-install.ps1. Lets a fleet tool opt a box out of the
  Excel add-in with -SkipExcelAddin, and passes -Port / -SkipSideload through.

  One-line install:
    powershell -ExecutionPolicy Bypass -File install\apply.ps1

.PARAMETER Port
  Bridge port (passed through). Default 0 selects the first available port
  starting at 8788.

.PARAMETER SkipExcelAddin
  Do not install the Excel add-in on this box (no-op success).

.PARAMETER SkipSideload
  Passed through to addin-install.ps1 (skip developer sideload registration).
#>
[CmdletBinding()]
param(
  [ValidateRange(0, 65535)]
  [int]    $Port           = 0,
  [switch] $SkipExcelAddin,
  [switch] $SkipSideload,
  [string] $ProfileName,
  [string] $HermesHome,
  [string] $OwnerReceiptPath,
  [switch] $AdoptLegacy
)

$ErrorActionPreference = 'Stop'

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Import-Module (Join-Path $ScriptDir 'profile-ownership.psm1') -Force

try {
  $context = Resolve-HermesExcelProfileContext -ProfileName $ProfileName -HermesHome $HermesHome -OwnerReceiptPath $OwnerReceiptPath
  if ($SkipExcelAddin) {
    Write-Host '[apply] -SkipExcelAddin set; this box opts out of the Excel add-in. Nothing to do.'
    exit 0
  }

  $installer = Join-Path $ScriptDir 'addin-install.ps1'
  if (-not (Test-Path -LiteralPath $installer)) {
    throw "addin-install.ps1 not found next to apply.ps1 ('$installer')."
  }

  $passthru = @{
    Port = $Port
    ProfileName = $context.ProfileName
    HermesHome = $context.ProfileHome
    OwnerReceiptPath = $context.ReceiptPath
  }
  if ($SkipSideload) { $passthru['SkipSideload'] = $true }
  if ($AdoptLegacy) { $passthru['AdoptLegacy'] = $true }

  $portLabel = if ($Port -eq 0) { 'auto (starting at 8788)' } else { "$Port" }
  Write-Host "[apply] Installing Hermes for Excel (port $portLabel)..."
  & powershell -NoProfile -ExecutionPolicy Bypass -File $installer @passthru
  $rc = $LASTEXITCODE
  if ($rc -ne 0) { throw "addin-install.ps1 exited with code $rc." }

  Write-Host '[apply] Done.'
  exit 0
}
catch {
  Write-Error "[apply] FAILED: $($_.Exception.Message)"
  exit 1
}
