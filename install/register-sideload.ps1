<# Register or unregister the Hermes manifest using Microsoft's supported
   per-user Office development registration tool. No elevation is required. #>
[CmdletBinding()]
param(
  [string] $InstallDir,
  [int] $Port = 8788,
  [switch] $Unregister,
  [switch] $RegisterExistingCatalog,
  [string] $ProfileName,
  [string] $HermesHome,
  [string] $OwnerReceiptPath,
  [string] $TransactionProof
)

$ErrorActionPreference = 'Stop'
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Import-Module (Join-Path $ScriptDir 'profile-ownership.psm1') -Force
$ProfileContext = Resolve-HermesExcelProfileContext -ProfileName $ProfileName -HermesHome $HermesHome -OwnerReceiptPath $OwnerReceiptPath
if ([string]::IsNullOrWhiteSpace($InstallDir)) { $InstallDir = $ProfileContext.InstallPath }
$InstallDir = [IO.Path]::GetFullPath($InstallDir).TrimEnd('\')
if (-not [string]::Equals($InstallDir, $ProfileContext.InstallPath, [StringComparison]::OrdinalIgnoreCase)) {
  throw 'InstallDir does not match the selected Hermes profile context.'
}
$OwnershipLock = $null

if ($Unregister -and $RegisterExistingCatalog) {
  throw 'Unregister and RegisterExistingCatalog are mutually exclusive.'
}

try {
if ($TransactionProof) {
  Assert-HermesExcelParentTransaction -Context $ProfileContext -Proof $TransactionProof
} else {
  $OwnershipLock = Enter-HermesExcelTransaction -Context $ProfileContext
  $owner = Assert-HermesExcelOwnership -Context $ProfileContext -Operation Mutate
  if (-not $owner) { throw 'Office sideload registration requires an existing owned install.' }
}
$CatalogDir = Join-Path $InstallDir 'OfficeAddinManifests'
$CatalogPath = Join-Path $CatalogDir 'hermes-excel-addin.xml'
$SourceManifest = Join-Path $InstallDir 'manifest.xml'
$LegacyKey = 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer'
$LegacyValue = 'HermesExcelAddinCatalog'
$ToolVersion = '3.1.2'

function Invoke-DevSettings([string[]] $Arguments) {
  $npx = (Get-Command npx.cmd -ErrorAction Stop).Source
  & $npx --yes "office-addin-dev-settings@$ToolVersion" @Arguments
  if ($LASTEXITCODE -ne 0) {
    throw "office-addin-dev-settings failed with exit code $LASTEXITCODE."
  }
}

function Remove-LegacyRegistration {
  if (Test-Path -LiteralPath $LegacyKey) {
    $legacy = Get-ItemProperty -LiteralPath $LegacyKey -Name $LegacyValue -ErrorAction SilentlyContinue
    if ($null -ne $legacy) {
      Remove-ItemProperty -LiteralPath $LegacyKey -Name $LegacyValue -Force -ErrorAction Stop
    }
    if ($null -ne (Get-ItemProperty -LiteralPath $LegacyKey -Name $LegacyValue -ErrorAction SilentlyContinue)) {
      throw "Legacy Office development registration '$LegacyValue' survived removal."
    }
  }
}

try {
  if ($Unregister) {
    if (Test-Path -LiteralPath $CatalogPath) {
      Invoke-DevSettings @('unregister', $CatalogPath)
      Remove-Item -LiteralPath $CatalogPath -Force -ErrorAction Stop
    }
    Remove-LegacyRegistration
    if (Test-Path -LiteralPath $CatalogPath) {
      throw "Office development catalog manifest survived removal: '$CatalogPath'."
    }
    if ($null -ne (Get-ItemProperty -LiteralPath $LegacyKey -Name $LegacyValue -ErrorAction SilentlyContinue)) {
      throw "Legacy Office development registration '$LegacyValue' survived removal."
    }
    Write-Host '[register-sideload] Hermes development manifest unregistered.'
    return
  }

  if ($RegisterExistingCatalog) {
    if (-not (Test-Path -LiteralPath $CatalogPath -PathType Leaf)) {
      throw "Existing Office development catalog manifest not found at '$CatalogPath'."
    }
    Invoke-DevSettings @('register', $CatalogPath)
    if (-not (Test-Path -LiteralPath $CatalogPath -PathType Leaf)) {
      throw "Office development catalog manifest disappeared during registration: '$CatalogPath'."
    }
    Write-Host "[register-sideload] Re-registered existing catalog '$CatalogPath' without rewriting legacy state."
    return
  }

  if (-not (Test-Path -LiteralPath $SourceManifest)) {
    throw "Source manifest not found at '$SourceManifest'."
  }
  New-Item -ItemType Directory -Path $CatalogDir -Force | Out-Null
  $xml = Get-Content -LiteralPath $SourceManifest -Raw
  $xml = $xml -replace 'localhost:8788', "localhost:$Port"
  Set-Content -LiteralPath $CatalogPath -Value $xml -Encoding UTF8 -Force

  Remove-LegacyRegistration
  Invoke-DevSettings @('register', $CatalogPath)
  Write-Host "[register-sideload] Registered '$CatalogPath' for Office development."
}
catch {
  Write-Error "[register-sideload] FAILED: $($_.Exception.Message)"
  exit 1
}
}
finally {
  Exit-HermesExcelTransaction -Lock $OwnershipLock
}
