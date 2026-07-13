<# Register or unregister the Hermes manifest using Microsoft's supported
   per-user Office development registration tool. No elevation is required. #>
[CmdletBinding()]
param(
  [string] $InstallDir = (Join-Path $env:LOCALAPPDATA 'hermes\excel-addin'),
  [int] $Port = 8787,
  [switch] $Unregister
)

$ErrorActionPreference = 'Stop'
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
    Remove-ItemProperty -LiteralPath $LegacyKey -Name $LegacyValue -Force -ErrorAction SilentlyContinue
  }
}

try {
  if ($Unregister) {
    if (Test-Path -LiteralPath $CatalogPath) {
      Invoke-DevSettings @('unregister', $CatalogPath)
      Remove-Item -LiteralPath $CatalogPath -Force -ErrorAction SilentlyContinue
    }
    Remove-LegacyRegistration
    Write-Host '[register-sideload] Hermes development manifest unregistered.'
    return
  }

  if (-not (Test-Path -LiteralPath $SourceManifest)) {
    throw "Source manifest not found at '$SourceManifest'."
  }
  New-Item -ItemType Directory -Path $CatalogDir -Force | Out-Null
  $xml = Get-Content -LiteralPath $SourceManifest -Raw
  $xml = $xml -replace 'localhost:8787', "localhost:$Port"
  Set-Content -LiteralPath $CatalogPath -Value $xml -Encoding UTF8 -Force

  Remove-LegacyRegistration
  Invoke-DevSettings @('register', $CatalogPath)
  Write-Host "[register-sideload] Registered '$CatalogPath' for Office development."
}
catch {
  Write-Error "[register-sideload] FAILED: $($_.Exception.Message)"
  exit 1
}
