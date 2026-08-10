[CmdletBinding(SupportsShouldProcess)]
param(
  [string] $TargetRoot,
  [switch] $SkipChecks,
  [switch] $RestartBridge,
  [switch] $RestartGateway,
  [string] $ProfileName,
  [string] $HermesHome,
  [string] $OwnerReceiptPath
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$SourceRoot = [IO.Path]::GetFullPath((Split-Path -Parent $PSScriptRoot))
Import-Module (Join-Path $PSScriptRoot 'profile-ownership.psm1') -Force
$ProfileContext = Resolve-HermesExcelProfileContext -ProfileName $ProfileName -HermesHome $HermesHome -OwnerReceiptPath $OwnerReceiptPath
if ([string]::IsNullOrWhiteSpace($TargetRoot)) { $TargetRoot = $ProfileContext.InstallPath }
$TargetRoot = [IO.Path]::GetFullPath($TargetRoot)
$TargetParent = Split-Path -Parent $TargetRoot
$TransactionLock = $null

if (-not [string]::Equals($TargetRoot.TrimEnd('\'), $ProfileContext.InstallPath.TrimEnd('\'), [StringComparison]::OrdinalIgnoreCase)) {
  throw "Deploy target '$TargetRoot' must equal the selected profile install '$($ProfileContext.InstallPath)'."
}

if ($TargetRoot -eq $SourceRoot) {
  throw 'Deploy target must differ from the canonical source directory.'
}
if (-not $TargetParent -or $TargetRoot -eq [IO.Path]::GetPathRoot($TargetRoot)) {
  throw "Refusing unsafe deploy target: $TargetRoot"
}

& git -C $SourceRoot rev-parse --is-inside-work-tree *> $null
if ($LASTEXITCODE -ne 0) {
  throw 'Canonical source must be initialized as a Git worktree before deployment.'
}

if (-not $SkipChecks) {
  Write-Host '[deploy] Running canonical verification gate...'
  & npm.cmd run verify
  if ($LASTEXITCODE -ne 0) {
    throw "Verification gate failed with exit code $LASTEXITCODE."
  }
}

try {
$TransactionLock = Enter-HermesExcelTransaction -Context $ProfileContext
$ownerReceipt = Assert-HermesExcelOwnership -Context $ProfileContext -Operation Mutate
if (-not $ownerReceipt) { throw 'Canonical deployment requires an existing owned install. Run addin-install.ps1 first.' }
$env:HERMES_HOME = $ProfileContext.ProfileHome

$TrackedFiles = @(& git -C $SourceRoot ls-files)
if ($LASTEXITCODE -ne 0 -or $TrackedFiles.Count -eq 0) {
  throw 'Git returned no tracked payload files.'
}

$Copied = 0
$Unchanged = 0
foreach ($RelativePath in $TrackedFiles) {
  $SourcePath = Join-Path $SourceRoot $RelativePath
  if (-not (Test-Path -LiteralPath $SourcePath -PathType Leaf)) {
    continue
  }

  $TargetPath = Join-Path $TargetRoot $RelativePath
  $NeedsCopy = -not (Test-Path -LiteralPath $TargetPath -PathType Leaf)
  if (-not $NeedsCopy) {
    $SourceHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $SourcePath).Hash
    $TargetHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $TargetPath).Hash
    $NeedsCopy = $SourceHash -ne $TargetHash
  }

  if (-not $NeedsCopy) {
    $Unchanged += 1
    continue
  }

  if ($PSCmdlet.ShouldProcess($TargetPath, 'Deploy canonical tracked file')) {
    $TargetDir = Split-Path -Parent $TargetPath
    if (-not (Test-Path -LiteralPath $TargetDir)) {
      New-Item -ItemType Directory -Path $TargetDir -Force | Out-Null
    }
    Copy-Item -LiteralPath $SourcePath -Destination $TargetPath -Force
    $Copied += 1
  }
}

if (-not $WhatIfPreference) {
  foreach ($RelativePath in $TrackedFiles) {
    $SourcePath = Join-Path $SourceRoot $RelativePath
    if (-not (Test-Path -LiteralPath $SourcePath -PathType Leaf)) {
      continue
    }
    $TargetPath = Join-Path $TargetRoot $RelativePath
    if (-not (Test-Path -LiteralPath $TargetPath -PathType Leaf)) {
      throw "Deployment verification failed; target is missing $RelativePath"
    }
    $SourceHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $SourcePath).Hash
    $TargetHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $TargetPath).Hash
    if ($SourceHash -ne $TargetHash) {
      throw "Deployment verification failed; hash mismatch for $RelativePath"
    }
  }
}

Write-Host "[deploy] Complete: $Copied copied, $Unchanged already current. Runtime data and untracked launchers were preserved."

if ($RestartBridge -and $PSCmdlet.ShouldProcess('Hermes_Excel_Bridge', 'Restart Scheduled Task')) {
  Get-ScheduledTask -TaskName 'Hermes_Excel_Bridge' -ErrorAction Stop | Out-Null
  Stop-ScheduledTask -TaskName 'Hermes_Excel_Bridge' -ErrorAction SilentlyContinue

  $EscapedTarget = [regex]::Escape($TargetRoot.TrimEnd('\'))
  $BridgeProcesses = @(Get-CimInstance Win32_Process | Where-Object {
    $_.Name -in @('wscript.exe', 'cmd.exe', 'node.exe') -and
    $_.CommandLine -and
    $_.CommandLine -match $EscapedTarget -and
    $_.CommandLine -match '(bridge-service\.(vbs|cmd)|run-bridge\.cmd|broker[\\/]server\.mjs)'
  })
  foreach ($Process in $BridgeProcesses) {
    Stop-Process -Id $Process.ProcessId -Force -ErrorAction SilentlyContinue
  }

  $LockPath = [IO.Path]::GetFullPath((Join-Path $TargetRoot 'data\bridge-supervisor.lock'))
  if (-not $LockPath.StartsWith(($TargetRoot.TrimEnd('\') + '\'), [StringComparison]::OrdinalIgnoreCase)) {
    throw "Refusing unsafe supervisor lock path: $LockPath"
  }
  if (Test-Path -LiteralPath $LockPath) {
    Remove-Item -LiteralPath $LockPath -Recurse -Force
  }
  Start-ScheduledTask -TaskName 'Hermes_Excel_Bridge'
  Write-Host '[deploy] Bridge restart requested.'
}

if ($RestartGateway -and $PSCmdlet.ShouldProcess('Hermes gateway', 'Restart')) {
  & hermes gateway restart
  if ($LASTEXITCODE -ne 0) {
    throw "Hermes gateway restart failed with exit code $LASTEXITCODE."
  }
  Write-Host '[deploy] Gateway restart completed.'
}
}
finally {
  Exit-HermesExcelTransaction -Lock $TransactionLock
}
