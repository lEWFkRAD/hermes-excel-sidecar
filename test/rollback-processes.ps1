$ErrorActionPreference = 'Stop'
$path = Join-Path $PSScriptRoot '..\install\rollback.ps1'
$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw 'Rollback must parse before testing.' }
# Evaluate only the two process helpers. Never execute the uninstall entry point.
foreach ($name in @('Get-OwnedBridgeProcesses', 'Stop-OwnedBridgeProcesses')) {
  $definition = $ast.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name }, $true)
  if (-not $definition) { throw "Missing helper $name" }
  . ([scriptblock]::Create($definition.Extent.Text))
}
$InstallDir = 'C:\Test Install\excel-addin'
function Start-Sleep { param($Milliseconds) }
$script:processes = @(
  [pscustomobject]@{ Name='node.exe'; ProcessId=11; CommandLine='node "C:\Test Install\excel-addin\broker\server.mjs"' },
  [pscustomobject]@{ Name='cmd.exe'; ProcessId=12; CommandLine='cmd /c "C:\Test Install\excel-addin\service\bridge-service.cmd"' },
  [pscustomobject]@{ Name='wscript.exe'; ProcessId=13; CommandLine='wscript "C:\Test Install\excel-addin\service\bridge-service.vbs"' },
  [pscustomobject]@{ Name='node.exe'; ProcessId=21; CommandLine='node "C:\Test Install\excel-addin-old\broker\server.mjs"' },
  [pscustomobject]@{ Name='node.exe'; ProcessId=22; CommandLine='node "C:\Test Install\excel-addin\other.mjs"' },
  [pscustomobject]@{ Name='node.exe'; ProcessId=23; CommandLine='node "C:\Test Install\excel-addin\broker\server.mjs.bak"' }
)
function Get-CimInstance { param($ClassName, $ErrorAction) $script:processes }
$script:stopped = @()
function Stop-Process {
  param($Id, [switch]$Force, $ErrorAction)
  $script:stopped += $Id
  $script:processes = @($script:processes | Where-Object ProcessId -ne $Id)
}
function Get-Process { param($Id, $ErrorAction) $script:processes | Where-Object ProcessId -eq $Id }
$owned = @(Get-OwnedBridgeProcesses)
if (($owned.ProcessId -join ',') -ne '11,12,13') { throw 'Ownership matcher selected the wrong processes.' }
Stop-OwnedBridgeProcesses
if ($script:stopped[-1] -ne 11) { throw 'Node was stopped before its launchers.' }
if (($script:processes.ProcessId -join ',') -ne '21,22,23') { throw 'Unrelated processes were changed.' }
Write-Output 'PASS exact ownership and launcher-before-node cleanup'
Stop-OwnedBridgeProcesses
if ($script:stopped.Count -ne 3) { throw 'Idempotent cleanup stopped an unrelated process.' }
Write-Output 'PASS absent-process idempotence'
$script:processes = @([pscustomobject]@{ Name='cmd.exe'; ProcessId=12; CommandLine='cmd /c "C:\Test Install\excel-addin\service\bridge-service.cmd"' })
function Stop-Process { param($Id, [switch]$Force, $ErrorAction) throw 'simulated access denied' }
$refused = $false
try { Stop-OwnedBridgeProcesses } catch { $refused = $_.Exception.Message -match 'simulated access denied' }
if (-not $refused) { throw 'Stop failure was swallowed.' }
Write-Output 'PASS live-process stop failure aborts cleanup'
function Stop-Process { param($Id, [switch]$Force, $ErrorAction) }
$refused = $false
try { Stop-OwnedBridgeProcesses } catch { $refused = $_.Exception.Message -match 'survived termination' }
if (-not $refused) { throw 'Surviving process was not detected.' }
Write-Output 'PASS surviving supervisor fails closed after bounded retries'
