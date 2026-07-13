<#
.SYNOPSIS
  Standalone, idempotent, per-box installer for the Hermes for Excel add-in.

.DESCRIPTION
  requires elevation? NO. Everything lives under %LOCALAPPDATA% and HKCU.

  Steps (in order):
    (a) Ensure Node.js LTS (winget install OpenJS.NodeJS.LTS if node missing).
    (b) Copy the apps/excel payload to the install dir (excludes uploads/exports/*.log).
    (c) Create the HERMES_EXCEL_DATA_DIR.
    (d) Mark-of-the-web unblock (Unblock-File) on the whole install tree.
    (e) Generate + persist a per-install HERMES_EXCEL_BRIDGE_TOKEN.
    (f) Emit run-bridge.cmd + run-bridge.vbs from templates (absolute paths,
        env baked in).
    (g) Register one tracked autostart supervisor via Scheduled Task
        (register-task.ps1), removing the legacy Startup shortcut if present.
        Any previously-running bridge from this install dir is stopped first.
    (h) Register the developer sideload (register-sideload.ps1).
    (i) Health-check: node ... --check-hermes, start the bridge, GET /api/health.
        Fail loudly (non-zero exit) if the bridge does not come up.

  Re-runnable: existence checks + -Force everywhere; never hard-fails on re-run.

.PARAMETER Port
  Bridge port. Use 0 (the default) to select the first available port starting
  at 8787. An explicitly selected occupied port fails instead of silently
  accepting an unrelated service.

.PARAMETER SkipSideload
  Skip the developer sideload registration (component 4).

.PARAMETER SourceDir
  Where the add-in payload is copied FROM. Defaults to the apps/excel tree that
  contains this script (this script lives in <SourceDir>\install).

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File install\addin-install.ps1
  powershell -ExecutionPolicy Bypass -File install\addin-install.ps1 -Port 8790
#>
[CmdletBinding()]
param(
  [ValidateRange(0, 65535)]
  [int]    $Port         = 0,
  [switch] $SkipSideload,
  [string] $SourceDir
)

$ErrorActionPreference = 'Stop'

# ============================================================================
#  Config block -- a deployer can retarget paths/port/modes here.
# ============================================================================
$ScriptDir   = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $SourceDir) { $SourceDir = Split-Path -Parent $ScriptDir }   # apps/excel
$InstallDir  = Join-Path $env:LOCALAPPDATA 'hermes\excel-addin'
$StageDir    = Join-Path $env:LOCALAPPDATA 'hermes\excel-addin.stage'
$BackupDir   = Join-Path $env:LOCALAPPDATA 'hermes\excel-addin.previous'
$HadExistingInstall = Test-Path -LiteralPath $InstallDir
$PreviousTask = Get-ScheduledTask -TaskName 'Hermes_Excel_Bridge' -ErrorAction SilentlyContinue
$HadPreviousTask = [bool]$PreviousTask
$PreviousTaskWasDisabled = $HadPreviousTask -and $PreviousTask.State -eq 'Disabled'
$WefDevKey = 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer'
$PreviousWef = Get-ItemProperty -LiteralPath $WefDevKey -Name 'HermesExcelAddinCatalog' -ErrorAction SilentlyContinue
$HadPreviousSideload = $null -ne $PreviousWef
$PreviousPort = 8787
if (Test-Path -LiteralPath (Join-Path $InstallDir 'run-bridge.cmd')) {
  $priorLauncher = Get-Content -LiteralPath (Join-Path $InstallDir 'run-bridge.cmd') -Raw
  if ($priorLauncher -match 'set "PORT=(\d+)"') { $PreviousPort = [int]$Matches[1] }
}
$DataDir     = Join-Path $InstallDir 'data'                            # HERMES_EXCEL_DATA_DIR
$TokenFile   = Join-Path $DataDir   '.bridge-token'                    # 0600-ish secret
$IngestTokenFile = Join-Path $DataDir '.ingest-token'
$DoclingMode = 'native'                                                # v1 uses in-body base64; no shared result path
$WslDistro   = 'Ubuntu'                                                # local Docling Serve runtime on this box
$NodeWingetId = 'OpenJS.NodeJS.LTS'
$HealthTimeoutSec = 45
$PreviousIngestToken = [Environment]::GetEnvironmentVariable('HERMES_EXCEL_INGEST_TOKEN', 'User')
$PreviousBridgeToken = [Environment]::GetEnvironmentVariable('HERMES_EXCEL_BRIDGE_TOKEN', 'User')
$PreviousAllowAllUsers = [Environment]::GetEnvironmentVariable('HERMES_EXCEL_ALLOW_ALL_USERS', 'User')
$GatewayRestartAttempted = $false
$ExcelPlatformEnableAttempted = $false
$PreviousExcelEnabled = $null
try {
  $cfgPath = Join-Path $env:LOCALAPPDATA 'hermes\config.yaml'
  $cfgText = if (Test-Path -LiteralPath $cfgPath) { Get-Content -LiteralPath $cfgPath -Raw } else { '' }
  if ($cfgText -match '(?ms)^excel:\s*\r?\n(?:^[ \t].*\r?\n)*?^[ \t]+enabled:\s*(true|false)\s*$') {
    $PreviousExcelEnabled = $Matches[1] -eq 'true'
  }
} catch { }
$PreferredPort = 8787
# Items copied; uploads/exports/logs are deliberately excluded.
$ExcludeDirs  = @('uploads', 'exports', 'node_modules', '.git')
$ExcludeFiles = @('*.log')
# ============================================================================

function Write-Step($msg) { Write-Host "==> $msg" -ForegroundColor Cyan }
function Write-Info($msg) { Write-Host "    $msg" }
function Write-Warn2($msg) { Write-Warning $msg }

function Protect-SecretFile([string]$Path) {
  $userSid = [System.Security.Principal.WindowsIdentity]::GetCurrent().User
  $systemSid = New-Object System.Security.Principal.SecurityIdentifier('S-1-5-18')
  # Set-Acl can request SeSecurityPrivilege even when the SACL is unchanged on
  # some non-admin Windows builds. icacls changes only the DACL needed here.
  & icacls.exe $Path /inheritance:r /grant:r "*$($userSid.Value):(F)" "*$($systemSid.Value):(F)" | Out-Null
  if ($LASTEXITCODE -ne 0) { throw "Could not secure secret ACL: $Path" }
  $verified = Get-Acl -LiteralPath $Path
  if (-not $verified.AreAccessRulesProtected) { throw "Secret ACL inheritance remains enabled: $Path" }
  $ownerSid = $verified.Owner | ForEach-Object { (New-Object System.Security.Principal.NTAccount($_)).Translate([System.Security.Principal.SecurityIdentifier]).Value }
  if ($ownerSid -ne $userSid.Value) { throw "Secret owner is not the current user: $Path" }
  $seen = @{}
  foreach ($rule in $verified.Access) {
    $sid = $rule.IdentityReference.Translate([System.Security.Principal.SecurityIdentifier]).Value
    if ($rule.AccessControlType -ne 'Allow' -or @($userSid.Value, $systemSid.Value) -notcontains $sid -or
        $rule.FileSystemRights -ne [System.Security.AccessControl.FileSystemRights]::FullControl) {
      throw "Unexpected secret ACL entry on '$Path': $sid $($rule.AccessControlType) $($rule.FileSystemRights)"
    }
    $seen[$sid] = $true
  }
  if (-not $seen[$userSid.Value] -or -not $seen[$systemSid.Value]) { throw "Secret ACL is missing current-user or SYSTEM FullControl: $Path" }
}

function Test-TcpPortAvailable([int]$Candidate) {
  $listener = $null
  try {
    $listener = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, $Candidate)
    $listener.Start()
    return $true
  } catch {
    return $false
  } finally {
    if ($listener) { try { $listener.Stop() } catch {} }
  }
}

function Resolve-BridgePort {
  if ($Port -ne 0) {
    if (-not (Test-TcpPortAvailable $Port)) {
      throw "Requested port $Port is already in use. Stop the conflicting service or rerun without -Port to auto-select."
    }
    return $Port
  }

  foreach ($candidate in $PreferredPort..($PreferredPort + 100)) {
    if (Test-TcpPortAvailable $candidate) {
      if ($candidate -ne $PreferredPort) {
        Write-Warn2 "Port $PreferredPort is occupied; selected $candidate for the Excel bridge."
      }
      return $candidate
    }
  }
  throw "No available bridge port found in $PreferredPort-$($PreferredPort + 100)."
}

# ---------------------------------------------------------------------------
# (a) Ensure Node.js LTS
# ---------------------------------------------------------------------------
function Ensure-Node {
  Write-Step '(a) Ensuring Node.js is available'
  $node = Get-Command node -ErrorAction SilentlyContinue
  if ($node) {
    $ver = (& node --version) 2>$null
    Write-Info "Node present: $ver ($($node.Source))"
    return
  }
  Write-Info 'node not on PATH; attempting install via winget.'
  $winget = Get-Command winget -ErrorAction SilentlyContinue
  if (-not $winget) {
    throw @"
Node.js is not installed and winget is not available.
Install Node.js LTS manually from https://nodejs.org/en/download (or
'App Installer' from the Microsoft Store to get winget), then re-run this script.
"@
  }
  Write-Info "Running: winget install $NodeWingetId"
  & winget install --id $NodeWingetId --silent --accept-source-agreements --accept-package-agreements
  # winget may have updated PATH only for new processes; probe common location.
  $node = Get-Command node -ErrorAction SilentlyContinue
  if (-not $node) {
    $candidate = Join-Path $env:ProgramFiles 'nodejs\node.exe'
    if (Test-Path -LiteralPath $candidate) {
      $env:PATH = "$($env:ProgramFiles)\nodejs;$env:PATH"
      $node = Get-Command node -ErrorAction SilentlyContinue
    }
  }
  if (-not $node) {
    throw "Node was installed by winget but is still not on PATH for this session. Open a new shell and re-run this installer."
  }
  Write-Info "Node now available: $(& node --version)"
}

# ---------------------------------------------------------------------------
# (b) Copy payload
# ---------------------------------------------------------------------------
function Copy-Payload {
  Write-Step "(b) Staging add-in payload for atomic activation"
  if (-not (Test-Path -LiteralPath (Join-Path $SourceDir 'broker\server.mjs'))) {
    throw "Source payload looks wrong: '$SourceDir\broker\server.mjs' not found."
  }
  Remove-Item -LiteralPath $StageDir -Recurse -Force -ErrorAction SilentlyContinue
  if (Test-Path -LiteralPath $BackupDir) {
    throw "Recovery backup already exists at '$BackupDir'. Resolve or preserve it manually before reinstalling."
  }
  New-Item -ItemType Directory -Path $StageDir -Force | Out-Null

  # robocopy is the most reliable mirror on Windows; /XD/XF exclude dirs/files.
  # Note: we do NOT /MIR so we never delete the live data dir on re-runs.
  $robo = Get-Command robocopy -ErrorAction SilentlyContinue
  if ($robo) {
    $args = @($SourceDir, $StageDir, '/E', '/PURGE', '/NJH', '/NJS', '/NDL', '/NP', '/R:1', '/W:1')
    foreach ($d in $ExcludeDirs)  { $args += @('/XD', (Join-Path $SourceDir $d)) }
    foreach ($f in $ExcludeFiles) { $args += @('/XF', $f) }
    & robocopy @args | Out-Null
    # robocopy exit codes 0-7 are success; 8+ is a real error.
    if ($LASTEXITCODE -ge 8) { throw "robocopy failed copying payload (exit $LASTEXITCODE)." }
  } else {
    # Fallback: manual recursive copy honoring excludes.
    Get-ChildItem -LiteralPath $SourceDir -Recurse -Force | ForEach-Object {
      $rel = $_.FullName.Substring($SourceDir.Length).TrimStart('\')
      $top = ($rel -split '\\')[0]
      if ($ExcludeDirs -contains $top) { return }
      if ($_.PSIsContainer) { return }
      foreach ($pat in $ExcludeFiles) { if ($_.Name -like $pat) { return } }
      $dest = Join-Path $StageDir $rel
      $destDir = Split-Path -Parent $dest
      if (-not (Test-Path -LiteralPath $destDir)) { New-Item -ItemType Directory -Path $destDir -Force | Out-Null }
      Copy-Item -LiteralPath $_.FullName -Destination $dest -Force
    }
  }
  if (-not (Test-Path -LiteralPath (Join-Path $StageDir 'broker\server.mjs')) -or
      -not (Test-Path -LiteralPath (Join-Path $StageDir 'manifest.xml'))) {
    throw 'Staged payload validation failed.'
  }
  Stop-ExistingBridge
  if ($HadExistingInstall) { Move-Item -LiteralPath $InstallDir -Destination $BackupDir }
  Move-Item -LiteralPath $StageDir -Destination $InstallDir
  if ($HadExistingInstall -and (Test-Path -LiteralPath (Join-Path $BackupDir 'data'))) {
    Move-Item -LiteralPath (Join-Path $BackupDir 'data') -Destination (Join-Path $InstallDir 'data')
  }
  if ($HadExistingInstall -and (Test-Path -LiteralPath (Join-Path $BackupDir 'OfficeAddinManifests'))) {
    Move-Item -LiteralPath (Join-Path $BackupDir 'OfficeAddinManifests') -Destination (Join-Path $InstallDir 'OfficeAddinManifests')
  }
  Write-Info 'Staged payload activated; previous version retained until certification passes.'
}

# ---------------------------------------------------------------------------
# (c) Data dir
# ---------------------------------------------------------------------------
function Ensure-DataDir {
  Write-Step "(c) Ensuring data dir '$DataDir'"
  if (-not (Test-Path -LiteralPath $DataDir)) {
    New-Item -ItemType Directory -Path $DataDir -Force | Out-Null
  }
  Write-Info 'Data dir ready (uploads/exports/logs live here, outside the web root).'
}

# ---------------------------------------------------------------------------
# (d) Mark-of-the-web unblock
# ---------------------------------------------------------------------------
function Unblock-Tree {
  Write-Step '(d) Unblocking files (mark-of-the-web)'
  Get-ChildItem -LiteralPath $InstallDir -Recurse -File -ErrorAction SilentlyContinue | Unblock-File -ErrorAction SilentlyContinue
  Write-Info 'Unblock-File applied across install tree.'
}

# ---------------------------------------------------------------------------
# (e) Bridge token
# ---------------------------------------------------------------------------
function Ensure-BridgeToken {
  Write-Step '(e) Ensuring per-install bridge token'
  if (Test-Path -LiteralPath $TokenFile) {
    $tok = (Get-Content -LiteralPath $TokenFile -Raw).Trim()
    if ($tok -notmatch '^[0-9a-f]{64}$') { throw 'Existing bridge token is invalid; remove it and reinstall to rotate.' }
    Protect-SecretFile $TokenFile
    [Environment]::SetEnvironmentVariable('HERMES_EXCEL_BRIDGE_TOKEN', $tok, 'User')
    $env:HERMES_EXCEL_BRIDGE_TOKEN = $tok
    Write-Info 'Existing bridge token validated and secured.'; return $tok
  }
  # Two GUIDs concatenated (hex, no braces) -> 64 hex chars of entropy.
  $tok = ([guid]::NewGuid().ToString('N')) + ([guid]::NewGuid().ToString('N'))
  Set-Content -LiteralPath $TokenFile -Value $tok -Encoding ASCII -Force -NoNewline

  Protect-SecretFile $TokenFile

  # Also expose as a user env var so other tooling/tests can read it.
  [Environment]::SetEnvironmentVariable('HERMES_EXCEL_BRIDGE_TOKEN', $tok, 'User')
  Write-Info 'Generated new bridge token (persisted to data dir + User env).'
  return $tok
}

function Ensure-IngestToken {
  Write-Step '(e2) Ensuring Hermes Excel adapter ingest token'
  if (Test-Path -LiteralPath $IngestTokenFile) {
    $tok = (Get-Content -LiteralPath $IngestTokenFile -Raw).Trim()
    if ($tok -notmatch '^[0-9a-f]{64}$') { throw 'Existing ingest token is invalid; remove it and reinstall to rotate.' }
    Protect-SecretFile $IngestTokenFile
    [Environment]::SetEnvironmentVariable('HERMES_EXCEL_INGEST_TOKEN', $tok, 'User')
    $env:HERMES_EXCEL_INGEST_TOKEN = $tok
    return $tok
  }
  $tok = ([guid]::NewGuid().ToString('N')) + ([guid]::NewGuid().ToString('N'))
  Set-Content -LiteralPath $IngestTokenFile -Value $tok -Encoding ASCII -Force -NoNewline
  Protect-SecretFile $IngestTokenFile
  [Environment]::SetEnvironmentVariable('HERMES_EXCEL_INGEST_TOKEN', $tok, 'User')
  $env:HERMES_EXCEL_INGEST_TOKEN = $tok
  Write-Info 'Generated adapter ingest token (ACL-restricted; shared with gateway and bridge).'
  return $tok
}

# ---------------------------------------------------------------------------
# (f) Emit run-bridge.cmd + run-bridge.vbs from templates
# ---------------------------------------------------------------------------
function Emit-Launchers([string]$BridgeToken, [string]$IngestToken) {
  Write-Step '(f) Generating run-bridge.cmd and run-bridge.vbs'
  $cmdTpl = Join-Path $InstallDir 'install\run-bridge.cmd.template'
  $vbsTpl = Join-Path $InstallDir 'install\run-bridge.vbs.template'
  if (-not (Test-Path -LiteralPath $cmdTpl)) { throw "Missing template '$cmdTpl'." }
  if (-not (Test-Path -LiteralPath $vbsTpl)) { throw "Missing template '$vbsTpl'." }

  # InstallDir without trailing backslash for clean substitution.
  $idir = $InstallDir.TrimEnd('\')

  $cmd = Get-Content -LiteralPath $cmdTpl -Raw
  $cmd = $cmd.Replace('__INSTALL_DIR__',   $idir)
  $cmd = $cmd.Replace('__PORT__',          "$Port")
  $cmd = $cmd.Replace('__DATA_DIR__',      $DataDir)
  $cmd = $cmd.Replace('__DOCLING_MODE__',  $DoclingMode)
  $cmd = $cmd.Replace('__WSL_DISTRO__',    $WslDistro)
  Set-Content -LiteralPath (Join-Path $InstallDir 'run-bridge.cmd') -Value $cmd -Encoding ASCII -Force

  $vbs = Get-Content -LiteralPath $vbsTpl -Raw
  $vbs = $vbs.Replace('__INSTALL_DIR__', $idir)
  Set-Content -LiteralPath (Join-Path $InstallDir 'run-bridge.vbs') -Value $vbs -Encoding ASCII -Force

  Write-Info 'Launchers written with absolute paths and env baked in.'
}

# ---------------------------------------------------------------------------
# Stop any bridge previously launched from this install dir.
# ---------------------------------------------------------------------------
function Stop-ExistingBridge {
  Write-Info 'Stopping any previously-running bridge from this install dir...'
  $idir = $InstallDir.TrimEnd('\')
  try {
    Get-CimInstance Win32_Process -Filter "Name='node.exe'" -ErrorAction SilentlyContinue |
      Where-Object { $_.CommandLine -and ($_.CommandLine -like "*$idir*") } |
      ForEach-Object {
        Write-Info "  killing stale bridge PID $($_.ProcessId)"
        Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue
      }
  } catch {
    Write-Warn2 "Could not enumerate node processes: $($_.Exception.Message)"
  }
}

# ---------------------------------------------------------------------------
# (g) Autostart: one tracked scheduled task
# ---------------------------------------------------------------------------
function Register-Autostart {
  Write-Step '(g) Registering the tracked Scheduled Task supervisor'
  Stop-ExistingBridge

  # Remove the legacy Startup shortcut. Running both it and the Scheduled Task
  # creates competing supervisors after logon and can cause restart churn.
  $startup = [Environment]::GetFolderPath('Startup')
  $lnkPath = Join-Path $startup 'Hermes Excel Bridge.lnk'
  if (Test-Path -LiteralPath $lnkPath) {
    Remove-Item -LiteralPath $lnkPath -Force -ErrorAction SilentlyContinue
    Write-Info "Removed legacy Startup shortcut '$lnkPath'."
  }

  # Scheduled Task (logon-triggered supervisor).
  $regTask = Join-Path $InstallDir 'install\register-task.ps1'
  if (Test-Path -LiteralPath $regTask) {
    & powershell -NoProfile -ExecutionPolicy Bypass -File $regTask `
      -InstallDir $InstallDir -DataDir $DataDir -Port $Port
    $taskExit = $LASTEXITCODE
    if ($taskExit -eq 5) {
      Stop-ScheduledTask -TaskName 'Hermes_Excel_Bridge' -ErrorAction SilentlyContinue
      Unregister-ScheduledTask -TaskName 'Hermes_Excel_Bridge' -Confirm:$false -ErrorAction SilentlyContinue
      $shell = New-Object -ComObject WScript.Shell
      $shortcut = $shell.CreateShortcut($lnkPath)
      $shortcut.TargetPath = (Join-Path $env:SystemRoot 'System32\wscript.exe')
      $shortcut.Arguments = '"' + (Join-Path $InstallDir 'run-bridge.vbs') + '"'
      $shortcut.WorkingDirectory = $InstallDir
      $shortcut.WindowStyle = 7
      $shortcut.Description = 'Hermes for Excel bridge supervisor'
      $shortcut.Save()
      if (-not (Test-Path -LiteralPath $lnkPath)) { throw 'Startup supervisor fallback was not created.' }
      Write-Warn2 'Task Scheduler denied current-user registration; installed one per-user Startup supervisor fallback.'
    } elseif ($taskExit -ne 0) {
      throw "Scheduled Task registration failed (exit $taskExit)."
    }
  } else {
    Write-Warn2 "register-task.ps1 not found at '$regTask'; scheduled task skipped."
  }
  $taskPresent = $null -ne (Get-ScheduledTask -TaskName 'Hermes_Excel_Bridge' -ErrorAction SilentlyContinue)
  $shortcutPresent = Test-Path -LiteralPath $lnkPath
  if ([int]$taskPresent + [int]$shortcutPresent -ne 1) {
    throw 'Supervisor invariant failed: expected exactly one Scheduled Task or Startup shortcut.'
  }
}

function Ensure-DoclingAutostart {
  Write-Step '(f1) Configuring independent Docling Serve startup when available'
  $doclingExe = '/root/.local/share/docling-serve/venv/bin/docling-serve'
  & wsl.exe -d $WslDistro -- test -x $doclingExe 2>$null
  if ($LASTEXITCODE -ne 0) {
    Write-Warn2 'Docling Serve is not installed in WSL; attachment parsing remains optional and ordinary edits stay available.'
    return
  }
  $shortcutPath = Join-Path ([Environment]::GetFolderPath('Startup')) 'Hermes Docling Serve.lnk'
  $shell = New-Object -ComObject WScript.Shell
  $shortcut = $shell.CreateShortcut($shortcutPath)
  $shortcut.TargetPath = (Join-Path $env:SystemRoot 'System32\wsl.exe')
  $shortcut.Arguments = "-d $WslDistro -- env DOCLING_DEVICE=cpu $doclingExe run --host 127.0.0.1 --port 8200"
  $shortcut.WorkingDirectory = $env:LOCALAPPDATA
  $shortcut.WindowStyle = 7
  $shortcut.Description = 'Independent CPU Docling Serve for Hermes Excel attachments'
  $shortcut.Save()
  Write-Info "Docling startup launcher ready: '$shortcutPath'."
}

function Ensure-OfficeTlsCertificate {
  Write-Step '(g1) Ensuring trusted localhost TLS certificate for Office'
  $certDir = Join-Path $env:USERPROFILE '.office-addin-dev-certs'
  $cert = Join-Path $certDir 'localhost.crt'
  $key = Join-Path $certDir 'localhost.key'
  $ca = Join-Path $certDir 'ca.crt'
  $createdByInstaller = -not ((Test-Path -LiteralPath $cert) -and (Test-Path -LiteralPath $key) -and (Test-Path -LiteralPath $ca))
  if ($createdByInstaller) {
    $npx = (Get-Command npx.cmd -ErrorAction Stop).Source
    & $npx --yes 'office-addin-dev-certs@2.0.10' install --days 365
    if ($LASTEXITCODE -ne 0) { throw "TLS certificate generation failed with exit code $LASTEXITCODE." }
  }
  $caObject = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($ca)
  $preTrusted = Test-Path -LiteralPath ("Cert:\CurrentUser\Root\" + $caObject.Thumbprint)
  & certutil.exe -user -f -addstore Root $ca | Out-Null
  if ($LASTEXITCODE -ne 0) { throw "Could not trust the Office development CA (certutil exit $LASTEXITCODE)." }
  $leaf = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($cert)
  if ($leaf.NotAfter -le (Get-Date).AddDays(7)) { throw 'Office TLS certificate is expired or expires within seven days.' }
  $san = ($leaf.Extensions | Where-Object { $_.Oid.Value -eq '2.5.29.17' }).Format($false)
  if ($san -notmatch '(?i)(DNS Name=localhost|DNS:localhost)') { throw 'Office TLS certificate lacks a localhost SAN.' }
  $chain = [System.Security.Cryptography.X509Certificates.X509Chain]::new()
  $chainOk = $chain.Build($leaf)
  $unexpectedChainErrors = @($chain.ChainStatus | Where-Object { $_.Status -ne 'RevocationStatusUnknown' })
  if ($unexpectedChainErrors.Count -gt 0 -or $chain.ChainElements.Count -lt 2 -or
      $chain.ChainElements[$chain.ChainElements.Count - 1].Certificate.Thumbprint -ne $caObject.Thumbprint) {
    throw 'Office TLS certificate does not chain to the exact trusted development CA.'
  }
  $principal = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
  & icacls.exe $key /inheritance:r /grant:r "${principal}:(R)" 'SYSTEM:(F)' | Out-Null
  if ($LASTEXITCODE -ne 0) { throw 'Could not restrict the localhost private-key ACL.' }
  & icacls.exe $key /remove:g '*S-1-5-32-544' | Out-Null
  $receipt = [ordered]@{ schema = 1; leaf_thumbprint = $leaf.Thumbprint; ca_thumbprint = $caObject.Thumbprint;
    cert_created_by_installer = $createdByInstaller; ca_was_trusted_before = $preTrusted;
    trust_policy = 'Office development CA intentionally persists because it may be shared by other add-ins.' }
  $receipt | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $DataDir 'tls-receipt.json') -Encoding UTF8
  Write-Info 'Trusted localhost TLS certificate is ready.'
}

# ---------------------------------------------------------------------------
# (h) Developer sideload
# ---------------------------------------------------------------------------
function Register-Sideload {
  if ($SkipSideload) { Write-Step '(h) Skipping sideload (-SkipSideload)'; return }
  Write-Step '(h) Registering developer sideload'
  $regSide = Join-Path $InstallDir 'install\register-sideload.ps1'
  if (Test-Path -LiteralPath $regSide) {
    & powershell -NoProfile -ExecutionPolicy Bypass -File $regSide `
      -InstallDir $InstallDir -Port $Port
  } else {
    Write-Warn2 "register-sideload.ps1 not found at '$regSide'; sideload skipped."
  }
}

# ---------------------------------------------------------------------------
# (i) Health check
# ---------------------------------------------------------------------------
function Invoke-HealthCheck([string]$BridgeToken) {
  Write-Step '(i) Health check'
  $server = Join-Path $InstallDir 'broker\server.mjs'

  # --check-hermes is allowed to fail (the Hermes gateway may simply be down);
  # we report but do not abort on it.
  Write-Info 'Running: node broker\server.mjs --check-hermes'
  try {
    & node $server --check-hermes
    if ($LASTEXITCODE -eq 0) { Write-Info 'Hermes endpoint: OK' }
    else { Write-Warn2 'Hermes endpoint check failed (gateway may be down). Continuing.' }
  } catch {
    Write-Warn2 "Hermes endpoint check error: $($_.Exception.Message). Continuing."
  }

  # Start the bridge via the generated launcher (hidden), then poll /api/health.
  Write-Info "Starting bridge and probing https://localhost:$Port/api/health ..."
  $vbs = Join-Path $InstallDir 'run-bridge.vbs'
  & (Join-Path $env:SystemRoot 'System32\wscript.exe') "$vbs"

  $deadline = (Get-Date).AddSeconds($HealthTimeoutSec)
  $health = $null
  while ((Get-Date) -lt $deadline) {
    Start-Sleep -Milliseconds 800
    try {
      $resp = Invoke-RestMethod -Uri "https://localhost:$Port/api/health" -Headers @{ 'X-Hermes-Token' = $BridgeToken } -TimeoutSec 3 -Method Get
      if ($resp -and $resp.service -eq 'hermes-excel-bridge' -and [int]$resp.port -eq $Port -and
          $resp.hermes_adapter_ready -eq $true -and $resp.raw_fallback_enabled -eq $false) {
        $health = $resp
        break
      }
    } catch { }
  }

  if (-not $health) {
    throw "Bridge did NOT respond on https://localhost:$Port/api/health within $HealthTimeoutSec s. Install FAILED."
  }

  $hermesOk  = if ($health.hermes_adapter_ready) { 'OK' } else { 'DOWN' }
  $doclingOk = if ($health.docling -and $health.docling.ok) { 'OK' } else { 'DOWN' }
  Write-Info 'Running typed Excel adapter certification turn...'
  $certId = [guid]::NewGuid().ToString('N')
  $certBody = @{
    request_id = "request-cert-$certId"
    workbook_id = "workbook-cert-$certId"
    conversation_id = "conversation-cert-$certId"
    prompt = 'Certification only: make no workbook changes. Return an empty actions proposal and a short readiness message.'
    history = @()
    workbook = @{ activeSheet = 'Sheet1'; sheets = @(@{ name = 'Sheet1'; usedRange = 'A1:A1' }) }
    selection = @{ address = 'Sheet1!A1'; values = @(@('')); formulas = @(@('')); rowCount = 1; columnCount = 1 }
    files = @()
    loop_count = 0
  } | ConvertTo-Json -Depth 8
  $cert = Invoke-RestMethod -Uri "https://localhost:$Port/api/chat" -Headers @{ 'X-Hermes-Token' = $BridgeToken } `
    -ContentType 'application/json' -Method Post -Body $certBody -TimeoutSec 180
  if ($cert.source -ne 'hermes-platform' -or $null -eq $cert.actions -or @($cert.actions).Count -ne 0 -or
      [string]::IsNullOrWhiteSpace([string]$cert.message) -or $null -ne $cert.fallback_reason) {
    throw "Typed Excel certification turn failed (source=$($cert.source), actions=$(@($cert.actions).Count))."
  }
  Write-Host ''
  Write-Host "    Bridge:  UP  (port $Port)" -ForegroundColor Green
  Write-Host "    Hermes:  $hermesOk"
  Write-Host "    Docling: $doclingOk"
}

# ============================================================================
#  Main
# ============================================================================
if (Test-Path -LiteralPath $BackupDir) {
  Write-Error "Recovery backup already exists at '$BackupDir'. Resolve or preserve it manually before reinstalling."
  exit 1
}

function Wait-ExcelAdapter([string]$IngestToken) {
  Write-Step '(f3) Verifying authenticated Excel platform adapter'
  $deadline = (Get-Date).AddSeconds(150)
  while ((Get-Date) -lt $deadline) {
    Start-Sleep -Milliseconds 750
    try {
      $health = Invoke-RestMethod -Uri 'http://127.0.0.1:8794/health' `
        -Headers @{ 'X-Excel-Token' = $IngestToken } -TimeoutSec 3 -Method Get
      if ($health.ok -eq $true -and $health.service -eq 'hermes-excel-adapter' -and
          [int]$health.protocol -eq 1 -and $health.capability -eq 'typed-proposals') {
        Write-Info 'Excel platform adapter: authenticated typed-proposals capability ready.'
        return
      }
    } catch { }
  }
  throw 'Excel platform adapter did not become ready on 127.0.0.1:8794 within 150 seconds.'
}
try {
  # A re-run may own the requested port already. Stop only bridges launched
  # from this install directory before deciding whether the port is available.
  if ($HadPreviousTask) {
    Stop-ScheduledTask -TaskName 'Hermes_Excel_Bridge' -ErrorAction SilentlyContinue
    Unregister-ScheduledTask -TaskName 'Hermes_Excel_Bridge' -Confirm:$false -ErrorAction SilentlyContinue
  }
  Stop-ExistingBridge
  $Port = Resolve-BridgePort
  Write-Host ''
  Write-Host "Hermes for Excel installer  (port $Port)" -ForegroundColor White
  Write-Host "Source : $SourceDir"
  Write-Host "Install: $InstallDir"
  Write-Host ''

  Ensure-Node
  Copy-Payload
  Ensure-DataDir
  Unblock-Tree
  $token = Ensure-BridgeToken
  $ingestToken = Ensure-IngestToken
  Emit-Launchers $token $ingestToken
  Ensure-DoclingAutostart
  Ensure-OfficeTlsCertificate
  $hermesCommand = Get-Command hermes -ErrorAction SilentlyContinue
  if ($hermesCommand) {
    Write-Step '(f2) Restarting Hermes gateway to activate Excel adapter'
    [Environment]::SetEnvironmentVariable('HERMES_EXCEL_ALLOW_ALL_USERS', 'true', 'User')
    $env:HERMES_EXCEL_ALLOW_ALL_USERS = 'true'
    & hermes config set excel.enabled true
    if ($LASTEXITCODE -ne 0) { throw 'Could not enable the Excel platform in Hermes config.' }
    $ExcelPlatformEnableAttempted = $true
    $GatewayRestartAttempted = $true
    & hermes gateway restart
    if ($LASTEXITCODE -ne 0) { throw 'Hermes gateway restart failed; Excel adapter was not activated.' }
  } else {
    throw 'Hermes CLI not found; cannot activate the Excel platform adapter.'
  }
  Wait-ExcelAdapter $ingestToken
  Register-Autostart
  Register-Sideload
  Invoke-HealthCheck $token
  if ($HadPreviousTask -and $PreviousTaskWasDisabled) {
    Disable-ScheduledTask -TaskName 'Hermes_Excel_Bridge' | Out-Null
    Stop-ExistingBridge
    Write-Info 'Preserved the operator-disabled bridge supervisor state after certification.'
  }
  Remove-Item -LiteralPath $BackupDir -Recurse -Force -ErrorAction SilentlyContinue

  Write-Host ''
  Write-Host 'Hermes for Excel installed successfully.' -ForegroundColor Green
  Write-Host "  - Restart Excel; open the pane from the Hermes ribbon group."
  Write-Host "  - Uninstall with: install\rollback.ps1"
  Write-Host ''
  exit 0
}
catch {
  $InstallError = $_.Exception.Message
  $RestoreFailed = $false
  try { Stop-ExistingBridge } catch { }
  # External state must be removed while cleanup scripts and the live payload
  # still exist. Payload restoration/deletion happens only after this block.
  try {
    Stop-ScheduledTask -TaskName 'Hermes_Excel_Bridge' -ErrorAction SilentlyContinue
    Unregister-ScheduledTask -TaskName 'Hermes_Excel_Bridge' -Confirm:$false -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath (Join-Path ([Environment]::GetFolderPath('Startup')) 'Hermes Excel Bridge.lnk') -Force -ErrorAction SilentlyContinue
    if (-not $HadExistingInstall) {
      & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $ScriptDir 'register-sideload.ps1') -InstallDir $InstallDir -Unregister
    }
  } catch { Write-Warning "Could not fully remove failed install external state: $($_.Exception.Message)" }
  try {
    Remove-Item -LiteralPath $StageDir -Recurse -Force -ErrorAction SilentlyContinue
    if (Test-Path -LiteralPath $BackupDir) {
      $liveData = Join-Path $InstallDir 'data'
      $backupData = Join-Path $BackupDir 'data'
      if ((Test-Path -LiteralPath $liveData) -and -not (Test-Path -LiteralPath $backupData)) {
        Move-Item -LiteralPath $liveData -Destination $backupData
      }
      $liveCatalog = Join-Path $InstallDir 'OfficeAddinManifests'
      $backupCatalog = Join-Path $BackupDir 'OfficeAddinManifests'
      if ((Test-Path -LiteralPath $liveCatalog) -and -not (Test-Path -LiteralPath $backupCatalog)) {
        Move-Item -LiteralPath $liveCatalog -Destination $backupCatalog
      }
      Remove-Item -LiteralPath $InstallDir -Recurse -Force -ErrorAction Stop
      Move-Item -LiteralPath $BackupDir -Destination $InstallDir -ErrorAction Stop
    } elseif (-not $HadExistingInstall) {
      Remove-Item -LiteralPath $InstallDir -Recurse -Force -ErrorAction SilentlyContinue
    }
  } catch {
    $RestoreFailed = $true
    Write-Warning "CRITICAL: could not restore previous Excel payload: $($_.Exception.Message). Backup: '$BackupDir'; live: '$InstallDir'."
  }
  try {
    Stop-ScheduledTask -TaskName 'Hermes_Excel_Bridge' -ErrorAction SilentlyContinue
    Unregister-ScheduledTask -TaskName 'Hermes_Excel_Bridge' -Confirm:$false -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath (Join-Path ([Environment]::GetFolderPath('Startup')) 'Hermes Excel Bridge.lnk') `
      -Force -ErrorAction SilentlyContinue
  } catch { }
  if ($HadExistingInstall -and -not $RestoreFailed) {
    try {
      if ($HadPreviousTask) {
        & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $ScriptDir 'register-task.ps1') `
          -InstallDir $InstallDir -DataDir (Join-Path $InstallDir 'data') -Port $PreviousPort
        if ($PreviousTaskWasDisabled) { Disable-ScheduledTask -TaskName 'Hermes_Excel_Bridge' | Out-Null }
      }
      if ($HadPreviousSideload) {
        & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $ScriptDir 'register-sideload.ps1') `
          -InstallDir $InstallDir -Port $PreviousPort
      }
    } catch {
      $RestoreFailed = $true
      Write-Warning "PARTIAL RESTORE: could not restore prior task/sideload state: $($_.Exception.Message)"
    }
  }
  [Environment]::SetEnvironmentVariable('HERMES_EXCEL_INGEST_TOKEN', $PreviousIngestToken, 'User')
  [Environment]::SetEnvironmentVariable('HERMES_EXCEL_BRIDGE_TOKEN', $PreviousBridgeToken, 'User')
  [Environment]::SetEnvironmentVariable('HERMES_EXCEL_ALLOW_ALL_USERS', $PreviousAllowAllUsers, 'User')
  if ($ExcelPlatformEnableAttempted) {
    try {
      if ($null -eq $PreviousExcelEnabled) { & hermes config set excel.enabled false | Out-Null }
      else { & hermes config set excel.enabled $PreviousExcelEnabled.ToString().ToLowerInvariant() | Out-Null }
    } catch { Write-Warning 'Could not restore prior Excel platform enabled state.' }
  }
  if ($PreviousIngestToken) { $env:HERMES_EXCEL_INGEST_TOKEN = $PreviousIngestToken }
  else { Remove-Item Env:HERMES_EXCEL_INGEST_TOKEN -ErrorAction SilentlyContinue }
  if ($PreviousBridgeToken) { $env:HERMES_EXCEL_BRIDGE_TOKEN = $PreviousBridgeToken }
  else { Remove-Item Env:HERMES_EXCEL_BRIDGE_TOKEN -ErrorAction SilentlyContinue }
  if ($GatewayRestartAttempted -and -not $RestoreFailed) {
    try { & hermes gateway restart | Out-Null } catch { Write-Warning 'Could not restore gateway environment after failed install.' }
  }
  Write-Host ''
  Write-Error "INSTALL FAILED: $InstallError"
  exit 1
}
