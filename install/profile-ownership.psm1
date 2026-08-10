Set-StrictMode -Version Latest

$script:OwnerSchema = 1
$script:ManifestId = '4fd4d435-7f9a-4d6d-9251-32f154f83a1f'
$script:ReceiptFields = @(
  'schema', 'profile_name', 'profile_home', 'install_path',
  'bridge_port', 'plugin_version', 'manifest_id'
)
$script:ProfileNamePattern = '^[a-z0-9][a-z0-9_-]{0,63}$'
$script:SemVerPattern = '^(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)(-[0-9A-Za-z-]+(\.[0-9A-Za-z-]+)*)?(\+[0-9A-Za-z-]+(\.[0-9A-Za-z-]+)*)?$'
$script:MaxReceiptBytes = 16KB

function ConvertTo-HermesExcelCanonicalPath {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)]
    [string] $Path,
    [switch] $RequireCanonical
  )

  if ([string]::IsNullOrWhiteSpace($Path) -or $Path -ne $Path.Trim() -or $Path.IndexOf([char]0) -ge 0) {
    throw 'Profile paths must be non-empty and cannot contain surrounding whitespace or NUL bytes.'
  }
  if (-not [IO.Path]::IsPathRooted($Path)) {
    throw "Profile path must be absolute: '$Path'."
  }
  $full = [IO.Path]::GetFullPath($Path).TrimEnd([IO.Path]::DirectorySeparatorChar, [IO.Path]::AltDirectorySeparatorChar)
  if ($full.Length -eq 2 -and $full[1] -eq ':') { $full += [IO.Path]::DirectorySeparatorChar }
  if ($RequireCanonical -and -not [string]::Equals($Path, $full, [StringComparison]::OrdinalIgnoreCase)) {
    throw "Owner receipt path is not already canonical: '$Path'."
  }
  return $full
}

function Get-HermesExcelDefaultHome {
  if ([string]::IsNullOrWhiteSpace($env:LOCALAPPDATA)) {
    throw 'LOCALAPPDATA is required to resolve the shared Excel singleton owner.'
  }
  return ConvertTo-HermesExcelCanonicalPath -Path (Join-Path $env:LOCALAPPDATA 'hermes')
}

function Resolve-HermesExcelProfileContext {
  [CmdletBinding()]
  param(
    [string] $ProfileName,
    [string] $HermesHome,
    [string] $OwnerReceiptPath
  )

  $defaultHome = Get-HermesExcelDefaultHome
  $hint = if ([string]::IsNullOrWhiteSpace($ProfileName)) { $null } else { $ProfileName.Trim().ToLowerInvariant() }
  if ($hint -and $hint -notmatch $script:ProfileNamePattern) {
    throw 'ProfileName must match [a-z0-9][a-z0-9_-]{0,63}.'
  }

  $rawHome = if (-not [string]::IsNullOrWhiteSpace($HermesHome)) {
    $HermesHome
  } elseif (-not [string]::IsNullOrWhiteSpace($env:HERMES_HOME)) {
    $env:HERMES_HOME
  } else {
    $defaultHome
  }
  if ([string]::IsNullOrWhiteSpace($HermesHome) -and [string]::IsNullOrWhiteSpace($env:HERMES_HOME) -and $hint -and $hint -ne 'default') {
    throw 'HermesHome (or HERMES_HOME) is required for a non-default profile.'
  }
  $profileHome = ConvertTo-HermesExcelCanonicalPath -Path $rawHome

  if ([string]::Equals($profileHome, $defaultHome, [StringComparison]::OrdinalIgnoreCase)) {
    $inferred = 'default'
    $hermesRoot = $defaultHome
  } else {
    $profileParent = Split-Path -Parent $profileHome
    if ([string]::Equals((Split-Path -Leaf $profileParent), 'profiles', [StringComparison]::OrdinalIgnoreCase)) {
      $inferred = (Split-Path -Leaf $profileHome).ToLowerInvariant()
      if ($inferred -notmatch $script:ProfileNamePattern) {
        throw "HermesHome implies an invalid profile name: '$inferred'."
      }
      $hermesRoot = ConvertTo-HermesExcelCanonicalPath -Path (Split-Path -Parent $profileParent)
    } else {
      $inferred = 'custom'
      $hermesRoot = $profileHome
    }
  }

  if ($inferred -eq 'custom' -and $hint -eq 'default') {
    $inferred = 'default'
  } elseif ($hint -and $hint -ne $inferred) {
    throw "ProfileName '$hint' does not match the HermesHome identity '$inferred'."
  }

  $sharedReceipt = ConvertTo-HermesExcelCanonicalPath -Path (Join-Path $defaultHome 'shared\excel-sidecar\owner-v1.json')
  if (-not [string]::IsNullOrWhiteSpace($OwnerReceiptPath)) {
    $explicitReceipt = ConvertTo-HermesExcelCanonicalPath -Path $OwnerReceiptPath
    if (-not [string]::Equals($explicitReceipt, $sharedReceipt, [StringComparison]::OrdinalIgnoreCase)) {
      throw "OwnerReceiptPath must identify the per-user shared receipt '$sharedReceipt'."
    }
  }

  return [pscustomobject]@{
    ProfileName = $inferred
    ProfileHome = $profileHome
    HermesRoot = $hermesRoot
    InstallPath = ConvertTo-HermesExcelCanonicalPath -Path (Join-Path $profileHome 'excel-addin')
    DataPath = ConvertTo-HermesExcelCanonicalPath -Path (Join-Path $profileHome 'excel-addin\data')
    ConfigPath = ConvertTo-HermesExcelCanonicalPath -Path (Join-Path $profileHome 'config.yaml')
    ReceiptPath = $sharedReceipt
  }
}

function Read-HermesExcelOwnerReceipt {
  [CmdletBinding()]
  param([Parameter(Mandatory)] [string] $Path)

  $canonicalPath = ConvertTo-HermesExcelCanonicalPath -Path $Path
  if (-not (Test-Path -LiteralPath $canonicalPath -PathType Leaf)) { return $null }
  $item = Get-Item -LiteralPath $canonicalPath -Force
  if ($item.Length -gt $script:MaxReceiptBytes) { throw 'Excel owner receipt exceeds the 16 KiB limit.' }
  $rawBytes = [IO.File]::ReadAllBytes($canonicalPath)
  if ($rawBytes.Length -ge 3 -and $rawBytes[0] -eq 0xEF -and $rawBytes[1] -eq 0xBB -and $rawBytes[2] -eq 0xBF) {
    throw 'Excel owner receipt must be UTF-8 JSON without a byte-order mark.'
  }
  try { $raw = [Text.UTF8Encoding]::new($false, $true).GetString($rawBytes) } catch {
    throw 'Excel owner receipt must contain strictly valid UTF-8 JSON.'
  }

  # ConvertFrom-Json accepts duplicate keys. Count every simple JSON object key
  # first so a crafted receipt cannot override an earlier owner field.
  $keyMatches = [regex]::Matches($raw, '"([A-Za-z_][A-Za-z0-9_]*)"\s*:')
  $keyCounts = @{}
  foreach ($match in $keyMatches) {
    $key = $match.Groups[1].Value
    if ($keyCounts.ContainsKey($key)) { throw "Duplicate Excel owner receipt field: '$key'." }
    $keyCounts[$key] = 1
  }
  try { $receipt = $raw | ConvertFrom-Json -ErrorAction Stop } catch { throw "Excel owner receipt is not valid JSON: $($_.Exception.Message)" }
  if ($null -eq $receipt -or $receipt -is [Array]) { throw 'Excel owner receipt must be a JSON object.' }
  $actualFields = @($receipt.PSObject.Properties.Name)
  $missing = @($script:ReceiptFields | Where-Object { $actualFields -notcontains $_ })
  $extra = @($actualFields | Where-Object { $script:ReceiptFields -notcontains $_ })
  if ($missing.Count -or $extra.Count -or $keyCounts.Count -ne $script:ReceiptFields.Count) {
    throw "Excel owner receipt fields are invalid (missing: $($missing -join ', '); extra: $($extra -join ', '))."
  }
  if ($receipt.schema -isnot [int] -or $receipt.schema -ne $script:OwnerSchema) { throw 'Excel owner receipt schema must be the integer 1.' }
  if ($receipt.profile_name -isnot [string] -or $receipt.profile_name -notmatch $script:ProfileNamePattern -or
      $receipt.profile_name -cne $receipt.profile_name.ToLowerInvariant()) { throw 'Excel owner receipt profile_name is invalid.' }
  if ($receipt.profile_home -isnot [string] -or $receipt.install_path -isnot [string]) { throw 'Excel owner receipt paths must be strings.' }
  $profileHome = ConvertTo-HermesExcelCanonicalPath -Path $receipt.profile_home -RequireCanonical
  $installPath = ConvertTo-HermesExcelCanonicalPath -Path $receipt.install_path -RequireCanonical
  if ($receipt.bridge_port -isnot [int] -or $receipt.bridge_port -lt 1 -or $receipt.bridge_port -gt 65535) { throw 'Excel owner receipt bridge_port must be an integer from 1 through 65535.' }
  if ($receipt.plugin_version -isnot [string] -or $receipt.plugin_version -notmatch $script:SemVerPattern) { throw 'Excel owner receipt plugin_version is not semantic versioning.' }
  if ($receipt.manifest_id -isnot [string] -or -not [string]::Equals($receipt.manifest_id, $script:ManifestId, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'Excel owner receipt manifest_id is not the Hermes Excel add-in.'
  }

  return [pscustomobject]@{
    Schema = 1
    ProfileName = [string]$receipt.profile_name
    ProfileHome = $profileHome
    InstallPath = $installPath
    BridgePort = [int]$receipt.bridge_port
    PluginVersion = [string]$receipt.plugin_version
    ManifestId = $script:ManifestId
  }
}

function Test-HermesExcelOwnerMatch {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)] $Receipt,
    [Parameter(Mandatory)] $Context
  )
  return $Receipt.Schema -eq 1 -and
    [string]::Equals($Receipt.ManifestId, $script:ManifestId, [StringComparison]::OrdinalIgnoreCase) -and
    [string]::Equals($Receipt.ProfileName, $Context.ProfileName, [StringComparison]::Ordinal) -and
    [string]::Equals($Receipt.ProfileHome, $Context.ProfileHome, [StringComparison]::OrdinalIgnoreCase) -and
    [string]::Equals($Receipt.InstallPath, $Context.InstallPath, [StringComparison]::OrdinalIgnoreCase)
}

function Get-HermesExcelLegacyArtifacts {
  [CmdletBinding()]
  param([Parameter(Mandatory)] $Context)

  $artifacts = New-Object System.Collections.Generic.List[string]
  $candidateInstalls = @($Context.InstallPath, (Join-Path (Get-HermesExcelDefaultHome) 'excel-addin'))
  $profilesRoot = Join-Path (Get-HermesExcelDefaultHome) 'profiles'
  if (Test-Path -LiteralPath $profilesRoot -PathType Container) {
    $candidateInstalls += @(Get-ChildItem -LiteralPath $profilesRoot -Directory -ErrorAction SilentlyContinue | ForEach-Object { Join-Path $_.FullName 'excel-addin' })
  }
  foreach ($candidate in ($candidateInstalls | Select-Object -Unique)) {
    if (Test-Path -LiteralPath $candidate) { $artifacts.Add("install:$candidate") }
  }
  if (Get-ScheduledTask -TaskName 'Hermes_Excel_Bridge' -ErrorAction SilentlyContinue) { $artifacts.Add('scheduled-task:Hermes_Excel_Bridge') }
  foreach ($shortcutName in @('Hermes Excel Bridge.lnk', 'Hermes Docling Serve.lnk')) {
    $shortcut = Join-Path ([Environment]::GetFolderPath('Startup')) $shortcutName
    if (Test-Path -LiteralPath $shortcut) { $artifacts.Add("startup:$shortcutName") }
  }
  $wefKey = 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer'
  if (Test-Path -LiteralPath $wefKey) {
    $registryKey = Get-Item -LiteralPath $wefKey
    if ($registryKey.GetValueNames() -contains 'HermesExcelAddinCatalog') { $artifacts.Add('office-sideload:HermesExcelAddinCatalog') }
  }
  $environmentKey = Get-Item -LiteralPath 'HKCU:\Environment' -ErrorAction SilentlyContinue
  if ($environmentKey) {
    $names = @($environmentKey.GetValueNames())
    foreach ($name in @('HERMES_EXCEL_BRIDGE_TOKEN', 'HERMES_EXCEL_INGEST_TOKEN', 'HERMES_EXCEL_ALLOW_ALL_USERS')) {
      if ($names -contains $name) { $artifacts.Add("user-environment:$name") }
    }
  }
  return @($artifacts)
}

function Assert-HermesExcelOwnership {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)] $Context,
    [ValidateSet('Install', 'Mutate', 'ReadOnly')]
    [string] $Operation = 'Mutate',
    [switch] $AdoptLegacy
  )

  if ($AdoptLegacy -and $Operation -ne 'Install') { throw 'AdoptLegacy is permitted only for installation.' }
  $receipt = Read-HermesExcelOwnerReceipt -Path $Context.ReceiptPath
  if ($receipt) {
    if (-not (Test-HermesExcelOwnerMatch -Receipt $receipt -Context $Context)) {
      throw "Excel singleton is owned by profile '$($receipt.ProfileName)' at '$($receipt.ProfileHome)'; requested profile '$($Context.ProfileName)' may not inspect tokens or mutate it."
    }
    return $receipt
  }

  $legacy = @(Get-HermesExcelLegacyArtifacts -Context $Context)
  if ($legacy.Count -gt 0 -and -not ($Operation -eq 'Install' -and $AdoptLegacy)) {
    throw "Excel singleton artifacts exist without an owner receipt. Refusing $Operation. Run the installer once with explicit -AdoptLegacy after verifying the artifacts belong to this profile. Artifacts: $($legacy -join '; ')"
  }
  if ($legacy.Count -gt 0 -and $Operation -eq 'Install' -and $AdoptLegacy) {
    Assert-HermesExcelLegacyAdoption -Context $Context
  }
  return $null
}

function Assert-HermesExcelLegacyAdoption {
  [CmdletBinding()]
  param([Parameter(Mandatory)] $Context)

  $expectedInstall = $Context.InstallPath
  $defaultHome = Get-HermesExcelDefaultHome
  $candidateInstalls = @($expectedInstall, (Join-Path $defaultHome 'excel-addin'))
  $profilesRoot = Join-Path $defaultHome 'profiles'
  if (Test-Path -LiteralPath $profilesRoot -PathType Container) {
    $candidateInstalls += @(Get-ChildItem -LiteralPath $profilesRoot -Directory -ErrorAction SilentlyContinue | ForEach-Object { Join-Path $_.FullName 'excel-addin' })
  }
  $existingInstalls = @($candidateInstalls | Select-Object -Unique | Where-Object { Test-Path -LiteralPath $_ })
  $divergent = @($existingInstalls | Where-Object { -not [string]::Equals((ConvertTo-HermesExcelCanonicalPath -Path $_), $expectedInstall, [StringComparison]::OrdinalIgnoreCase) })
  if ($divergent.Count) {
    throw "AdoptLegacy found an install belonging to another path: $($divergent -join ', '). Move or recover it explicitly; live takeover is not supported."
  }
  if (-not (Test-Path -LiteralPath $expectedInstall -PathType Container)) {
    throw "AdoptLegacy requires the legacy payload at the requested install path '$expectedInstall'."
  }
  $manifest = Join-Path $expectedInstall 'manifest.xml'
  if (-not (Test-Path -LiteralPath $manifest -PathType Leaf)) { throw "AdoptLegacy cannot verify the manifest at '$manifest'." }
  try { $manifestXml = [xml](Get-Content -LiteralPath $manifest -Raw) } catch { throw "AdoptLegacy manifest is invalid XML: $($_.Exception.Message)" }
  if ($manifestXml.OfficeApp.Id -ne $script:ManifestId) { throw 'AdoptLegacy manifest ID is not the Hermes Excel add-in.' }

  $task = Get-ScheduledTask -TaskName 'Hermes_Excel_Bridge' -ErrorAction SilentlyContinue
  if ($task) {
    $taskActions = @($task.Actions)
    if ($taskActions.Count -ne 1) { throw 'AdoptLegacy requires exactly one verifiable scheduled-task action.' }
    foreach ($action in $taskActions) {
      $working = [string]$action.WorkingDirectory
      $arguments = [string]$action.Arguments
      if (-not [string]::Equals((ConvertTo-HermesExcelCanonicalPath -Path $working), $expectedInstall, [StringComparison]::OrdinalIgnoreCase) -or
          $arguments.IndexOf($expectedInstall, [StringComparison]::OrdinalIgnoreCase) -lt 0) {
        throw "AdoptLegacy scheduled task is not anchored to '$expectedInstall'."
      }
    }
  }

  $shortcut = Join-Path ([Environment]::GetFolderPath('Startup')) 'Hermes Excel Bridge.lnk'
  if (Test-Path -LiteralPath $shortcut -PathType Leaf) {
    $shell = New-Object -ComObject WScript.Shell
    $link = $shell.CreateShortcut($shortcut)
    if (([string]$link.Arguments).IndexOf($expectedInstall, [StringComparison]::OrdinalIgnoreCase) -lt 0 -and
        ([string]$link.TargetPath).IndexOf($expectedInstall, [StringComparison]::OrdinalIgnoreCase) -lt 0) {
      throw "AdoptLegacy Startup shortcut is not anchored to '$expectedInstall'."
    }
  }

  $doclingShortcut = Join-Path ([Environment]::GetFolderPath('Startup')) 'Hermes Docling Serve.lnk'
  if (Test-Path -LiteralPath $doclingShortcut -PathType Leaf) {
    $shell = New-Object -ComObject WScript.Shell
    $link = $shell.CreateShortcut($doclingShortcut)
    $expectedWsl = ConvertTo-HermesExcelCanonicalPath -Path (Join-Path $env:SystemRoot 'System32\wsl.exe')
    $reportedWsl = ConvertTo-HermesExcelCanonicalPath -Path ([string]$link.TargetPath)
    $doclingArguments = [string]$link.Arguments
    if (-not [string]::Equals($reportedWsl, $expectedWsl, [StringComparison]::OrdinalIgnoreCase) -or
        $doclingArguments -notmatch '^-d [A-Za-z0-9._-]+ -- env DOCLING_DEVICE=cpu /root/\.local/share/docling-serve/venv/bin/docling-serve run --host 127\.0\.0\.1 --port 8200$') {
      throw 'AdoptLegacy Docling Startup shortcut is not the fixed loopback Docling service.'
    }
  }

  $wefKey = 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer'
  if (Test-Path -LiteralPath $wefKey) {
    $key = Get-Item -LiteralPath $wefKey
    if ($key.GetValueNames() -contains 'HermesExcelAddinCatalog') {
      $catalogValue = [string]$key.GetValue('HermesExcelAddinCatalog', $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
      if ([string]::IsNullOrWhiteSpace($catalogValue) -or $catalogValue.IndexOf($expectedInstall, [StringComparison]::OrdinalIgnoreCase) -lt 0) {
        throw "AdoptLegacy Office registration is not anchored to '$expectedInstall'."
      }
    }
  }
  $catalogManifest = Join-Path $expectedInstall 'OfficeAddinManifests\hermes-excel-addin.xml'
  if (Test-Path -LiteralPath $catalogManifest -PathType Leaf) {
    try { $catalogXml = [xml](Get-Content -LiteralPath $catalogManifest -Raw) } catch { throw "AdoptLegacy catalog manifest is invalid XML: $($_.Exception.Message)" }
    if ($catalogXml.OfficeApp.Id -ne $script:ManifestId) { throw 'AdoptLegacy catalog manifest ID is not the Hermes Excel add-in.' }
  }
}

function Enter-HermesExcelTransaction {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)] $Context,
    [string] $Proof
  )

  $lockDir = Split-Path -Parent $Context.ReceiptPath
  if (-not (Test-Path -LiteralPath $lockDir)) { New-Item -ItemType Directory -Path $lockDir -Force | Out-Null }
  $sid = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
  $lockPath = Join-Path $lockDir "transaction-$sid.lock"
  Protect-HermesExcelOwnerFile -Path $lockDir
  if (-not (Test-Path -LiteralPath $lockPath)) {
    [IO.File]::WriteAllText($lockPath, '', [Text.UTF8Encoding]::new($false))
  }
  Protect-HermesExcelOwnerFile -Path $lockPath
  try {
    $stream = [IO.File]::Open($lockPath, [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
  } catch {
    throw "Another Hermes Excel ownership transaction is running for this Windows user ('$lockPath')."
  }
  try {
    $stream.SetLength(0)
    $proofLine = if ($Proof) { "proof=$Proof`n" } else { '' }
    $bytes = [Text.Encoding]::ASCII.GetBytes("pid=$PID`nstarted=$([DateTime]::UtcNow.ToString('o'))`nprofile=$($Context.ProfileName)`n$proofLine")
    $stream.Write($bytes, 0, $bytes.Length)
    $stream.Flush($true)
    return $stream
  } catch {
    $stream.Dispose()
    throw
  }
}

function Get-HermesExcelOwnerFingerprint {
  [CmdletBinding()]
  param([Parameter(Mandatory)] $Receipt)
  function ConvertTo-AsciiJsonString([string] $Value) {
    $builder = New-Object Text.StringBuilder
    [void]$builder.Append('"')
    foreach ($character in $Value.ToCharArray()) {
      $code = [int][char]$character
      $escaped = $null
      switch ($code) {
        8 { $escaped = '\b' }
        9 { $escaped = '\t' }
        10 { $escaped = '\n' }
        12 { $escaped = '\f' }
        13 { $escaped = '\r' }
        34 { $escaped = '\"' }
        92 { $escaped = '\\' }
      }
      if ($null -ne $escaped) { [void]$builder.Append($escaped); continue }
      if ($code -lt 32 -or $code -gt 126) { [void]$builder.Append(('\u{0:x4}' -f $code)) }
      else { [void]$builder.Append($character) }
    }
    [void]$builder.Append('"')
    return $builder.ToString()
  }
  $installKey = ([string]$Receipt.InstallPath).Replace('/', '\').ToLowerInvariant()
  $profileKey = ([string]$Receipt.ProfileHome).Replace('/', '\').ToLowerInvariant()
  # Match Python json.dumps(sort_keys=True, separators=(',', ':')) exactly.
  $json = '{"bridge_port":' + [int]$Receipt.BridgePort +
    ',"install_path":' + (ConvertTo-AsciiJsonString $installKey) +
    ',"manifest_id":' + (ConvertTo-AsciiJsonString ([string]$Receipt.ManifestId).ToLowerInvariant()) +
    ',"plugin_version":' + (ConvertTo-AsciiJsonString ([string]$Receipt.PluginVersion)) +
    ',"profile_home":' + (ConvertTo-AsciiJsonString $profileKey) +
    ',"profile_name":' + (ConvertTo-AsciiJsonString ([string]$Receipt.ProfileName)) +
    ',"schema":' + [int]$Receipt.Schema + '}'
  $sha = [Security.Cryptography.SHA256]::Create()
  try { $hash = $sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($json)) } finally { $sha.Dispose() }
  return 'sha256:' + (($hash | ForEach-Object { $_.ToString('x2') }) -join '')
}

function Exit-HermesExcelTransaction {
  [CmdletBinding()]
  param($Lock)
  if ($Lock) {
    try { $Lock.SetLength(0); $Lock.Flush($true) } catch { }
    $Lock.Dispose()
  }
}

function Assert-HermesExcelParentTransaction {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)] $Context,
    [Parameter(Mandatory)] [string] $Proof
  )
  if ([string]::IsNullOrWhiteSpace($Proof) -or
      -not [string]::Equals($Proof, $env:HERMES_EXCEL_TRANSACTION_PROOF, [StringComparison]::Ordinal) -or
      -not [string]::Equals($Context.ProfileName, $env:HERMES_EXCEL_TRANSACTION_PROFILE, [StringComparison]::Ordinal) -or
      -not [string]::Equals($Context.ProfileHome, $env:HERMES_EXCEL_TRANSACTION_HOME, [StringComparison]::OrdinalIgnoreCase) -or
      -not [string]::Equals($Context.ReceiptPath, $env:HERMES_EXCEL_TRANSACTION_RECEIPT, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'The parent ownership transaction proof is absent or does not match this profile context.'
  }
  $sid = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
  $lockPath = Join-Path (Split-Path -Parent $Context.ReceiptPath) "transaction-$sid.lock"
  if (-not (Test-Path -LiteralPath $lockPath -PathType Leaf)) { throw 'The parent ownership transaction lock file is absent.' }
  $probe = $null
  try {
    $probe = [IO.File]::Open($lockPath, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
    throw 'The claimed parent ownership transaction does not hold the exclusive per-user lock.'
  } catch [IO.IOException] {
    # FileShare.None in the parent is the expected proof that the transaction
    # which supplied the inherited 128-bit nonce still owns the lock.
  } finally {
    if ($probe) { $probe.Dispose() }
  }
}

function Protect-HermesExcelOwnerFile {
  [CmdletBinding()]
  param([Parameter(Mandatory)] [string] $Path)
  $sid = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
  & icacls.exe $Path /inheritance:r /grant:r "*$sid`:(F)" '*S-1-5-18:(F)' | Out-Null
  if ($LASTEXITCODE -ne 0) { throw "Could not apply a current-user/SYSTEM ACL to '$Path'." }
  $acl = Get-Acl -LiteralPath $Path
  if (-not $acl.AreAccessRulesProtected) { throw "ACL inheritance remains enabled for '$Path'." }
}

function Write-HermesExcelOwnerReceipt {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)] $Context,
    [ValidateRange(1, 65535)] [int] $BridgePort,
    [Parameter(Mandatory)] [string] $PluginVersion,
    [string] $ExpectedFingerprint
  )
  if ($PluginVersion -notmatch $script:SemVerPattern) { throw "PluginVersion '$PluginVersion' is not semantic versioning." }
  if ($ExpectedFingerprint -and $ExpectedFingerprint -notmatch '^sha256:[0-9a-f]{64}$') { throw 'ExpectedFingerprint is not a lowercase SHA-256 owner fingerprint.' }
  $parent = Split-Path -Parent $Context.ReceiptPath
  if (-not (Test-Path -LiteralPath $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
  $payload = [ordered]@{
    schema = 1
    profile_name = $Context.ProfileName
    profile_home = $Context.ProfileHome
    install_path = $Context.InstallPath
    bridge_port = $BridgePort
    plugin_version = $PluginVersion
    manifest_id = $script:ManifestId
  }
  $json = $payload | ConvertTo-Json -Compress
  $temp = Join-Path $parent ('.owner-v1.' + [guid]::NewGuid().ToString('N') + '.tmp')
  $backup = Join-Path $parent ('.owner-v1.' + [guid]::NewGuid().ToString('N') + '.bak')
  $hadPrevious = Test-Path -LiteralPath $Context.ReceiptPath
  $committed = $false
  try {
    [IO.File]::WriteAllText($temp, $json, [Text.UTF8Encoding]::new($false))
    Protect-HermesExcelOwnerFile -Path $temp
    $candidate = Read-HermesExcelOwnerReceipt -Path $temp
    if (-not (Test-HermesExcelOwnerMatch -Receipt $candidate -Context $Context)) { throw 'Candidate owner receipt failed identity verification.' }
    if ($ExpectedFingerprint -and (Get-HermesExcelOwnerFingerprint -Receipt $candidate) -ne $ExpectedFingerprint) {
      throw 'Candidate owner receipt fingerprint differs from the certified launcher identity.'
    }
    if ($hadPrevious) {
      [IO.File]::Replace($temp, $Context.ReceiptPath, $backup, $true)
    } else {
      [IO.File]::Move($temp, $Context.ReceiptPath)
    }
    $committed = $true
    Protect-HermesExcelOwnerFile -Path $Context.ReceiptPath
    $verified = Read-HermesExcelOwnerReceipt -Path $Context.ReceiptPath
    if (-not (Test-HermesExcelOwnerMatch -Receipt $verified -Context $Context)) { throw 'Committed owner receipt failed identity verification.' }
    if ($hadPrevious) { Remove-Item -LiteralPath $backup -Force -ErrorAction Stop }
    return $verified
  } catch {
    $commitError = $_.Exception.Message
    if ($committed -and $hadPrevious -and (Test-Path -LiteralPath $backup)) {
      $failedReplacement = Join-Path $parent ('.owner-v1.' + [guid]::NewGuid().ToString('N') + '.failed')
      try {
        [IO.File]::Replace($backup, $Context.ReceiptPath, $failedReplacement, $true)
        $committed = $false
        Remove-Item -LiteralPath $failedReplacement -Force -ErrorAction SilentlyContinue
      } catch {
        throw "Owner receipt commit failed ('$commitError') and the exact prior receipt could not be restored. Recovery receipt retained at '$backup': $($_.Exception.Message)"
      }
    } elseif ($committed -and -not $hadPrevious) {
      try { Remove-Item -LiteralPath $Context.ReceiptPath -Force -ErrorAction Stop; $committed = $false } catch {
        throw "Fresh owner receipt commit failed ('$commitError') and its safe same-owner receipt could not be removed: $($_.Exception.Message)"
      }
    }
    throw
  } finally {
    Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue
    # Never discard a receipt recovery backup on failure.
  }
}

function Restore-HermesExcelOwnerReceiptSnapshot {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)] $Context,
    [Parameter(Mandatory)] [string] $ExpectedCurrentFingerprint,
    [Parameter(Mandatory)] [bool] $HadPrevious,
    [byte[]] $PreviousBytes
  )
  if ($ExpectedCurrentFingerprint -notmatch '^sha256:[0-9a-f]{64}$') {
    throw 'ExpectedCurrentFingerprint is not a lowercase SHA-256 owner fingerprint.'
  }
  $current = Read-HermesExcelOwnerReceipt -Path $Context.ReceiptPath
  if (-not $current -or -not (Test-HermesExcelOwnerMatch -Receipt $current -Context $Context) -or
      (Get-HermesExcelOwnerFingerprint -Receipt $current) -ne $ExpectedCurrentFingerprint) {
    throw 'Refusing to restore an owner receipt snapshot because the current receipt is absent or differs from the failed install.'
  }

  if (-not $HadPrevious) {
    Remove-HermesExcelOwnerReceipt -Context $Context
    if (Test-Path -LiteralPath $Context.ReceiptPath) {
      throw 'Fresh-install owner receipt still exists after removal.'
    }
    return
  }
  if ($null -eq $PreviousBytes -or $PreviousBytes.Length -eq 0 -or $PreviousBytes.Length -gt $script:MaxReceiptBytes) {
    throw 'The previous owner receipt snapshot is absent or exceeds the receipt size limit.'
  }

  $parent = Split-Path -Parent $Context.ReceiptPath
  $temp = Join-Path $parent ('.owner-v1.' + [guid]::NewGuid().ToString('N') + '.restore')
  $backup = Join-Path $parent ('.owner-v1.' + [guid]::NewGuid().ToString('N') + '.failed-install')
  $replaced = $false
  try {
    [IO.File]::WriteAllBytes($temp, $PreviousBytes)
    Protect-HermesExcelOwnerFile -Path $temp
    $candidate = Read-HermesExcelOwnerReceipt -Path $temp
    if (-not (Test-HermesExcelOwnerMatch -Receipt $candidate -Context $Context)) {
      throw 'The previous owner receipt snapshot does not belong to this profile context.'
    }

    # File.Replace preserves an exact copy of the new, same-owner receipt in
    # $backup until the previous bytes have been parsed, ACL-protected, and
    # compared byte-for-byte. Any post-replace failure restores that safe receipt.
    [IO.File]::Replace($temp, $Context.ReceiptPath, $backup, $true)
    $replaced = $true
    Protect-HermesExcelOwnerFile -Path $Context.ReceiptPath
    $verified = Read-HermesExcelOwnerReceipt -Path $Context.ReceiptPath
    if (-not (Test-HermesExcelOwnerMatch -Receipt $verified -Context $Context)) {
      throw 'Restored owner receipt failed identity verification.'
    }
    $actualBytes = [IO.File]::ReadAllBytes($Context.ReceiptPath)
    if ($actualBytes.Length -ne $PreviousBytes.Length) {
      throw 'Restored owner receipt differs from the exact previous snapshot.'
    }
    for ($i = 0; $i -lt $actualBytes.Length; $i++) {
      if ($actualBytes[$i] -ne $PreviousBytes[$i]) {
        throw 'Restored owner receipt differs from the exact previous snapshot.'
      }
    }
    Remove-Item -LiteralPath $backup -Force -ErrorAction Stop
  } catch {
    $restoreError = $_.Exception.Message
    if ($replaced -and (Test-Path -LiteralPath $backup)) {
      $failedSnapshot = Join-Path $parent ('.owner-v1.' + [guid]::NewGuid().ToString('N') + '.failed-restore')
      try {
        [IO.File]::Replace($backup, $Context.ReceiptPath, $failedSnapshot, $true)
        Protect-HermesExcelOwnerFile -Path $Context.ReceiptPath
        Remove-Item -LiteralPath $failedSnapshot -Force -ErrorAction SilentlyContinue
      } catch {
        throw "Owner receipt snapshot restore failed ('$restoreError') and the safe failed-install receipt could not be restored. Recovery receipt retained at '$backup': $($_.Exception.Message)"
      }
    }
    throw
  } finally {
    Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue
    # Never discard the failed-install recovery receipt on a restore failure.
  }
}

function Remove-HermesExcelOwnerReceipt {
  [CmdletBinding()]
  param([Parameter(Mandatory)] $Context)
  $receipt = Read-HermesExcelOwnerReceipt -Path $Context.ReceiptPath
  if (-not $receipt) { return }
  if (-not (Test-HermesExcelOwnerMatch -Receipt $receipt -Context $Context)) {
    throw "Refusing to remove the receipt owned by profile '$($receipt.ProfileName)'."
  }
  Remove-Item -LiteralPath $Context.ReceiptPath -Force
}

Export-ModuleMember -Function @(
  'Resolve-HermesExcelProfileContext',
  'Read-HermesExcelOwnerReceipt',
  'Test-HermesExcelOwnerMatch',
  'Get-HermesExcelLegacyArtifacts',
  'Assert-HermesExcelLegacyAdoption',
  'Assert-HermesExcelOwnership',
  'Enter-HermesExcelTransaction',
  'Exit-HermesExcelTransaction',
  'Assert-HermesExcelParentTransaction',
  'Write-HermesExcelOwnerReceipt',
  'Restore-HermesExcelOwnerReceiptSnapshot',
  'Remove-HermesExcelOwnerReceipt',
  'Get-HermesExcelOwnerFingerprint'
)
