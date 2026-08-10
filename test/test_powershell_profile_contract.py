from __future__ import annotations

import ast
import base64
import json
import os
import pathlib
import re
import shutil
import subprocess
import sys
import unittest
import uuid


ROOT = pathlib.Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))
INSTALL = ROOT / "install"
OWNERSHIP_MODULE = INSTALL / "profile-ownership.psm1"
POWERSHELL = shutil.which("powershell.exe") or shutil.which("powershell")


def _function_source(path: pathlib.Path, function_name: str) -> str:
    source = path.read_text(encoding="utf-8")
    tree = ast.parse(source, filename=str(path))
    for node in tree.body:
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)) and node.name == function_name:
            segment = ast.get_source_segment(source, node)
            if segment is None:
                raise AssertionError(f"could not recover source for {function_name}")
            return segment
    raise AssertionError(f"{function_name} is missing from {path.name}")


@unittest.skipUnless(os.name == "nt" and POWERSHELL, "Windows PowerShell contract")
class PowerShellModuleContractTests(unittest.TestCase):
    maxDiff = None

    def setUp(self) -> None:
        fixture_root = ROOT / ".test-temp"
        fixture_root.mkdir(exist_ok=True)
        # Let the PowerShell child create the fixture hierarchy. Some managed
        # Windows sandboxes give Python-created temp directories a process-only
        # DACL that a nested powershell.exe cannot traverse.
        self.temp_path = fixture_root / (
            "powershell-profile-contract-" + uuid.uuid4().hex
        )
        self.addCleanup(self._cleanup_fixture)
        self.local_appdata = self.temp_path / "Local App Data"
        self.profile_home = self.local_appdata / "hermes" / "profiles" / "finance"
        self.env = os.environ.copy()
        self.env.update(
            {
                "HERMES_TEST_MODULE": str(OWNERSHIP_MODULE),
                "HERMES_TEST_INSTALLER": str(INSTALL / "addin-install.ps1"),
                "HERMES_TEST_LOCALAPPDATA": str(self.local_appdata),
                "HERMES_TEST_HOME": str(self.profile_home),
            }
        )

    def _cleanup_fixture(self) -> None:
        # The module deliberately protects receipt/lock ACLs. Restore inheritance
        # before Python recursively removes the isolated fixture tree.
        if self.temp_path.exists():
            subprocess.run(
                ["icacls.exe", str(self.temp_path), "/reset", "/T", "/C", "/Q"],
                text=True,
                capture_output=True,
                check=False,
            )
            subprocess.run(
                ["icacls.exe", str(self.temp_path), "/inheritance:e", "/T", "/C", "/Q"],
                text=True,
                capture_output=True,
                check=False,
            )
            shutil.rmtree(self.temp_path, ignore_errors=True)

    def run_module(self, body: str) -> object:
        script = rf"""
$ErrorActionPreference = 'Stop'
$env:LOCALAPPDATA = $env:HERMES_TEST_LOCALAPPDATA
$env:HERMES_HOME = $env:HERMES_TEST_HOME
$module = Import-Module -Name $env:HERMES_TEST_MODULE -Force -PassThru
# ACL behavior is covered by the installer checks. Keep these runtime tests on
# pure temporary filesystem semantics so managed CI/sandbox ACL policy cannot
# mask ownership, locking, or atomic-replacement failures.
& $module {{ function script:Protect-HermesExcelOwnerFile {{ param($Path) }} }}
{body}
"""
        completed = subprocess.run(
            [
                str(POWERSHELL),
                "-NoLogo",
                "-NoProfile",
                "-NonInteractive",
                "-ExecutionPolicy",
                "Bypass",
                "-Command",
                script,
            ],
            cwd=ROOT,
            env=self.env,
            text=True,
            capture_output=True,
            timeout=30,
            check=False,
        )
        self.assertEqual(
            completed.returncode,
            0,
            msg=f"PowerShell failed.\nSTDOUT:\n{completed.stdout}\nSTDERR:\n{completed.stderr}",
        )
        marker = "@@HERMES_RESULT@@"
        result_lines = [
            line[len(marker) :]
            for line in completed.stdout.splitlines()
            if line.startswith(marker)
        ]
        self.assertEqual(
            len(result_lines),
            1,
            msg=f"PowerShell emitted no unique result marker:\n{completed.stdout}",
        )
        return json.loads(result_lines[0])

    def test_missing_same_other_and_corrupt_receipt_states_fail_closed(self) -> None:
        result = self.run_module(
            r"""
function Test-Throws([scriptblock] $Action) {
  try { & $Action | Out-Null; return $false } catch { return $true }
}
$context = Resolve-HermesExcelProfileContext -ProfileName 'finance' -HermesHome $env:HERMES_HOME

# Isolate receipt semantics from this machine's real Task Scheduler/HKCU state.
& $module { function script:Get-HermesExcelLegacyArtifacts { param($Context) return @() } }
$missing = $null -eq (Assert-HermesExcelOwnership -Context $context -Operation Install)
$written = Write-HermesExcelOwnerReceipt -Context $context -BridgePort 8788 -PluginVersion '0.2.0'
$same = $null -ne (Assert-HermesExcelOwnership -Context $context -Operation Install)

$otherHome = Join-Path $env:LOCALAPPDATA 'hermes\profiles\personal'
$other = Resolve-HermesExcelProfileContext -ProfileName 'personal' -HermesHome $otherHome
$otherRefused = Test-Throws { Assert-HermesExcelOwnership -Context $other -Operation ReadOnly }

[IO.File]::WriteAllText($context.ReceiptPath, '{}', [Text.UTF8Encoding]::new($false))
$corruptRefused = Test-Throws { Assert-HermesExcelOwnership -Context $context -Operation ReadOnly }

$result = [ordered]@{
  missing_allowed = $missing
  same_allowed = $same
  other_refused = $otherRefused
  corrupt_refused = $corruptRefused
}
Write-Output ('@@HERMES_RESULT@@' + ($result | ConvertTo-Json -Compress))
"""
        )
        self.assertEqual(
            result,
            {
                "missing_allowed": True,
                "same_allowed": True,
                "other_refused": True,
                "corrupt_refused": True,
            },
        )

    def test_legacy_adoption_is_explicit_and_install_only(self) -> None:
        result = self.run_module(
            r"""
function Test-Throws([scriptblock] $Action) {
  try { & $Action | Out-Null; return $false } catch { return $true }
}
$context = Resolve-HermesExcelProfileContext -ProfileName 'finance' -HermesHome $env:HERMES_HOME
& $module { function script:Get-HermesExcelLegacyArtifacts { param($Context) return @('install:legacy-fixture') } }
$withoutFlag = Test-Throws { Assert-HermesExcelOwnership -Context $context -Operation Install }
$wrongOperation = Test-Throws { Assert-HermesExcelOwnership -Context $context -Operation Mutate -AdoptLegacy }
$result = [ordered]@{
  without_flag_refused = $withoutFlag
  wrong_operation_refused = $wrongOperation
}
Write-Output ('@@HERMES_RESULT@@' + ($result | ConvertTo-Json -Compress))
"""
        )
        self.assertEqual(
            result,
            {
                "without_flag_refused": True,
                "wrong_operation_refused": True,
            },
        )
        module_text = OWNERSHIP_MODULE.read_text(encoding="utf-8")
        self.assertRegex(
            module_text,
            r"(?is)\$Operation\s+-eq\s+['\"]Install['\"].*\$AdoptLegacy",
        )
        self.assertRegex(
            module_text,
            r"(?i)Assert-HermesExcelLegacyAdoption\s+-Context\s+\$Context",
        )
        self.assertIn("'Hermes Docling Serve.lnk'", module_text)
        self.assertIn("AdoptLegacy Docling Startup shortcut", module_text)

    def test_transaction_lock_is_exclusive_and_released_by_dispose(self) -> None:
        result = self.run_module(
            r"""
function Test-Throws([scriptblock] $Action) {
  try { & $Action | Out-Null; return $false } catch { return $true }
}
$context = Resolve-HermesExcelProfileContext -ProfileName 'finance' -HermesHome $env:HERMES_HOME
$first = Enter-HermesExcelTransaction -Context $context
try {
  $contended = Test-Throws { Enter-HermesExcelTransaction -Context $context }
} finally {
  Exit-HermesExcelTransaction -Lock $first
}
$second = Enter-HermesExcelTransaction -Context $context
$reacquired = $null -ne $second
Exit-HermesExcelTransaction -Lock $second
$result = [ordered]@{ contended = $contended; reacquired = $reacquired }
Write-Output ('@@HERMES_RESULT@@' + ($result | ConvertTo-Json -Compress))
"""
        )
        self.assertEqual(result, {"contended": True, "reacquired": True})

    def test_parent_transaction_proof_requires_the_live_exclusive_lock(self) -> None:
        result = self.run_module(
            r"""
function Test-Throws([scriptblock] $Action) {
  try { & $Action | Out-Null; return $false } catch { return $true }
}
$context = Resolve-HermesExcelProfileContext -ProfileName 'finance' -HermesHome $env:HERMES_HOME
$proof = [guid]::NewGuid().ToString('N')
$env:HERMES_EXCEL_TRANSACTION_PROOF = $proof
$env:HERMES_EXCEL_TRANSACTION_PROFILE = $context.ProfileName
$env:HERMES_EXCEL_TRANSACTION_HOME = $context.ProfileHome
$env:HERMES_EXCEL_TRANSACTION_RECEIPT = $context.ReceiptPath
$lock = Enter-HermesExcelTransaction -Context $context -Proof $proof
try {
  $accepted_while_held = -not (Test-Throws {
    Assert-HermesExcelParentTransaction -Context $context -Proof $proof
  })
} finally {
  Exit-HermesExcelTransaction -Lock $lock
}
$refused_after_release = Test-Throws {
  Assert-HermesExcelParentTransaction -Context $context -Proof $proof
}
$result = [ordered]@{
  accepted_while_held = $accepted_while_held
  refused_after_release = $refused_after_release
}
Write-Output ('@@HERMES_RESULT@@' + ($result | ConvertTo-Json -Compress))
"""
        )
        self.assertEqual(
            result,
            {"accepted_while_held": True, "refused_after_release": True},
        )

    def test_receipt_replacement_is_valid_and_leaves_no_partial_files(self) -> None:
        result = self.run_module(
            r"""
$context = Resolve-HermesExcelProfileContext -ProfileName 'finance' -HermesHome $env:HERMES_HOME
$first = Write-HermesExcelOwnerReceipt -Context $context -BridgePort 8788 -PluginVersion '0.2.0'
$second = Write-HermesExcelOwnerReceipt -Context $context -BridgePort 8790 -PluginVersion '0.2.1'
$readBack = Read-HermesExcelOwnerReceipt -Path $context.ReceiptPath
$parent = Split-Path -Parent $context.ReceiptPath
$partials = @(Get-ChildItem -LiteralPath $parent -Force -File | Where-Object {
  $_.Name -like '.owner-v1.*.tmp' -or $_.Name -like '.owner-v1.*.bak'
})
$result = [ordered]@{
  first_port = $first.BridgePort
  final_port = $readBack.BridgePort
  final_version = $readBack.PluginVersion
  partial_count = $partials.Count
}
Write-Output ('@@HERMES_RESULT@@' + ($result | ConvertTo-Json -Compress))
"""
        )
        self.assertEqual(
            result,
            {
                "first_port": 8788,
                "final_port": 8790,
                "final_version": "0.2.1",
                "partial_count": 0,
            },
        )
        module_text = OWNERSHIP_MODULE.read_text(encoding="utf-8")
        self.assertIn("[IO.File]::Replace", module_text)
        self.assertRegex(module_text, r"(?i)Read-HermesExcelOwnerReceipt\s+-Path\s+\$temp")

    def test_failed_post_replace_verification_restores_exact_prior_receipt(self) -> None:
        result = self.run_module(
            r"""
function Test-Throws([scriptblock] $Action) {
  try { & $Action | Out-Null; return $false } catch { return $true }
}
$context = Resolve-HermesExcelProfileContext -ProfileName 'finance' -HermesHome $env:HERMES_HOME
$first = Write-HermesExcelOwnerReceipt -Context $context -BridgePort 8788 -PluginVersion '0.2.0'
$before = [IO.File]::ReadAllBytes($context.ReceiptPath)
$env:HERMES_TEST_RECEIPT = $context.ReceiptPath
& $module {
  function script:Protect-HermesExcelOwnerFile {
    param($Path)
    if ([string]::Equals($Path, $env:HERMES_TEST_RECEIPT, [StringComparison]::OrdinalIgnoreCase)) {
      throw 'injected post-replace ACL verification failure'
    }
  }
}
$failed = Test-Throws {
  Write-HermesExcelOwnerReceipt -Context $context -BridgePort 8790 -PluginVersion '0.2.1'
}
$after = [IO.File]::ReadAllBytes($context.ReceiptPath)
$readBack = Read-HermesExcelOwnerReceipt -Path $context.ReceiptPath
$result = [ordered]@{
  failed = $failed
  exact_restore = [Convert]::ToBase64String($before) -eq [Convert]::ToBase64String($after)
  port = $readBack.BridgePort
  version = $readBack.PluginVersion
}
Write-Output ('@@HERMES_RESULT@@' + ($result | ConvertTo-Json -Compress))
"""
        )
        self.assertEqual(
            result,
            {
                "failed": True,
                "exact_restore": True,
                "port": 8788,
                "version": "0.2.0",
            },
        )

    def test_owner_fingerprint_matches_python_reference(self) -> None:
        result = self.run_module(
            r"""
$context = Resolve-HermesExcelProfileContext -ProfileName 'finance' -HermesHome $env:HERMES_HOME
$receipt = [pscustomobject]@{
  Schema = 1
  ProfileName = $context.ProfileName
  ProfileHome = $context.ProfileHome
  InstallPath = $context.InstallPath
  BridgePort = 8788
  PluginVersion = '0.2.0'
  ManifestId = '4fd4d435-7f9a-4d6d-9251-32f154f83a1f'
}
$result = [ordered]@{
  fingerprint = Get-HermesExcelOwnerFingerprint -Receipt $receipt
  receipt = [ordered]@{
    schema = $receipt.Schema
    profile_name = $receipt.ProfileName
    profile_home = $receipt.ProfileHome
    install_path = $receipt.InstallPath
    bridge_port = $receipt.BridgePort
    plugin_version = $receipt.PluginVersion
    manifest_id = $receipt.ManifestId
  }
}
Write-Output ('@@HERMES_RESULT@@' + ($result | ConvertTo-Json -Compress -Depth 3))
"""
        )
        import profile_ownership

        receipt = profile_ownership.parse_owner_receipt(result["receipt"])
        self.assertEqual(
            result["fingerprint"],
            profile_ownership.owner_fingerprint(receipt),
        )

    def test_failed_install_receipt_snapshot_restores_exact_bytes_or_removes_fresh(self) -> None:
        result = self.run_module(
            r"""
$context = Resolve-HermesExcelProfileContext -ProfileName 'finance' -HermesHome $env:HERMES_HOME
$old = Write-HermesExcelOwnerReceipt -Context $context -BridgePort 8788 -PluginVersion '0.2.0'
$oldBytes = [IO.File]::ReadAllBytes($context.ReceiptPath)
$new = Write-HermesExcelOwnerReceipt -Context $context -BridgePort 8790 -PluginVersion '0.2.1'
$newFingerprint = Get-HermesExcelOwnerFingerprint -Receipt $new
Restore-HermesExcelOwnerReceiptSnapshot -Context $context `
  -ExpectedCurrentFingerprint $newFingerprint -HadPrevious $true -PreviousBytes $oldBytes
$restoredBytes = [IO.File]::ReadAllBytes($context.ReceiptPath)
$restored = Read-HermesExcelOwnerReceipt -Path $context.ReceiptPath
$exactOld = [Convert]::ToBase64String($oldBytes) -eq [Convert]::ToBase64String($restoredBytes)

Remove-HermesExcelOwnerReceipt -Context $context
$fresh = Write-HermesExcelOwnerReceipt -Context $context -BridgePort 8791 -PluginVersion '0.2.2'
$freshFingerprint = Get-HermesExcelOwnerFingerprint -Receipt $fresh
Restore-HermesExcelOwnerReceiptSnapshot -Context $context `
  -ExpectedCurrentFingerprint $freshFingerprint -HadPrevious $false
$partials = @(Get-ChildItem -LiteralPath (Split-Path -Parent $context.ReceiptPath) -Force -File |
  Where-Object { $_.Name -like '.owner-v1.*' })
$result = [ordered]@{
  exact_old = $exactOld
  restored_port = $restored.BridgePort
  fresh_removed = -not (Test-Path -LiteralPath $context.ReceiptPath)
  partial_count = $partials.Count
}
Write-Output ('@@HERMES_RESULT@@' + ($result | ConvertTo-Json -Compress))
"""
        )
        self.assertEqual(
            result,
            {
                "exact_old": True,
                "restored_port": 8788,
                "fresh_removed": True,
                "partial_count": 0,
            },
        )

    def test_utf16_vbs_launcher_preserves_accented_and_cjk_profile_paths(self) -> None:
        result = self.run_module(
            r"""
$template = Join-Path (Split-Path -Parent $env:HERMES_TEST_MODULE) 'run-bridge.vbs.template'
$profileHome = Join-Path $env:HERMES_TEST_LOCALAPPDATA 'Hérmès\配置\profiles\finance'
$installDir = Join-Path $profileHome 'excel-addin'
$dataDir = Join-Path $installDir 'data'
$configPath = Join-Path $profileHome 'config.yaml'
$receiptPath = Join-Path $env:HERMES_TEST_LOCALAPPDATA 'Hérmès\共享\owner-v1.json'
$vbs = Get-Content -LiteralPath $template -Raw
$vbs = $vbs.Replace('__INSTALL_DIR__', $installDir)
$vbs = $vbs.Replace('__DATA_DIR__', $dataDir)
$vbs = $vbs.Replace('__HERMES_HOME__', $profileHome)
$vbs = $vbs.Replace('__HERMES_CONFIG__', $configPath)
$vbs = $vbs.Replace('__PROFILE_NAME__', 'finance')
$vbs = $vbs.Replace('__OWNER_RECEIPT__', $receiptPath)
$vbs = $vbs.Replace('__OWNER_FINGERPRINT__', ('sha256:' + ('a' * 64)))
$output = Join-Path $env:HERMES_TEST_LOCALAPPDATA 'launcher.vbs'
New-Item -ItemType Directory -Path (Split-Path -Parent $output) -Force | Out-Null
Set-Content -LiteralPath $output -Value $vbs -Encoding Unicode -Force
$bytes = [IO.File]::ReadAllBytes($output)
$result = [ordered]@{ bytes = [Convert]::ToBase64String($bytes) }
Write-Output ('@@HERMES_RESULT@@' + ($result | ConvertTo-Json -Compress))
"""
        )
        raw = base64.b64decode(result["bytes"])
        self.assertTrue(raw.startswith(b"\xff\xfe"), "VBS launcher must be UTF-16LE")
        rendered = raw[2:].decode("utf-16le")
        self.assertIn("Hérmès", rendered)
        self.assertIn("配置", rendered)
        self.assertIn("共享", rendered)
        self.assertNotIn("__HERMES_HOME__", rendered)

    def test_exact_legacy_wef_kind_and_file_bytes_restore_then_remove(self) -> None:
        result = self.run_module(
            r"""
$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile(
  $env:HERMES_TEST_INSTALLER, [ref]$tokens, [ref]$errors)
if ($errors.Count -ne 0) { throw 'Could not parse installer helpers for executable contract test.' }
foreach ($name in @('Test-ExactValue', 'Restore-LegacyWefSnapshot', 'Restore-FileSnapshot')) {
  $definition = $ast.Find({
    param($node)
    $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name
  }, $true)
  if (-not $definition) { throw "Missing installer helper: $name" }
  Invoke-Expression $definition.Extent.Text
}

$testId = [guid]::NewGuid().ToString('N')
$WefDevSubKey = "Software\HermesExcelSidecarTests\$testId"
$WefDevKey = "HKCU:\$WefDevSubKey"
$LegacyWefValueName = 'HermesExcelAddinCatalog'
$HadPreviousLegacyWef = $true
$PreviousLegacyWefKind = 'ExpandString'
$PreviousLegacyWefValue = '%TEMP%\Hérmès\配置'
try {
  Restore-LegacyWefSnapshot
  $key = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($WefDevSubKey, $false)
  try {
    $kindExact = [string]$key.GetValueKind($LegacyWefValueName) -ceq 'ExpandString'
    $raw = $key.GetValue($LegacyWefValueName, $null,
      [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
    $valueExact = $raw -ceq $PreviousLegacyWefValue
  } finally {
    if ($key) { $key.Dispose() }
  }
  $HadPreviousLegacyWef = $false
  Restore-LegacyWefSnapshot
  $verify = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($WefDevSubKey, $false)
  try {
    $wefAbsent = -not $verify -or (@($verify.GetValueNames()) -cnotcontains $LegacyWefValueName)
  } finally {
    if ($verify) { $verify.Dispose() }
  }

  $artifact = Join-Path $env:HERMES_TEST_LOCALAPPDATA 'catalog\hermes-excel-addin.xml'
  New-Item -ItemType Directory -Path (Split-Path -Parent $artifact) -Force | Out-Null
  [IO.File]::WriteAllBytes($artifact, [byte[]](9, 9, 9))
  $expected = [byte[]](0, 255, 1, 128, 13, 10)
  Restore-FileSnapshot -Path $artifact -HadPrevious $true -PreviousBytes $expected -Label 'test catalog'
  $fileExact = [Convert]::ToBase64String([IO.File]::ReadAllBytes($artifact)) -ceq `
    [Convert]::ToBase64String($expected)
  Restore-FileSnapshot -Path $artifact -HadPrevious $false -PreviousBytes $null -Label 'test catalog'
  $fileAbsent = -not (Test-Path -LiteralPath $artifact)

  $result = [ordered]@{
    kind_exact = $kindExact
    value_exact = $valueExact
    wef_absent = $wefAbsent
    file_exact = $fileExact
    file_absent = $fileAbsent
  }
  Write-Output ('@@HERMES_RESULT@@' + ($result | ConvertTo-Json -Compress))
} finally {
  try { [Microsoft.Win32.Registry]::CurrentUser.DeleteSubKeyTree($WefDevSubKey, $false) } catch { }
}
"""
        )
        self.assertEqual(
            result,
            {
                "kind_exact": True,
                "value_exact": True,
                "wef_absent": True,
                "file_exact": True,
                "file_absent": True,
            },
        )

    def test_skip_sideload_does_not_cross_mutation_boundary(self) -> None:
        result = self.run_module(
            r"""
$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile(
  $env:HERMES_TEST_INSTALLER, [ref]$tokens, [ref]$errors)
$definition = $ast.Find({
  param($node)
  $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
    $node.Name -eq 'Register-Sideload'
}, $true)
if (-not $definition) { throw 'Missing Register-Sideload helper.' }
Invoke-Expression $definition.Extent.Text

function Write-Step { param($Message) }
function Write-Warn2 { param($Message) }
$script:childCalls = 0
$script:flagSeenByChild = $false
function powershell {
  $script:childCalls++
  $script:flagSeenByChild = $script:SideloadMutationAttempted
  $global:LASTEXITCODE = 0
}
$InstallDir = Join-Path $env:HERMES_TEST_LOCALAPPDATA 'sideload-boundary'
New-Item -ItemType Directory -Path (Join-Path $InstallDir 'install') -Force | Out-Null
New-Item -ItemType File -Path (Join-Path $InstallDir 'install\register-sideload.ps1') -Force | Out-Null
$ProfileContext = [pscustomobject]@{
  ProfileName = 'finance'
  ProfileHome = $env:HERMES_HOME
  ReceiptPath = (Join-Path $env:HERMES_TEST_LOCALAPPDATA 'owner-v1.json')
}
$Port = 8788
$TransactionProof = 'fixture-proof'
$script:SideloadMutationAttempted = $false
$SkipSideload = $true
Register-Sideload
$skipPreserved = -not $script:SideloadMutationAttempted -and $script:childCalls -eq 0

$SkipSideload = $false
Register-Sideload
$attemptObserved = $script:SideloadMutationAttempted -and $script:flagSeenByChild -and $script:childCalls -eq 1
$result = [ordered]@{ skip_preserved = $skipPreserved; attempt_observed = $attemptObserved }
Write-Output ('@@HERMES_RESULT@@' + ($result | ConvertTo-Json -Compress))
"""
        )
        self.assertEqual(
            result,
            {"skip_preserved": True, "attempt_observed": True},
        )


class PowerShellEntrypointStaticContracts(unittest.TestCase):
    ENTRYPOINTS = (
        "apply.ps1",
        "addin-install.ps1",
        "check.ps1",
        "deploy.ps1",
        "register-sideload.ps1",
        "register-task.ps1",
        "rollback.ps1",
    )

    def test_every_entrypoint_accepts_explicit_profile_identity(self) -> None:
        declarations = {
            "ProfileName": r"(?im)\[\s*string\s*\]\s*\$ProfileName\b",
            "HermesHome": r"(?im)\[\s*string\s*\]\s*\$HermesHome\b",
            "OwnerReceiptPath": r"(?im)\[\s*string\s*\]\s*\$OwnerReceiptPath\b",
        }
        for name in self.ENTRYPOINTS:
            text = (INSTALL / name).read_text(encoding="utf-8")
            for parameter, pattern in declarations.items():
                with self.subTest(script=name, parameter=parameter):
                    self.assertRegex(text, pattern)

    def test_no_entrypoint_or_cli_exposes_live_takeover(self) -> None:
        for name in self.ENTRYPOINTS:
            text = (INSTALL / name).read_text(encoding="utf-8")
            with self.subTest(script=name):
                self.assertNotRegex(text, r"(?im)^\s*(?:\[[^\]]+\]\s*)*\$Takeover\b")
        cli_text = (ROOT / "cli.py").read_text(encoding="utf-8")
        self.assertNotIn("--takeover", cli_text.lower())

    def test_installer_guards_owner_before_singleton_mutation_or_token_access(self) -> None:
        text = (INSTALL / "addin-install.ps1").read_text(encoding="utf-8")
        guard = re.search(r"(?i)Assert-HermesExcelOwnership\s+-Context\b", text)
        self.assertIsNotNone(guard, "installer must perform the ownership assertion")
        assert guard is not None
        sensitive = (
            r"GetEnvironmentVariable\(\s*['\"]HERMES_EXCEL_",
            r"Stop-ScheduledTask\b",
            r"Unregister-ScheduledTask\b",
            r"Stop-ExistingBridge\b",
            r"Copy-Payload\b",
            r"Ensure-BridgeToken\b",
            r"Ensure-IngestToken\b",
            r"hermes\s+config\s+set\b",
            r"hermes\s+gateway\s+restart\b",
        )
        for pattern in sensitive:
            matches = list(re.finditer(pattern, text, re.IGNORECASE))
            protected_matches = [match for match in matches if match.start() > guard.start()]
            with self.subTest(action=pattern):
                self.assertTrue(
                    protected_matches,
                    f"expected an owner-guarded action matching: {pattern}",
                )

    def test_rollback_guards_first_and_removes_receipt_only_after_owned_cleanup(self) -> None:
        text = (INSTALL / "rollback.ps1").read_text(encoding="utf-8")
        guard = re.search(r"(?i)Assert-HermesExcelOwnership\s+-Context\b", text)
        remove_receipt = re.search(r"(?i)Remove-HermesExcelOwnerReceipt\s+-Context\b", text)
        self.assertIsNotNone(guard, "rollback must assert ownership")
        self.assertIsNotNone(remove_receipt, "rollback must retire its owner receipt")
        assert guard is not None and remove_receipt is not None

        protected_actions = (
            r"Stop-ScheduledTask\b",
            r"Unregister-ScheduledTask\b",
            r"Stop-OwnedBridgeProcesses\b",
            r"powershell[^\r\n]*-File[^\r\n]*\$regSide[^\r\n]*-Unregister",
            r"Remove-Item[^\r\n]*\$InstallDir",
        )
        positions: list[int] = []
        for pattern in protected_actions:
            matches = [
                match
                for match in re.finditer(pattern, text, re.IGNORECASE)
                if match.start() > guard.start()
            ]
            with self.subTest(action=pattern):
                self.assertTrue(matches, f"expected rollback action missing: {pattern}")
            positions.extend(match.start() for match in matches)
        self.assertLess(guard.start(), min(positions))
        self.assertGreater(
            remove_receipt.start(),
            max(positions),
            "the receipt remains authoritative until owned singleton cleanup completes",
        )

    def test_apply_threads_profile_identity_and_adoption_to_installer(self) -> None:
        text = (INSTALL / "apply.ps1").read_text(encoding="utf-8")
        context_properties = {
            "ProfileName": "ProfileName",
            "HermesHome": "ProfileHome",
            "OwnerReceiptPath": "ReceiptPath",
        }
        for argument, property_name in context_properties.items():
            self.assertRegex(
                text,
                rf"(?i)['\"]?{argument}['\"]?\s*=\s*\$context\.{property_name}\b",
            )
        self.assertRegex(text, r"(?i)\$AdoptLegacy\b")
        self.assertRegex(
            text,
            r"(?i)['\"]?AdoptLegacy['\"]?\s*\]?\s*=\s*\$true\b",
        )

    def test_launcher_and_installer_certify_profile_owner_identity(self) -> None:
        template = (INSTALL / "run-bridge.cmd.template").read_text(encoding="utf-8")
        installer = (INSTALL / "addin-install.ps1").read_text(encoding="utf-8")
        broker = (ROOT / "broker" / "server.mjs").read_text(encoding="utf-8")

        self.assertIn("HERMES_EXCEL_PROFILE_NAME=__PROFILE_NAME__", template)
        self.assertIn(
            "HERMES_EXCEL_OWNER_FINGERPRINT=__OWNER_FINGERPRINT__",
            template,
        )
        self.assertRegex(installer, r"\$resp\.profile_name\s+-eq\s+\$ProfileContext\.ProfileName")
        self.assertRegex(installer, r"\$resp\.owner_fingerprint\s+-eq\s+\$OwnerFingerprint")
        self.assertRegex(broker, r"profile_name:\s*runtimeOwnerIdentity\.profile_name")
        self.assertRegex(
            broker,
            r"owner_fingerprint:\s*runtimeOwnerIdentity\.owner_fingerprint",
        )

    def test_installer_commits_receipt_before_gateway_and_attests_adapter_owner(self) -> None:
        text = (INSTALL / "addin-install.ps1").read_text(encoding="utf-8")
        receipt_commit = text.index("[void](Write-HermesExcelOwnerReceipt")
        gateway_restart = text.index("& hermes gateway restart", receipt_commit)
        self.assertLess(receipt_commit, gateway_restart)
        self.assertIn("$ReceiptCommitted = $true", text[receipt_commit:gateway_restart])
        self.assertRegex(text, r"\$health\.profile_name\s+-eq\s+\$ProfileContext\.ProfileName")
        self.assertRegex(text, r"\$health\.owner_fingerprint\s+-eq\s+\$OwnerFingerprint")
        self.assertRegex(text, r"\[int\]\$health\.bridge_port\s+-eq\s+\$Port")
        restore = text.index("Restore-HermesExcelOwnerReceiptSnapshot")
        quiescent = text.rfind("Assert-OwnedBridgeQuiescent", 0, restore)
        self.assertGreater(quiescent, gateway_restart)
        final_gateway_restart = text.index("& hermes gateway restart | Out-Null", restore)
        self.assertLess(restore, final_gateway_restart)
        self.assertIn("$ReceiptCommitted -and -not $RestoreFailed", text)

    def test_failed_install_cleanup_checks_native_failures(self) -> None:
        text = (INSTALL / "addin-install.ps1").read_text(encoding="utf-8")
        self.assertRegex(
            text,
            r"\$HadPreviousSideload\s*=\s*\$HadPreviousLegacyWef\s+-or\s+\$HadPreviousCatalog",
        )
        self.assertRegex(
            text,
            r"(?is)register-sideload\.ps1.*?\$cleanupSideloadExit\s*=\s*\$LASTEXITCODE.*?"
            r"if\s*\(\$cleanupSideloadExit\s+-ne\s+0\).*?throw",
        )
        self.assertRegex(
            text,
            r"(?is)hermes\s+config\s+set\s+excel\.enabled.*?"
            r"\$configRestoreExit\s*=\s*\$LASTEXITCODE.*?\$RestoreFailed\s*=\s*\$true",
        )
        self.assertRegex(
            text,
            r"(?is)if\s*\(\$GatewayRestartAttempted\b.*?hermes\s+gateway\s+restart.*?"
            r"\$gatewayRestoreExit\s*=\s*\$LASTEXITCODE.*?\$RestoreFailed\s*=\s*\$true",
        )

    def test_profile_launchers_are_unicode_safe_without_unicode_cmd_literals(self) -> None:
        installer = (INSTALL / "addin-install.ps1").read_text(encoding="utf-8")
        cmd_template = (INSTALL / "run-bridge.cmd.template").read_text(encoding="ascii")
        vbs_template = (INSTALL / "run-bridge.vbs.template").read_text(encoding="utf-8")
        register_task = (INSTALL / "register-task.ps1").read_text(encoding="utf-8")

        for placeholder in (
            "__INSTALL_DIR__",
            "__DATA_DIR__",
            "__HERMES_HOME__",
            "__HERMES_CONFIG__",
            "__OWNER_RECEIPT__",
        ):
            self.assertNotIn(f"$cmd.Replace('{placeholder}'", installer)
            self.assertNotIn(placeholder, cmd_template)
            self.assertIn(placeholder, vbs_template)
        self.assertRegex(
            installer,
            r"Set-Content[^\r\n]+run-bridge\.vbs[^\r\n]+-Encoding\s+Unicode",
        )
        self.assertRegex(
            register_task,
            r"Set-Content[^\r\n]+\$ServiceVbs[^\r\n]+-Encoding\s+Unicode",
        )
        self.assertIn('%~dp0', cmd_template)

    def test_rollback_stops_exact_supervisor_chain_and_proves_port_quiescence(self) -> None:
        text = (INSTALL / "rollback.ps1").read_text(encoding="utf-8")
        for executable, target in (
            ("wscript.exe", "service\\bridge-service.vbs"),
            ("cmd.exe", "service\\bridge-service.cmd"),
            ("node.exe", "broker\\server.mjs"),
        ):
            self.assertIn(f"'{executable}'", text)
            self.assertIn(target, text)
        self.assertIn("Get-CimInstance Win32_Process -ErrorAction Stop", text)
        self.assertIn("Stop-Process -Id $process.ProcessId -Force -ErrorAction Stop", text)
        self.assertIn("GetActiveTcpListeners()", text)
        remove_receipt = text.index("Remove-HermesExcelOwnerReceipt -Context")
        final_quiescence = text.rfind("Assert-OwnedBridgeQuiescent", 0, remove_receipt)
        self.assertGreater(final_quiescence, text.index("& hermes gateway restart"))
        self.assertNotRegex(text, r"(?i)Name='node\.exe'.*?-like\s+\"\*\$idir\*\"")

    def test_receipt_gates_verified_shortcut_and_inline_wef_cleanup(self) -> None:
        rollback = (INSTALL / "rollback.ps1").read_text(encoding="utf-8")
        installer = (INSTALL / "addin-install.ps1").read_text(encoding="utf-8")
        guard = rollback.index("Assert-HermesExcelOwnership -Context")
        for shortcut in ("Hermes Excel Bridge.lnk", "Hermes Docling Serve.lnk"):
            position = rollback.index(shortcut)
            self.assertGreater(position, guard)
        self.assertRegex(
            rollback,
            r"Remove-Item\s+-LiteralPath\s+\$lnk\s+-Force\s+-ErrorAction\s+Stop",
        )
        self.assertRegex(
            rollback,
            r"Remove-Item\s+-LiteralPath\s+\$doclingLnk\s+-Force\s+-ErrorAction\s+Stop",
        )
        self.assertRegex(
            rollback,
            r"(?is)Remove-ItemProperty.*?HermesExcelAddinCatalog.*?"
            r"Get-ItemProperty.*?HermesExcelAddinCatalog.*?throw",
        )
        self.assertIn("$HadPreviousDoclingShortcut", installer)
        self.assertIn(
            "Restore-FileSnapshot -Path $DoclingShortcutPath -HadPrevious $HadPreviousDoclingShortcut",
            installer,
        )
        self.assertRegex(
            installer,
            r"Remove-Item\s+-LiteralPath\s+\$Path\s+-Force\s+-ErrorAction\s+Stop",
        )

    def test_failed_install_restores_exact_prior_excel_shortcut_snapshot(self) -> None:
        installer = (INSTALL / "addin-install.ps1").read_text(encoding="utf-8")
        capture = installer.index(
            "$PreviousExcelShortcutBytes = [IO.File]::ReadAllBytes($ExcelShortcutPath)"
        )
        mutation = installer.index("$InstallMutationStarted = $true")
        restore = installer.index("Restore-FileSnapshot -Path $ExcelShortcutPath", mutation)
        self.assertLess(capture, mutation)
        self.assertGreater(restore, mutation)
        self.assertIn("[IO.File]::Replace($temp, $Path, $backup, $true)", installer)
        self.assertIn("[Convert]::ToBase64String($PreviousBytes)", installer)
        docling_capture = installer.index(
            "$PreviousDoclingShortcutBytes = [IO.File]::ReadAllBytes($DoclingShortcutPath)"
        )
        docling_restore = installer.index(
            "Restore-FileSnapshot -Path $DoclingShortcutPath", mutation
        )
        self.assertLess(docling_capture, mutation)
        self.assertGreater(docling_restore, mutation)

    def test_existing_task_is_retired_fail_closed_before_payload_mutation(self) -> None:
        installer = (INSTALL / "addin-install.ps1").read_text(encoding="utf-8")
        start = installer.index("if ($HadPreviousTask) {", installer.index("$InstallMutationStarted"))
        end = installer.index("Stop-ExistingBridge", start)
        block = installer[start:end]
        self.assertIn("Stop-ScheduledTask", block)
        self.assertIn("-ErrorAction Stop", block)
        self.assertIn(
            "Unregister-ScheduledTask -TaskName 'Hermes_Excel_Bridge' -Confirm:$false -ErrorAction Stop",
            block,
        )
        self.assertIn("Get-ScheduledTask -TaskName 'Hermes_Excel_Bridge'", block)
        self.assertIn("survived pre-install unregistration", block)
        self.assertNotIn("Stop-ScheduledTask -TaskName 'Hermes_Excel_Bridge' -ErrorAction SilentlyContinue", block)

    def test_sideload_unregister_verifies_catalog_and_legacy_value_absence(self) -> None:
        text = (INSTALL / "register-sideload.ps1").read_text(encoding="utf-8")
        unregister = text.index("if ($Unregister)")
        success = text.index("development manifest unregistered", unregister)
        block = text[unregister:success]
        self.assertIn("Remove-Item -LiteralPath $CatalogPath -Force -ErrorAction Stop", block)
        self.assertIn("Test-Path -LiteralPath $CatalogPath", block)
        self.assertIn("Get-ItemProperty -LiteralPath $LegacyKey -Name $LegacyValue", block)
        self.assertRegex(
            text,
            r"Remove-ItemProperty\s+-LiteralPath\s+\$LegacyKey\s+-Name\s+\$LegacyValue\s+"
            r"-Force\s+-ErrorAction\s+Stop",
        )

    def test_failed_upgrade_restores_exact_modern_and_legacy_sideload_table(self) -> None:
        installer = (INSTALL / "addin-install.ps1").read_text(encoding="utf-8")
        helper = (INSTALL / "register-sideload.ps1").read_text(encoding="utf-8")
        mutation = installer.index("$InstallMutationStarted = $true")
        self.assertLess(
            installer.index("[Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames"),
            mutation,
        )
        self.assertLess(
            installer.index("$PreviousLegacyWefKind = [string]$previousWefKey.GetValueKind"),
            mutation,
        )
        self.assertLess(
            installer.index("$PreviousCatalogBytes = [IO.File]::ReadAllBytes($CatalogPath)"),
            mutation,
        )

        restore_start = installer.index("function Restore-SideloadSnapshot")
        restore_end = installer.index("function Register-Autostart", restore_start)
        restore = installer[restore_start:restore_end]
        modern_if = restore.index("if ($HadPreviousCatalog)")
        modern_register = restore.index("-RegisterExistingCatalog", modern_if)
        legacy_restore = restore.index("Restore-LegacyWefSnapshot", modern_register)
        self.assertLess(modern_if, modern_register)
        self.assertLess(modern_register, legacy_restore)
        self.assertGreaterEqual(
            restore.count(
                "Restore-FileSnapshot -Path $CatalogPath -HadPrevious $true"
            ),
            2,
            "catalog bytes must be restored both before and after registration",
        )
        self.assertIn("-HadPrevious $false", restore)

        mode_start = helper.index("if ($RegisterExistingCatalog)")
        mode_end = helper.index("if (-not (Test-Path -LiteralPath $SourceManifest))", mode_start)
        existing_mode = helper[mode_start:mode_end]
        self.assertIn("Invoke-DevSettings @('register', $CatalogPath)", existing_mode)
        self.assertNotIn("Remove-LegacyRegistration", existing_mode)
        self.assertNotIn("Set-Content", existing_mode)

        cleanup = installer[installer.index("# Clear the possibly-mutated registration"):]
        self.assertIn("-Unregister", cleanup)
        self.assertIn("Restore-SideloadSnapshot", cleanup)

    def test_sideload_compensation_runs_only_after_mutation_attempt(self) -> None:
        installer = (INSTALL / "addin-install.ps1").read_text(encoding="utf-8")
        self.assertIn("$SideloadMutationAttempted = $false", installer)
        register_start = installer.index("function Register-Sideload")
        register_end = installer.index("function Invoke-HealthCheck", register_start)
        register = installer[register_start:register_end]
        skip = register.index("if ($SkipSideload)")
        boundary = register.index("$script:SideloadMutationAttempted = $true")
        child = register.index("& powershell", boundary)
        self.assertLess(skip, boundary)
        self.assertLess(boundary, child)

        cleanup_start = installer.index("if ($SideloadMutationAttempted) {", child)
        cleanup_child = installer.index("-Unregister", cleanup_start)
        cleanup_end = installer.index("} catch {", cleanup_start)
        self.assertLess(cleanup_start, cleanup_child)
        self.assertLess(cleanup_child, cleanup_end)
        self.assertIn(
            "if ($SideloadMutationAttempted -and -not $RestoreFailed)",
            installer,
        )

    def test_tokens_are_profile_files_not_successful_user_scope_secrets(self) -> None:
        text = (INSTALL / "addin-install.ps1").read_text(encoding="utf-8")
        self.assertNotRegex(
            text,
            r"SetEnvironmentVariable\(\s*['\"]HERMES_EXCEL_(?:BRIDGE|INGEST)_TOKEN['\"]\s*,\s*\$tok",
        )
        guard = text.index("Assert-HermesExcelOwnership -Context")
        clear_bridge = text.index(
            "SetEnvironmentVariable('HERMES_EXCEL_BRIDGE_TOKEN', $null, 'User')"
        )
        clear_ingest = text.index(
            "SetEnvironmentVariable('HERMES_EXCEL_INGEST_TOKEN', $null, 'User')"
        )
        self.assertGreater(clear_bridge, guard)
        self.assertGreater(clear_ingest, guard)

    def test_rollback_cancellation_cannot_retire_the_owner_receipt(self) -> None:
        text = (INSTALL / "rollback.ps1").read_text(encoding="utf-8")
        cancellation = re.search(
            r"(?is)if\s*\(\s*\$proceed\s*\).*?else\s*\{(?P<body>.*?)\}",
            text,
        )
        self.assertIsNotNone(cancellation)
        assert cancellation is not None
        self.assertRegex(cancellation.group("body"), r"(?i)throw\b.*receipt.*retain")


class PythonStatusStaticContracts(unittest.TestCase):
    def test_status_uses_owned_token_file_and_verified_https(self) -> None:
        cli_path = ROOT / "cli.py"
        status = _function_source(cli_path, "_status")
        token_reader = _function_source(cli_path, "_read_bridge_token")

        self.assertIn("https://localhost:", status)
        self.assertNotIn("http://127.0.0.1:", status)
        self.assertIn("_local_tls_context()", status)
        self.assertIn("_read_bridge_token(context)", status)
        self.assertNotRegex(status, r"os\.environ\b.*HERMES_EXCEL_BRIDGE_TOKEN")
        self.assertLess(status.index("owner_matches"), status.index("_read_bridge_token(context)"))
        self.assertIn("context.data_path", token_reader)
        self.assertIn(".bridge-token", token_reader)
        self.assertNotIn("HERMES_EXCEL_BRIDGE_TOKEN", token_reader)


if __name__ == "__main__":
    unittest.main()
