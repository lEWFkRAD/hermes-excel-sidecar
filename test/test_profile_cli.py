from __future__ import annotations

import argparse
import contextlib
import importlib.util
import io
import json
import os
import pathlib
import sys
import tempfile
import types
import unittest
from unittest import mock

ROOT = pathlib.Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location("excel_profile_cli_bare", ROOT / "cli.py")
profile_cli = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = profile_cli
assert SPEC.loader
SPEC.loader.exec_module(profile_cli)


class ProfileCliBindingTests(unittest.TestCase):
    def test_registration_carries_the_selected_profile_without_global_state(self):
        parser = argparse.ArgumentParser()
        profile_cli.register_cli(parser, profile_name="finance")

        first = parser.parse_args(["status"])
        self.assertEqual(first.excel_profile_name, "finance")

        other_parser = argparse.ArgumentParser()
        profile_cli.register_cli(other_parser, profile_name="personal")
        second = other_parser.parse_args(["check"])

        self.assertEqual(second.excel_profile_name, "personal")
        self.assertEqual(first.excel_profile_name, "finance")

    def test_runtime_environment_uses_call_time_hermes_profile_override(self):
        first_home = pathlib.Path("C:/Hermes Profiles/finance")
        second_home = pathlib.Path("C:/Hermes Profiles/personal")
        homes = iter((first_home, second_home))
        constants = types.ModuleType("hermes_constants")
        constants.get_hermes_home = lambda: next(homes)

        with mock.patch.dict(sys.modules, {"hermes_constants": constants}), mock.patch.dict(
            os.environ, {"HERMES_HOME": "C:/stale-process-home"}, clear=False
        ):
            first = profile_cli._runtime_environment()
            second = profile_cli._runtime_environment()

        self.assertEqual(first["HERMES_HOME"], str(first_home))
        self.assertEqual(second["HERMES_HOME"], str(second_home))


class ProfileCliInvocationTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory(prefix="Hermes Profile With Spaces ")
        self.addCleanup(self.temp_dir.cleanup)
        self.local_appdata = pathlib.Path(self.temp_dir.name)
        # Profile names cannot contain spaces, but every path component may.
        self.profile_home = self.local_appdata / "hermes" / "profiles" / "finance"
        self.runtime_env = {
            "LOCALAPPDATA": str(self.local_appdata),
            "HERMES_HOME": str(self.profile_home),
            "HERMES_EXCEL_BRIDGE_TOKEN": "test-token-must-not-be-logged",
        }
        self.context = profile_cli.profile_ownership.resolve_profile_context(
            "finance", environ=self.runtime_env
        )

    def _parser(self):
        parser = argparse.ArgumentParser()
        profile_cli.register_cli(parser, profile_name="finance")
        return parser

    def test_every_powershell_command_receives_explicit_profile_paths_and_env(self):
        cases = (
            (["install", "--port", "9123"], "apply.ps1", ["-Port", "9123"]),
            (["check"], "check.ps1", []),
            (["rollback"], "rollback.ps1", []),
        )
        for command_args, expected_script, extra in cases:
            with self.subTest(command=command_args), mock.patch.object(
                profile_cli, "_runtime_environment", return_value=self.runtime_env
            ), mock.patch.object(profile_cli.sys, "platform", "win32"), mock.patch.object(
                profile_cli.subprocess,
                "run",
                return_value=types.SimpleNamespace(returncode=29),
            ) as run:
                result = profile_cli.excel_sidecar_command(
                    self._parser().parse_args(command_args)
                )

            self.assertEqual(result, 29)
            argv = run.call_args.args[0]
            self.assertEqual(pathlib.Path(argv[argv.index("-File") + 1]).name, expected_script)
            self.assertEqual(argv[argv.index("-ProfileName") + 1], "finance")
            self.assertEqual(
                argv[argv.index("-HermesHome") + 1], str(self.context.profile_home)
            )
            self.assertEqual(
                argv[argv.index("-OwnerReceiptPath") + 1],
                str(self.context.receipt_path),
            )
            self.assertEqual(argv[-len(extra) :] if extra else [], extra)
            self.assertEqual(
                run.call_args.kwargs["env"]["HERMES_HOME"],
                str(self.context.profile_home),
            )
            self.assertFalse(run.call_args.kwargs["check"])

    def test_legacy_adoption_is_install_only_and_explicit(self):
        parser = self._parser()
        with mock.patch.object(
            profile_cli, "_runtime_environment", return_value=self.runtime_env
        ), mock.patch.object(profile_cli.sys, "platform", "win32"), mock.patch.object(
            profile_cli.subprocess,
            "run",
            return_value=types.SimpleNamespace(returncode=0),
        ) as run:
            self.assertEqual(
                profile_cli.excel_sidecar_command(
                    parser.parse_args(["install", "--adopt-legacy"])
                ),
                0,
            )

        self.assertIn("-AdoptLegacy", run.call_args.args[0])
        self.assertFalse(self._parser().parse_args(["install"]).adopt_legacy)

    def test_invalid_profile_selection_returns_usage_error_without_spawning(self):
        invalid_env = dict(self.runtime_env, HERMES_HOME=str(self.profile_home))
        with mock.patch.object(
            profile_cli, "_runtime_environment", return_value=invalid_env
        ), mock.patch.object(profile_cli.subprocess, "run") as run, contextlib.redirect_stderr(
            io.StringIO()
        ):
            args = self._parser().parse_args(["check"])
            args.excel_profile_name = "personal"
            result = profile_cli.excel_sidecar_command(args)

        self.assertEqual(result, 2)
        run.assert_not_called()


class ProfileCliStatusTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory(prefix="Hermes Status With Spaces ")
        self.addCleanup(self.temp_dir.cleanup)
        local_appdata = pathlib.Path(self.temp_dir.name)
        self.environ = {
            "LOCALAPPDATA": str(local_appdata),
            "HERMES_HOME": str(local_appdata / "hermes" / "profiles" / "finance"),
        }
        self.context = profile_cli.profile_ownership.resolve_profile_context(
            "finance", environ=self.environ
        )
        self.receipt = profile_cli.profile_ownership.OwnerReceipt.from_context(
            self.context, bridge_port=8788, plugin_version="0.1.0"
        )

    def test_status_rejects_wrong_owner_before_reading_token(self):
        other_env = dict(
            self.environ,
            HERMES_HOME=str(pathlib.Path(self.temp_dir.name) / "hermes" / "profiles" / "personal"),
        )
        other_context = profile_cli.profile_ownership.resolve_profile_context(
            "personal", environ=other_env
        )
        other_receipt = profile_cli.profile_ownership.OwnerReceipt.from_context(
            other_context, bridge_port=8788, plugin_version="0.1.0"
        )
        with mock.patch.object(
            profile_cli.profile_ownership,
            "load_owner_receipt",
            return_value=other_receipt,
        ), mock.patch.object(profile_cli, "_read_bridge_token") as read_token, mock.patch.object(
            profile_cli, "urlopen"
        ) as urlopen, contextlib.redirect_stderr(
            io.StringIO()
        ):
            result = profile_cli._status(self.context, None)

        self.assertEqual(result, 1)
        read_token.assert_not_called()
        urlopen.assert_not_called()

    def test_status_rejects_unowned_port_before_reading_token(self):
        with mock.patch.object(
            profile_cli.profile_ownership,
            "load_owner_receipt",
            return_value=self.receipt,
        ), mock.patch.object(profile_cli, "_read_bridge_token") as read_token, mock.patch.object(
            profile_cli, "urlopen"
        ) as urlopen, contextlib.redirect_stderr(
            io.StringIO()
        ):
            result = profile_cli._status(self.context, 9999)

        self.assertEqual(result, 1)
        read_token.assert_not_called()
        urlopen.assert_not_called()

    def test_status_uses_owned_port_https_and_verified_tls(self):
        tls_context = object()
        token = "ab" * 32
        fingerprint = profile_cli.profile_ownership.owner_fingerprint(self.receipt)
        response = io.StringIO(
            json.dumps(
                {
                    "service": "hermes-excel-bridge",
                    "port": 8788,
                    "profile_name": "finance",
                    "owner_fingerprint": fingerprint,
                    "debug": token,
                }
            )
        )
        opener = mock.MagicMock()
        opener.return_value.__enter__.return_value = response
        stdout = io.StringIO()
        stderr = io.StringIO()
        self.context.data_path.mkdir(parents=True)
        (self.context.data_path / ".bridge-token").write_text(token, encoding="ascii")

        with mock.patch.object(
            profile_cli.profile_ownership,
            "load_owner_receipt",
            return_value=self.receipt,
        ), mock.patch.object(
            profile_cli, "_local_tls_context", return_value=tls_context
        ), mock.patch.object(profile_cli, "urlopen", opener), contextlib.redirect_stdout(
            stdout
        ), contextlib.redirect_stderr(stderr):
            result = profile_cli._status(self.context, None)

        self.assertEqual(result, 0)
        request = opener.call_args.args[0]
        self.assertEqual(request.full_url, "https://localhost:8788/api/health")
        self.assertEqual(request.get_header("X-hermes-token"), token)
        self.assertEqual(opener.call_args.kwargs["timeout"], 3)
        self.assertIs(opener.call_args.kwargs["context"], tls_context)
        self.assertNotIn(token, stdout.getvalue())
        self.assertNotIn(token, stderr.getvalue())

    def test_status_rejects_mismatched_profile_attestation_without_details(self):
        token = "cd" * 32
        expected_fingerprint = profile_cli.profile_ownership.owner_fingerprint(
            self.receipt
        )
        self.context.data_path.mkdir(parents=True)
        (self.context.data_path / ".bridge-token").write_text(token, encoding="ascii")
        response = io.StringIO(
            json.dumps(
                {
                    "service": "hermes-excel-bridge",
                    "port": 8788,
                    "profile_name": "personal",
                    "owner_fingerprint": expected_fingerprint,
                }
            )
        )
        opener = mock.MagicMock()
        opener.return_value.__enter__.return_value = response
        stdout = io.StringIO()
        stderr = io.StringIO()

        with mock.patch.object(
            profile_cli.profile_ownership,
            "load_owner_receipt",
            return_value=self.receipt,
        ), mock.patch.object(profile_cli, "urlopen", opener), contextlib.redirect_stdout(
            stdout
        ), contextlib.redirect_stderr(stderr):
            result = profile_cli._status(self.context, None)

        self.assertEqual(result, 1)
        self.assertEqual(stdout.getvalue(), "")
        self.assertNotIn(token, stderr.getvalue())
        self.assertNotIn(expected_fingerprint, stderr.getvalue())
        self.assertNotIn(str(self.receipt.profile_home), stderr.getvalue())

    def test_status_rejects_mismatched_fingerprint_without_details(self):
        token = "ef" * 32
        expected_fingerprint = profile_cli.profile_ownership.owner_fingerprint(
            self.receipt
        )
        reported_fingerprint = "sha256:" + ("0" * 64)
        self.context.data_path.mkdir(parents=True)
        (self.context.data_path / ".bridge-token").write_text(token, encoding="ascii")
        response = io.StringIO(
            json.dumps(
                {
                    "service": "hermes-excel-bridge",
                    "port": 8788,
                    "profile_name": "finance",
                    "owner_fingerprint": reported_fingerprint,
                }
            )
        )
        opener = mock.MagicMock()
        opener.return_value.__enter__.return_value = response
        stdout = io.StringIO()
        stderr = io.StringIO()

        with mock.patch.object(
            profile_cli.profile_ownership,
            "load_owner_receipt",
            return_value=self.receipt,
        ), mock.patch.object(profile_cli, "urlopen", opener), contextlib.redirect_stdout(
            stdout
        ), contextlib.redirect_stderr(stderr):
            result = profile_cli._status(self.context, None)

        self.assertEqual(result, 1)
        self.assertEqual(stdout.getvalue(), "")
        self.assertNotIn(token, stderr.getvalue())
        self.assertNotIn(expected_fingerprint, stderr.getvalue())
        self.assertNotIn(reported_fingerprint, stderr.getvalue())
        self.assertNotIn(str(self.receipt.install_path), stderr.getvalue())

    def test_status_rejects_mismatched_reported_port(self):
        token = "12" * 32
        fingerprint = profile_cli.profile_ownership.owner_fingerprint(self.receipt)
        self.context.data_path.mkdir(parents=True)
        (self.context.data_path / ".bridge-token").write_text(token, encoding="ascii")
        response = io.StringIO(
            json.dumps(
                {
                    "service": "hermes-excel-bridge",
                    "port": 8799,
                    "profile_name": "finance",
                    "owner_fingerprint": fingerprint,
                }
            )
        )
        opener = mock.MagicMock()
        opener.return_value.__enter__.return_value = response
        stderr = io.StringIO()

        with mock.patch.object(
            profile_cli.profile_ownership,
            "load_owner_receipt",
            return_value=self.receipt,
        ), mock.patch.object(profile_cli, "urlopen", opener), contextlib.redirect_stderr(stderr):
            result = profile_cli._status(self.context, None)

        self.assertEqual(result, 1)
        self.assertNotIn(token, stderr.getvalue())
        self.assertNotIn(fingerprint, stderr.getvalue())

    def test_status_rejects_invalid_owned_token_without_leaking_it(self):
        leaked_value = "not-a-token-secret-value"
        self.context.data_path.mkdir(parents=True)
        (self.context.data_path / ".bridge-token").write_text(
            leaked_value, encoding="ascii"
        )
        stderr = io.StringIO()
        with mock.patch.object(
            profile_cli.profile_ownership,
            "load_owner_receipt",
            return_value=self.receipt,
        ), mock.patch.object(profile_cli, "urlopen") as urlopen, contextlib.redirect_stderr(
            stderr
        ):
            result = profile_cli._status(self.context, None)

        self.assertEqual(result, 1)
        urlopen.assert_not_called()
        self.assertNotIn(leaked_value, stderr.getvalue())

    def test_owned_token_read_is_bounded(self):
        self.context.data_path.mkdir(parents=True)
        token_path = self.context.data_path / ".bridge-token"
        token_path.write_bytes(b"a" * 65)

        with self.assertRaisesRegex(ValueError, "exactly 64"):
            profile_cli._read_bridge_token(self.context)

    def test_local_tls_context_loads_installer_ca_without_disabling_verification(self):
        ca_dir = pathlib.Path(self.temp_dir.name) / ".office-addin-dev-certs"
        ca_dir.mkdir()
        ca_file = ca_dir / "ca.crt"
        ca_file.write_text("test fixture", encoding="utf-8")
        expected = object()
        with mock.patch.object(
            profile_cli.os, "environ", {"USERPROFILE": self.temp_dir.name}
        ), mock.patch.object(
            profile_cli.ssl, "create_default_context", return_value=expected
        ) as create_context:
            result = profile_cli._local_tls_context()

        self.assertIs(result, expected)
        create_context.assert_called_once_with(cafile=str(ca_file))


if __name__ == "__main__":
    unittest.main()
