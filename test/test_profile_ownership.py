from __future__ import annotations

import json
import pathlib
import sys
import tempfile
import unittest

import importlib.util

ROOT = pathlib.Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location("profile_ownership_bare", ROOT / "profile_ownership.py")
ownership = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = ownership
assert SPEC.loader
SPEC.loader.exec_module(ownership)


class ProfileResolutionTests(unittest.TestCase):
    def test_default_named_and_custom_profiles_resolve_without_hermes_runtime(self):
        with tempfile.TemporaryDirectory() as temp:
            base = pathlib.Path(temp)
            env = {"LOCALAPPDATA": str(base / "Local")}

            default = ownership.resolve_profile_context(environ=env)
            self.assertEqual(default.profile_name, "default")
            self.assertEqual(default.profile_home, (base / "Local" / "hermes").resolve())

            env["HERMES_HOME"] = str(base / "Local" / "hermes" / "profiles" / "alpha")
            named = ownership.resolve_profile_context("alpha", env)
            self.assertEqual(named.profile_name, "alpha")
            self.assertEqual(named.hermes_root, default.profile_home)

            env["HERMES_HOME"] = str(base / "portable-hermes")
            custom = ownership.resolve_profile_context("custom", env)
            self.assertEqual(custom.profile_name, "custom")
            self.assertEqual(custom.hermes_root, custom.profile_home)

            custom_default = ownership.resolve_profile_context("default", env)
            self.assertEqual(custom_default.profile_name, "default")
            self.assertEqual(custom_default.profile_home, custom.profile_home)

    def test_resolution_reads_changed_environment_at_each_call(self):
        with tempfile.TemporaryDirectory() as temp:
            root = pathlib.Path(temp) / "hermes"
            env = {"LOCALAPPDATA": temp, "HERMES_HOME": str(root / "profiles" / "alpha")}
            alpha = ownership.resolve_profile_context(environ=env)
            env["HERMES_HOME"] = str(root / "profiles" / "beta")
            beta = ownership.resolve_profile_context(environ=env)

            self.assertEqual((alpha.profile_name, beta.profile_name), ("alpha", "beta"))
            self.assertNotEqual(alpha.install_path, beta.install_path)
            self.assertNotEqual(alpha.data_path, beta.data_path)
            self.assertEqual(alpha.data_path, alpha.install_path / "data")
            self.assertEqual(alpha.receipt_path, beta.receipt_path)

    def test_profile_hint_is_a_fail_closed_assertion(self):
        with tempfile.TemporaryDirectory() as temp:
            env = {"LOCALAPPDATA": temp}
            with self.assertRaisesRegex(ValueError, "HERMES_HOME is required"):
                ownership.resolve_profile_context("alpha", env)
            env["HERMES_HOME"] = str(pathlib.Path(temp) / "hermes" / "profiles" / "alpha")
            with self.assertRaisesRegex(ValueError, "does not match"):
                ownership.resolve_profile_context("beta", env)

    def test_shared_receipt_is_not_scoped_to_active_custom_home(self):
        with tempfile.TemporaryDirectory() as temp:
            env = {"LOCALAPPDATA": str(pathlib.Path(temp) / "Local")}
            expected = (pathlib.Path(env["LOCALAPPDATA"]) / "hermes" / "shared" /
                        "excel-sidecar" / "owner-v1.json").resolve()
            for home in (pathlib.Path(temp) / "one", pathlib.Path(temp) / "two"):
                env["HERMES_HOME"] = str(home)
                self.assertEqual(ownership.shared_receipt_path(env), expected)


class OwnerReceiptTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        root = pathlib.Path(self.temp.name) / "hermes"
        self.env = {"LOCALAPPDATA": self.temp.name,
                    "HERMES_HOME": str(root / "profiles" / "finance")}
        self.context = ownership.resolve_profile_context(environ=self.env)
        self.receipt = ownership.OwnerReceipt.from_context(
            self.context, bridge_port=8788, plugin_version="0.2.0"
        )

    def tearDown(self):
        self.temp.cleanup()

    def test_round_trip_and_owner_match_are_case_insensitive_for_windows_paths(self):
        payload = self.receipt.to_dict()
        payload["profile_home"] = str(payload["profile_home"]).upper()
        payload["install_path"] = str(payload["install_path"]).swapcase()
        reparsed = ownership.parse_owner_receipt(json.dumps(payload))

        self.assertTrue(ownership.owner_matches(reparsed, self.context))
        self.assertEqual(
            ownership.owner_fingerprint(reparsed),
            ownership.owner_fingerprint(self.receipt),
        )

    def test_owner_match_rejects_other_profile_even_with_same_port(self):
        other_env = dict(self.env)
        other_env["HERMES_HOME"] = str(
            pathlib.Path(self.temp.name) / "hermes" / "profiles" / "personal"
        )
        other = ownership.resolve_profile_context(environ=other_env)
        self.assertFalse(ownership.owner_matches(self.receipt, other))

    def test_parser_rejects_missing_extra_duplicate_and_coerced_fields(self):
        valid = self.receipt.to_dict()
        invalid = []
        missing = dict(valid)
        missing.pop("profile_home")
        invalid.append(missing)
        invalid.append({**valid, "token": "must-never-be-stored"})
        invalid.append({**valid, "schema": True})
        invalid.append({**valid, "bridge_port": "8788"})
        invalid.append({**valid, "plugin_version": "v0.2"})
        invalid.append({**valid, "profile_name": "Finance"})
        invalid.append({**valid, "profile_home": "relative-profile"})
        invalid.append({**valid, "profile_home": str(self.context.profile_home / ".." / "finance")})
        invalid.append({**valid, "manifest_id": "00000000-0000-0000-0000-000000000000"})
        for payload in invalid:
            with self.subTest(payload=payload), self.assertRaises(ValueError):
                ownership.parse_owner_receipt(payload)

        duplicate = json.dumps(valid)[:-1] + ',"bridge_port":9999}'
        with self.assertRaisesRegex(ValueError, "duplicate"):
            ownership.parse_owner_receipt(duplicate)

    def test_load_returns_none_only_for_missing_and_rejects_invalid_file(self):
        path = pathlib.Path(self.temp.name) / "owner.json"
        self.assertIsNone(ownership.load_owner_receipt(path))
        path.write_text(json.dumps(self.receipt.to_dict()), encoding="utf-8")
        self.assertEqual(ownership.load_owner_receipt(path), self.receipt)
        path.write_text("{}", encoding="utf-8")
        with self.assertRaises(ValueError):
            ownership.load_owner_receipt(path)

    def test_fingerprint_covers_runtime_identity_without_exposing_values(self):
        changed = ownership.OwnerReceipt.from_context(
            self.context, bridge_port=8789, plugin_version="0.2.0"
        )
        fingerprint = ownership.owner_fingerprint(self.receipt)
        self.assertRegex(fingerprint, r"^sha256:[0-9a-f]{64}$")
        self.assertNotEqual(fingerprint, ownership.owner_fingerprint(changed))
        self.assertNotIn("finance", fingerprint)


if __name__ == "__main__":
    unittest.main()
