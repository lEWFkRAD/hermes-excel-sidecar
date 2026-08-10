from __future__ import annotations

import argparse
import importlib.util
import pathlib
import sys
import unittest

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


if __name__ == "__main__":
    unittest.main()
