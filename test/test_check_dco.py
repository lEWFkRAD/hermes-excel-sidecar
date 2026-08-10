from __future__ import annotations

import unittest
from unittest.mock import patch

from scripts.check_dco import Commit, commits_between, has_author_signoff, missing_signoffs


def commit(message: str, *, name: str = "Alice Example", email: str = "alice@example.com") -> Commit:
    return Commit(sha="a" * 40, author_name=name, author_email=email, message=message)


class DcoIdentityTests(unittest.TestCase):
    def test_matching_author_identity_passes_case_and_whitespace_normalization(self):
        item = commit(
            "Subject\n\nSigned-off-by:  alice   example <ALICE@example.com>"
        )
        self.assertTrue(has_author_signoff(item))
        self.assertEqual(missing_signoffs([item]), [])

    def test_unrelated_signer_does_not_satisfy_dco(self):
        item = commit("Subject\n\nSigned-off-by: Mallory Example <mallory@example.com>")
        self.assertFalse(has_author_signoff(item))
        self.assertEqual(missing_signoffs([item]), [item])

    def test_matching_author_among_multiple_signers_passes(self):
        item = commit(
            "Subject\n\n"
            "Signed-off-by: Reviewer Example <reviewer@example.com>\n"
            "Signed-off-by: Alice Example <alice@example.com>"
        )
        self.assertTrue(has_author_signoff(item))

    @patch("scripts.check_dco.subprocess.check_output")
    def test_commit_metadata_is_loaded_separately_from_untrusted_message(self, check_output):
        check_output.side_effect = [
            "abc123\ndef456\n",
            "abc123\x00Alice Example\x00alice@example.com\x00One\x1e message\n",
            "def456\x00Bob Example\x00bob@example.com\x00Two\x1f message\n",
        ]

        commits = commits_between("base", "head")

        self.assertEqual([item.sha for item in commits], ["abc123", "def456"])
        self.assertEqual(commits[0].message, "One\x1e message")
        self.assertEqual(commits[1].author_email, "bob@example.com")


if __name__ == "__main__":
    unittest.main()
