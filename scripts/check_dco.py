#!/usr/bin/env python3
"""Require a Developer Certificate of Origin sign-off on every PR commit."""

from __future__ import annotations

import re
import subprocess
import sys
from dataclasses import dataclass

SIGNOFF_RE = re.compile(
    r"^Signed-off-by:\s+(?P<name>[^<>\r\n]+?)\s+"
    r"<(?P<email>[^<>\s]+@[^<>\s]+)>\s*$",
    re.IGNORECASE | re.MULTILINE,
)


@dataclass(frozen=True)
class Commit:
    sha: str
    author_name: str
    author_email: str
    message: str


def commits_between(base: str, head: str) -> list[Commit]:
    """Return commits reachable from ``head`` but not ``base``."""
    revision_output = subprocess.check_output(
        ["git", "rev-list", f"{base}..{head}"],
        text=True,
        encoding="utf-8",
    )
    commits: list[Commit] = []
    for sha in revision_output.splitlines():
        sha = sha.strip()
        if not sha:
            continue
        record = subprocess.check_output(
            ["git", "show", "-s", "--format=%H%x00%an%x00%ae%x00%B", sha],
            text=True,
            encoding="utf-8",
        )
        commit_sha, author_name, author_email, message = record.split("\x00", 3)
        commits.append(
            Commit(
                sha=commit_sha.strip(),
                author_name=author_name.strip(),
                author_email=author_email.strip(),
                message=message.strip(),
            )
        )
    return commits


def _normalize_name(value: str) -> str:
    return " ".join(value.split()).casefold()


def _normalize_email(value: str) -> str:
    return value.strip().casefold()


def has_author_signoff(commit: Commit) -> bool:
    """Return whether a trailer matches the commit author's identity."""
    expected_name = _normalize_name(commit.author_name)
    expected_email = _normalize_email(commit.author_email)
    return any(
        _normalize_name(match.group("name")) == expected_name
        and _normalize_email(match.group("email")) == expected_email
        for match in SIGNOFF_RE.finditer(commit.message)
    )


def missing_signoffs(commits: list[Commit]) -> list[Commit]:
    """Return commits without a trailer matching their author identity."""
    return [commit for commit in commits if not has_author_signoff(commit)]


def main(argv: list[str] | None = None) -> int:
    args = sys.argv[1:] if argv is None else argv
    if len(args) != 2:
        print("usage: check_dco.py <base-sha> <head-sha>", file=sys.stderr)
        return 2

    commits = commits_between(args[0], args[1])
    if not commits:
        print("No pull-request commits found in the requested range.", file=sys.stderr)
        return 1

    unsigned = missing_signoffs(commits)
    if not unsigned:
        print(f"DCO sign-off present on all {len(commits)} commit(s).")
        return 0

    print(
        "The following commits are missing a Signed-off-by trailer matching "
        "their author name and email:",
        file=sys.stderr,
    )
    for commit in unsigned:
        subject = commit.message.splitlines()[0] if commit.message else "(no subject)"
        print(f"- {commit.sha[:12]} {subject}", file=sys.stderr)
    print(
        "Amend each commit with `git commit --amend --signoff` (or `git commit -s`) "
        "and update the pull-request branch.",
        file=sys.stderr,
    )
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
