#!/usr/bin/env python3
"""Require a Developer Certificate of Origin sign-off on every PR commit."""

from __future__ import annotations

import re
import subprocess
import sys
from dataclasses import dataclass

SIGNOFF_RE = re.compile(
    r"^Signed-off-by:\s+[^<>\r\n]+\s+<[^<>\s]+@[^<>\s]+>\s*$",
    re.IGNORECASE | re.MULTILINE,
)


@dataclass(frozen=True)
class Commit:
    sha: str
    message: str


def commits_between(base: str, head: str) -> list[Commit]:
    """Return commits reachable from ``head`` but not ``base``."""
    output = subprocess.check_output(
        ["git", "log", "--format=%H%x1f%B%x1e", f"{base}..{head}"],
        text=True,
        encoding="utf-8",
    )
    commits: list[Commit] = []
    for record in output.split("\x1e"):
        record = record.strip()
        if not record:
            continue
        sha, message = record.split("\x1f", 1)
        commits.append(Commit(sha=sha.strip(), message=message.strip()))
    return commits


def missing_signoffs(commits: list[Commit]) -> list[Commit]:
    """Return commits that do not contain a valid ``Signed-off-by`` trailer."""
    return [commit for commit in commits if SIGNOFF_RE.search(commit.message) is None]


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

    print("The following commits are missing a valid Signed-off-by trailer:", file=sys.stderr)
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
