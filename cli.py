"""CLI integration for the Hermes Excel sidecar plugin."""

from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

ROOT = Path(__file__).resolve().parent


def register_cli(parser: argparse.ArgumentParser) -> None:
    subs = parser.add_subparsers(dest="excel_sidecar_command")
    install = subs.add_parser("install", help="Install or update the add-in and bridge")
    install.add_argument("--port", type=int, default=None, help="Explicit bridge port")
    subs.add_parser("check", help="Validate the plugin and install package")
    status = subs.add_parser("status", help="Check the running bridge identity")
    status.add_argument("--port", type=int, default=None)
    subs.add_parser("rollback", help="Remove the per-user add-in installation")
    parser.set_defaults(func=excel_sidecar_command)


def _powershell(script: str, extra: list[str] | None = None) -> int:
    if sys.platform != "win32":
        print("Hermes Excel sidecar installation currently requires Windows.", file=sys.stderr)
        return 2
    command = [
        "powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass",
        "-File", str(ROOT / "install" / script),
    ]
    if extra:
        command.extend(extra)
    return subprocess.run(command, cwd=ROOT, check=False).returncode


def _status(port: int | None) -> int:
    resolved = port or int(os.environ.get("HERMES_EXCEL_PORT", "8787"))
    request = Request(f"http://127.0.0.1:{resolved}/api/health")
    token = os.environ.get("HERMES_EXCEL_BRIDGE_TOKEN", "").strip()
    if token:
        request.add_header("x-hermes-token", token)
    try:
        with urlopen(request, timeout=3) as response:
            payload = json.load(response)
    except HTTPError as error:
        print(f"Excel bridge returned HTTP {error.code}", file=sys.stderr)
        return 1
    except (URLError, OSError, ValueError) as error:
        print(f"Excel bridge unavailable on port {resolved}: {error}", file=sys.stderr)
        return 1
    if payload.get("service") != "hermes-excel-bridge":
        print(f"Port {resolved} is not the Hermes Excel bridge.", file=sys.stderr)
        return 1
    print(json.dumps(payload, indent=2))
    return 0


def excel_sidecar_command(args: argparse.Namespace) -> int:
    command = getattr(args, "excel_sidecar_command", None)
    if command == "install":
        extra = ["-Port", str(args.port)] if args.port else None
        return _powershell("apply.ps1", extra)
    if command == "check":
        return _powershell("check.ps1")
    if command == "rollback":
        return _powershell("rollback.ps1")
    if command == "status":
        return _status(args.port)
    print("Choose install, check, status, or rollback.", file=sys.stderr)
    return 2
