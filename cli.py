"""CLI integration for the Hermes Excel sidecar plugin."""

from __future__ import annotations

import argparse
import json
import os
import re
import ssl
import subprocess
import sys
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

if __package__:
    from . import profile_ownership
else:  # Support the repository's bare-file validation tests.
    import profile_ownership

ROOT = Path(__file__).resolve().parent
_BRIDGE_TOKEN_RE = re.compile(rb"[0-9a-f]{64}")


def register_cli(parser: argparse.ArgumentParser, *, profile_name: str) -> None:
    subs = parser.add_subparsers(dest="excel_sidecar_command")
    install = subs.add_parser("install", help="Install or update the add-in and bridge")
    install.add_argument("--port", type=int, default=None, help="Explicit bridge port")
    install.add_argument(
        "--adopt-legacy",
        action="store_true",
        help="Adopt an existing unreceipted installation owned by this profile",
    )
    subs.add_parser("check", help="Validate the plugin and install package")
    status = subs.add_parser("status", help="Check the running bridge identity")
    status.add_argument("--port", type=int, default=None)
    subs.add_parser("rollback", help="Remove the per-user add-in installation")
    parser.set_defaults(func=excel_sidecar_command, excel_profile_name=profile_name)


def _powershell(
    script: str,
    context: profile_ownership.ProfileContext,
    extra: list[str] | None = None,
) -> int:
    if sys.platform != "win32":
        print("Hermes Excel sidecar installation currently requires Windows.", file=sys.stderr)
        return 2
    command = [
        "powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass",
        "-File", str(ROOT / "install" / script),
        "-ProfileName", context.profile_name,
        "-HermesHome", str(context.profile_home),
        "-OwnerReceiptPath", str(context.receipt_path),
    ]
    if extra:
        command.extend(extra)
    child_env = os.environ.copy()
    child_env["HERMES_HOME"] = str(context.profile_home)
    return subprocess.run(
        command,
        cwd=ROOT,
        check=False,
        env=child_env,
    ).returncode


def _runtime_environment() -> dict[str, str]:
    """Snapshot the call-time Hermes profile, including ContextVar overrides."""

    environ = os.environ.copy()
    try:
        from hermes_constants import get_hermes_home
    except ModuleNotFoundError as error:
        if error.name != "hermes_constants":
            raise
        # Bare plugin checkout validation has no Hermes runtime on sys.path.
        return environ
    environ["HERMES_HOME"] = str(get_hermes_home())
    return environ


def _local_tls_context() -> ssl.SSLContext:
    """Trust the installer-managed localhost CA without disabling TLS checks."""

    explicit_ca = os.environ.get("HERMES_EXCEL_TLS_CA", "").strip()
    if explicit_ca:
        ca_path = Path(explicit_ca).expanduser()
    else:
        user_profile = os.environ.get("USERPROFILE", "").strip()
        ca_path = Path(user_profile) / ".office-addin-dev-certs" / "ca.crt"
    if ca_path.is_file():
        return ssl.create_default_context(cafile=str(ca_path))
    return ssl.create_default_context()


def _read_bridge_token(context: profile_ownership.ProfileContext) -> str:
    """Read the selected owner's fixed-size token without consulting global env."""

    token_path = context.data_path / ".bridge-token"
    with token_path.open("rb") as handle:
        raw = handle.read(65)
    if len(raw) != 64 or _BRIDGE_TOKEN_RE.fullmatch(raw) is None:
        raise ValueError("owned bridge token must be exactly 64 lowercase hex characters")
    return raw.decode("ascii")


def _status(context: profile_ownership.ProfileContext, port: int | None) -> int:
    try:
        receipt = profile_ownership.load_owner_receipt(context.receipt_path)
    except (OSError, ValueError) as error:
        print(f"Excel bridge owner receipt is invalid: {error}", file=sys.stderr)
        return 1
    if receipt is None:
        print("Excel bridge has no owner receipt; run install for this profile.", file=sys.stderr)
        return 1
    if not profile_ownership.owner_matches(receipt, context):
        print(
            f"Excel bridge is owned by profile {receipt.profile_name!r}, "
            f"not {context.profile_name!r}.",
            file=sys.stderr,
        )
        return 1
    if port is not None and port != receipt.bridge_port:
        print(
            f"Port {port} does not match this profile's owned bridge port "
            f"{receipt.bridge_port}.",
            file=sys.stderr,
        )
        return 1

    resolved = receipt.bridge_port
    request = Request(f"https://localhost:{resolved}/api/health")
    # Ownership must be established before this process reads any credential.
    try:
        token = _read_bridge_token(context)
    except (OSError, ValueError):
        print(
            "This profile's owned bridge token is missing or invalid; reinstall it.",
            file=sys.stderr,
        )
        return 1
    request.add_header("x-hermes-token", token)
    try:
        with urlopen(request, timeout=3, context=_local_tls_context()) as response:
            payload = json.load(response)
    except HTTPError as error:
        print(f"Excel bridge returned HTTP {error.code}", file=sys.stderr)
        return 1
    except (URLError, OSError, ValueError) as error:
        print(f"Excel bridge unavailable on port {resolved}: {error}", file=sys.stderr)
        return 1
    if not isinstance(payload, dict) or payload.get("service") != "hermes-excel-bridge":
        print(f"Port {resolved} is not the Hermes Excel bridge.", file=sys.stderr)
        return 1
    expected_fingerprint = profile_ownership.owner_fingerprint(receipt)
    if (
        payload.get("port") != resolved
        or payload.get("profile_name") != context.profile_name
        or payload.get("owner_fingerprint") != expected_fingerprint
    ):
        print(
            "Excel bridge ownership attestation does not match this profile.",
            file=sys.stderr,
        )
        return 1
    # Never echo the credential even if an unhealthy local service reflects it.
    print(json.dumps(payload, indent=2).replace(token, "<redacted>"))
    return 0


def excel_sidecar_command(args: argparse.Namespace) -> int:
    command = getattr(args, "excel_sidecar_command", None)
    if command not in {"install", "check", "rollback", "status"}:
        print("Choose install, check, status, or rollback.", file=sys.stderr)
        return 2
    try:
        runtime_env = _runtime_environment()
        context = profile_ownership.resolve_profile_context(
            getattr(args, "excel_profile_name", None),
            environ=runtime_env,
        )
    except ValueError as error:
        print(f"Invalid Hermes profile selection: {error}", file=sys.stderr)
        return 2

    if command == "install":
        extra = ["-Port", str(args.port)] if args.port else []
        if getattr(args, "adopt_legacy", False):
            extra.append("-AdoptLegacy")
        return _powershell("apply.ps1", context, extra)
    if command == "check":
        return _powershell("check.ps1", context)
    if command == "rollback":
        return _powershell("rollback.ps1", context)
    if command == "status":
        return _status(context, args.port)
    raise AssertionError(f"unhandled Excel sidecar command: {command}")
