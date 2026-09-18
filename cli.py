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
    from . import profile_ownership, remote_mode
else:  # Support the repository's bare-file validation tests.
    import profile_ownership
    import remote_mode

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
    remote_server = subs.add_parser("remote-server", help="Configure the VPS-owned Excel adapter")
    remote_server_subs = remote_server.add_subparsers(dest="remote_server_command", required=True)
    remote_server_subs.add_parser("setup", help="Create profile-owned loopback adapter state")
    remote_server_subs.add_parser("status", help="Inspect remote adapter configuration")
    remote_server_subs.add_parser("rollback", help="Remove only remote adapter state")
    remote_client = subs.add_parser("remote-client", help="Install the Windows Office client for a tunnel")
    remote_client_subs = remote_client.add_subparsers(dest="remote_client_command", required=True)
    client_setup = remote_client_subs.add_parser("setup", help="Install bridge/manifest without a local gateway")
    client_setup.add_argument("--adapter-url", default="http://127.0.0.1:8794/ingest")
    client_setup.add_argument("--token-file", required=True, help="Path to the securely provisioned adapter token")
    client_setup.add_argument("--bridge-port", type=int, default=8788)
    client_setup.add_argument("--ssh-host", required=True, help="Existing SSH host alias (not a password/key)")
    remote_client_subs.add_parser("rollback", help="Remove only the Windows remote client resources")
    remote_client_subs.add_parser("status", help="Inspect the Windows remote client")
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


def _remote_server_command(context: profile_ownership.ProfileContext, command: str) -> int:
    state = remote_mode.server_state_dir(context.profile_home)
    if command == "setup":
        config, token_path, created = remote_mode.write_server_state(context.profile_home, context.profile_name)
        print(json.dumps({"mode": "remote-server", "profile_name": config.profile_name,
                          "bind": f"{config.host}:{config.port}", "token_file": str(token_path),
                          "token_created": created, "gateway_restart_required": False}, indent=2))
        return 0
    if command == "status":
        try:
            config = remote_mode.load_server_config(context.profile_home)
            if config is None:
                print("Remote Excel server is not configured.", file=sys.stderr)
                return 1
            remote_mode.read_token(state / config.token_path)
        except (OSError, ValueError) as error:
            print(f"Remote Excel server configuration is invalid: {error}", file=sys.stderr)
            return 1
        print(json.dumps({"mode": "remote-server", "profile_name": config.profile_name,
                          "bind": f"{config.host}:{config.port}", "token": "<present>"}, indent=2))
        return 0
    if command == "rollback":
        for name in ("remote-server.json", ".ingest-token"):
            (state / name).unlink(missing_ok=True)
        try:
            state.rmdir()
        except OSError:
            pass
        print("Remote Excel server state removed; Hermes gateway was not restarted.")
        return 0
    return 2


def _remote_client_command(args: argparse.Namespace, context: profile_ownership.ProfileContext) -> int:
    if args.remote_client_command == "setup":
        if sys.platform != "win32":
            print("Remote Excel client installation requires Windows.", file=sys.stderr)
            return 2
        if not Path(args.token_file).is_file():
            print("Adapter token file does not exist; provision it securely first.", file=sys.stderr)
            return 2
        extra = ["-AdapterUrl", args.adapter_url, "-AdapterTokenFile", args.token_file,
                 "-BridgePort", str(args.bridge_port), "-SshHost", args.ssh_host]
        return _powershell("remote-client-install.ps1", context, extra)
    if args.remote_client_command == "rollback":
        return _powershell("remote-client-rollback.ps1", context)
    if args.remote_client_command == "status":
        return _powershell("remote-client-status.ps1", context)
    return 2


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
    if command not in {"install", "check", "rollback", "status", "remote-server", "remote-client"}:
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

    if command == "remote-server":
        return _remote_server_command(context, args.remote_server_command)
    if command == "remote-client":
        return _remote_client_command(args, context)
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
