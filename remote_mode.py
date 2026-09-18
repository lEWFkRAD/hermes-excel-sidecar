"""Explicit remote client/server deployment contracts for Hermes for Excel.

The server owns the Excel platform adapter and its credential on the VPS.  The
client owns only Office, the localhost bridge, and the SSH forward.  This module
contains the shared, non-secret configuration format and intentionally performs
no gateway restart or Office mutation.
"""
from __future__ import annotations

import hashlib
import json
import os
import secrets
from dataclasses import dataclass
from pathlib import Path
from typing import Mapping

TOKEN_RE = __import__("re").compile(r"^[0-9a-f]{64}$")
SCHEMA = 1
DEFAULT_ADAPTER_PORT = 8794
DEFAULT_BRIDGE_PORT = 8788


@dataclass(frozen=True)
class RemoteServerConfig:
    profile_name: str
    host: str = "127.0.0.1"
    port: int = DEFAULT_ADAPTER_PORT
    token_path: Path = Path("excel-adapter") / ".ingest-token"
    schema: int = SCHEMA

    def to_dict(self) -> dict[str, object]:
        return {"schema": self.schema, "mode": "remote-server", "profile_name": self.profile_name,
                "host": self.host, "port": self.port, "token_path": str(self.token_path)}


def server_state_dir(profile_home: str | os.PathLike[str]) -> Path:
    return Path(profile_home).expanduser().resolve() / "excel-adapter"


def server_config_path(profile_home: str | os.PathLike[str]) -> Path:
    return server_state_dir(profile_home) / "remote-server.json"


def load_server_config(profile_home: str | os.PathLike[str]) -> RemoteServerConfig | None:
    path = server_config_path(profile_home)
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except FileNotFoundError:
        return None
    except (OSError, json.JSONDecodeError) as exc:
        raise ValueError("remote Excel server configuration is invalid") from exc
    if not isinstance(payload, dict):
        raise ValueError("remote Excel server configuration is invalid")
    if (payload.get("schema"), payload.get("mode")) != (SCHEMA, "remote-server"):
        raise ValueError("remote Excel server configuration has an unsupported schema")
    profile = payload.get("profile_name")
    host = payload.get("host")
    port = payload.get("port")
    token_path = payload.get("token_path")
    if not isinstance(profile, str) or not profile or host != "127.0.0.1" or type(port) is not int or port != DEFAULT_ADAPTER_PORT:
        raise ValueError("remote Excel adapter must be bound to 127.0.0.1:8794")
    if not isinstance(token_path, str) or Path(token_path).is_absolute() or ".." in Path(token_path).parts:
        raise ValueError("remote Excel token path must be relative to the profile state directory")
    return RemoteServerConfig(profile_name=profile, host=host, port=port, token_path=Path(token_path))


def read_token(path: Path) -> str:
    if path.is_symlink():
        raise ValueError("remote Excel token must not be a symlink")
    raw = path.read_text(encoding="ascii").strip()
    if not TOKEN_RE.fullmatch(raw):
        raise ValueError("remote Excel token must be exactly 64 lowercase hex characters")
    return raw


def write_server_state(profile_home: str | os.PathLike[str], profile_name: str) -> tuple[RemoteServerConfig, Path, bool]:
    state = server_state_dir(profile_home)
    state.mkdir(mode=0o700, parents=True, exist_ok=True)
    token_path = state / ".ingest-token"
    created = not token_path.exists()
    token = secrets.token_hex(32) if created else read_token(token_path)
    if created:
        token_path.write_text(token, encoding="ascii")
        os.chmod(token_path, 0o600)
    config = RemoteServerConfig(profile_name=profile_name, token_path=Path(".ingest-token"))
    tmp = state / ".remote-server.json.tmp"
    tmp.write_text(json.dumps(config.to_dict(), indent=2) + "\n", encoding="utf-8")
    os.chmod(tmp, 0o600)
    tmp.replace(server_config_path(profile_home))
    return config, token_path, created


def owner_fingerprint(config: RemoteServerConfig) -> str:
    material = f"remote-server:{config.profile_name}:{config.host}:{config.port}".encode()
    return "sha256:" + hashlib.sha256(material).hexdigest()


def client_environment(*, adapter_url: str, token_file: str, profile_name: str = "default") -> dict[str, str]:
    if adapter_url != "http://127.0.0.1:8794/ingest":
        raise ValueError("remote client adapter URL must use the localhost tunnel")
    if not token_file or not profile_name:
        raise ValueError("remote client token file and profile name are required")
    return {"HERMES_EXCEL_ADAPTER_URL": adapter_url, "HERMES_EXCEL_INGEST_TOKEN_FILE": token_file,
            "HERMES_EXCEL_PROFILE_NAME": profile_name, "HERMES_EXCEL_REMOTE_CLIENT": "1",
            "HERMES_EXCEL_TRANSPORT": "platform-only", "HERMES_EXCEL_ALLOW_RAW_FALLBACK": "0"}
