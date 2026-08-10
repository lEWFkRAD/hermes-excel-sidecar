"""Profile and shared-owner identity for the Windows Excel sidecar.

This module is intentionally stdlib-only and resolves the process environment
at call time.  It can therefore be imported from a bare plugin checkout and it
does not retain a stale ``HERMES_HOME`` when a long-lived process changes
profile scope.
"""

from __future__ import annotations

import hashlib
import json
import ntpath
import os
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Mapping

OWNER_SCHEMA = 1
OWNER_RECEIPT_RELATIVE_PATH = Path("shared") / "excel-sidecar" / "owner-v1.json"
OFFICE_MANIFEST_ID = "4fd4d435-7f9a-4d6d-9251-32f154f83a1f"
MAX_RECEIPT_BYTES = 16 * 1024

_PROFILE_NAME_RE = re.compile(r"^[a-z0-9][a-z0-9_-]{0,63}$")
_SEMVER_RE = re.compile(
    r"^(?:0|[1-9][0-9]*)\."
    r"(?:0|[1-9][0-9]*)\."
    r"(?:0|[1-9][0-9]*)"
    r"(?:-[0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*)?"
    r"(?:\+[0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*)?$"
)
_RECEIPT_FIELDS = frozenset(
    {
        "schema",
        "profile_name",
        "profile_home",
        "install_path",
        "bridge_port",
        "plugin_version",
        "manifest_id",
    }
)


def _canonical_path(value: str | os.PathLike[str]) -> Path:
    if not isinstance(value, (str, os.PathLike)):
        raise ValueError("path must be a string or path-like value")
    raw = os.fspath(value)
    if not isinstance(raw, str) or not raw.strip():
        raise ValueError("path cannot be empty")
    return Path(raw).expanduser().resolve(strict=False)


def _windows_path_key(value: str | os.PathLike[str]) -> str:
    """Return a stable, case-insensitive identity key for a Windows path."""

    canonical = str(_canonical_path(value)).replace("/", "\\")
    return ntpath.normcase(ntpath.normpath(canonical))


def _strict_receipt_path(value: object, field: str) -> Path:
    if not isinstance(value, str) or not value.strip():
        raise ValueError(f"{field} must be a non-empty absolute path string")
    if value != value.strip() or not Path(value).is_absolute():
        raise ValueError(f"{field} must be an absolute canonical path")
    if os.path.normpath(value) != value:
        raise ValueError(f"{field} must not contain redundant or relative segments")
    canonical = _canonical_path(value)
    raw_key = ntpath.normcase(ntpath.normpath(value.replace("/", "\\")))
    if raw_key != _windows_path_key(canonical):
        raise ValueError(f"{field} must already be resolved")
    return canonical


def _default_hermes_home(environ: Mapping[str, str]) -> Path:
    local_appdata = str(environ.get("LOCALAPPDATA", "")).strip()
    if local_appdata:
        return _canonical_path(Path(local_appdata) / "hermes")
    if sys.platform == "win32":
        return _canonical_path(Path.home() / "AppData" / "Local" / "hermes")
    return _canonical_path(Path.home() / ".hermes")


def _normalize_profile_name(value: object) -> str:
    if not isinstance(value, str):
        raise ValueError("profile_name must be a string")
    name = value.strip().lower()
    if not name or not _PROFILE_NAME_RE.fullmatch(name):
        raise ValueError("profile_name must match [a-z0-9][a-z0-9_-]{0,63}")
    return name


def _parse_receipt_profile_name(value: object) -> str:
    name = _normalize_profile_name(value)
    if value != name:
        raise ValueError("profile_name must already be normalized lowercase")
    return name


@dataclass(frozen=True)
class ProfileContext:
    """Resolved profile-owned paths and the shared Windows-user receipt."""

    profile_name: str
    profile_home: Path
    hermes_root: Path
    install_path: Path
    data_path: Path
    config_path: Path
    receipt_path: Path


@dataclass(frozen=True)
class OwnerReceipt:
    """Strict, non-secret record of the active singleton owner."""

    schema: int
    profile_name: str
    profile_home: Path
    install_path: Path
    bridge_port: int
    plugin_version: str
    manifest_id: str

    @classmethod
    def from_context(
        cls,
        context: ProfileContext,
        *,
        bridge_port: int,
        plugin_version: str,
    ) -> "OwnerReceipt":
        return parse_owner_receipt(
            {
                "schema": OWNER_SCHEMA,
                "profile_name": context.profile_name,
                "profile_home": str(context.profile_home),
                "install_path": str(context.install_path),
                "bridge_port": bridge_port,
                "plugin_version": plugin_version,
                "manifest_id": OFFICE_MANIFEST_ID,
            }
        )

    def to_dict(self) -> dict[str, object]:
        return {
            "schema": self.schema,
            "profile_name": self.profile_name,
            "profile_home": str(self.profile_home),
            "install_path": str(self.install_path),
            "bridge_port": self.bridge_port,
            "plugin_version": self.plugin_version,
            "manifest_id": self.manifest_id,
        }


def shared_receipt_path(environ: Mapping[str, str] | None = None) -> Path:
    """Return the Windows-user shared receipt path, independent of profile."""

    env = os.environ if environ is None else environ
    return _default_hermes_home(env) / OWNER_RECEIPT_RELATIVE_PATH


def resolve_profile_context(
    profile_hint: str | None = None,
    environ: Mapping[str, str] | None = None,
) -> ProfileContext:
    """Resolve the active profile from the current environment.

    ``profile_hint`` is an assertion, not a path selector.  A non-default hint
    with no ``HERMES_HOME`` fails closed, as does a hint that disagrees with the
    home-derived identity.
    """

    env = os.environ if environ is None else environ
    default_home = _default_hermes_home(env)
    raw_home = str(env.get("HERMES_HOME", "")).strip()
    hint = _normalize_profile_name(profile_hint) if profile_hint is not None else None

    if raw_home:
        profile_home = _canonical_path(raw_home)
    else:
        if hint not in (None, "default"):
            raise ValueError("HERMES_HOME is required for a non-default profile")
        profile_home = default_home

    if _windows_path_key(profile_home) == _windows_path_key(default_home):
        inferred_name = "default"
        hermes_root = default_home
    elif profile_home.parent.name.casefold() == "profiles":
        inferred_name = _normalize_profile_name(profile_home.name)
        hermes_root = _canonical_path(profile_home.parent.parent)
    else:
        inferred_name = "custom"
        hermes_root = profile_home

    # In a custom deployment, Hermes treats the custom root itself as its
    # built-in default profile.  The registration context is authoritative for
    # that otherwise-ambiguous label; the path remains the ownership anchor.
    if inferred_name == "custom" and hint == "default":
        inferred_name = "default"
    elif hint is not None and hint != inferred_name:
        raise ValueError(
            f"profile hint {hint!r} does not match HERMES_HOME identity {inferred_name!r}"
        )

    return ProfileContext(
        profile_name=inferred_name,
        profile_home=profile_home,
        hermes_root=hermes_root,
        install_path=profile_home / "excel-addin",
        data_path=profile_home / "excel-addin" / "data",
        config_path=profile_home / "config.yaml",
        receipt_path=shared_receipt_path(env),
    )


def _decode_json_object(value: str | bytes | bytearray) -> Mapping[str, object]:
    if isinstance(value, (bytes, bytearray)):
        if len(value) > MAX_RECEIPT_BYTES:
            raise ValueError("owner receipt exceeds the size limit")
        try:
            text = bytes(value).decode("utf-8")
        except UnicodeDecodeError as error:
            raise ValueError("owner receipt must be UTF-8 JSON") from error
    else:
        text = value
        if len(text.encode("utf-8")) > MAX_RECEIPT_BYTES:
            raise ValueError("owner receipt exceeds the size limit")

    def no_duplicates(pairs: list[tuple[str, object]]) -> dict[str, object]:
        result: dict[str, object] = {}
        for key, item in pairs:
            if key in result:
                raise ValueError(f"duplicate owner receipt field: {key}")
            result[key] = item
        return result

    try:
        decoded = json.loads(text, object_pairs_hook=no_duplicates)
    except (json.JSONDecodeError, UnicodeError) as error:
        raise ValueError("owner receipt must be valid JSON") from error
    if not isinstance(decoded, dict):
        raise ValueError("owner receipt must be a JSON object")
    return decoded


def parse_owner_receipt(
    value: Mapping[str, object] | str | bytes | bytearray,
) -> OwnerReceipt:
    """Validate and canonicalize a receipt, rejecting extras and coercions."""

    if isinstance(value, (str, bytes, bytearray)):
        payload = _decode_json_object(value)
    elif isinstance(value, Mapping):
        payload = value
    else:
        raise ValueError("owner receipt must be a mapping or JSON object")

    if any(not isinstance(key, str) for key in payload):
        raise ValueError("owner receipt field names must be strings")
    keys = set(payload)
    if keys != _RECEIPT_FIELDS:
        missing = sorted(_RECEIPT_FIELDS - keys)
        extra = sorted(keys - _RECEIPT_FIELDS)
        details = []
        if missing:
            details.append(f"missing fields: {', '.join(missing)}")
        if extra:
            details.append(f"unknown fields: {', '.join(extra)}")
        raise ValueError("invalid owner receipt fields (" + "; ".join(details) + ")")

    if type(payload["schema"]) is not int or payload["schema"] != OWNER_SCHEMA:
        raise ValueError(f"schema must be the integer {OWNER_SCHEMA}")
    profile_name = _parse_receipt_profile_name(payload["profile_name"])

    profile_home = _strict_receipt_path(payload["profile_home"], "profile_home")
    install_path = _strict_receipt_path(payload["install_path"], "install_path")

    bridge_port = payload["bridge_port"]
    if type(bridge_port) is not int or not 1 <= bridge_port <= 65535:
        raise ValueError("bridge_port must be an integer from 1 through 65535")

    plugin_version = payload["plugin_version"]
    if not isinstance(plugin_version, str) or not _SEMVER_RE.fullmatch(plugin_version):
        raise ValueError("plugin_version must be a semantic version")

    manifest_id = payload["manifest_id"]
    if not isinstance(manifest_id, str) or manifest_id.casefold() != OFFICE_MANIFEST_ID:
        raise ValueError("manifest_id is not the Hermes Excel add-in manifest")

    return OwnerReceipt(
        schema=OWNER_SCHEMA,
        profile_name=profile_name,
        profile_home=profile_home,
        install_path=install_path,
        bridge_port=bridge_port,
        plugin_version=plugin_version,
        manifest_id=OFFICE_MANIFEST_ID,
    )


def load_owner_receipt(
    path: str | os.PathLike[str] | None = None,
    *,
    environ: Mapping[str, str] | None = None,
) -> OwnerReceipt | None:
    """Load a receipt, returning ``None`` only when it does not exist."""

    receipt_path = shared_receipt_path(environ) if path is None else _canonical_path(path)
    try:
        with receipt_path.open("rb") as handle:
            raw = handle.read(MAX_RECEIPT_BYTES + 1)
    except FileNotFoundError:
        return None
    if len(raw) > MAX_RECEIPT_BYTES:
        raise ValueError("owner receipt exceeds the size limit")
    return parse_owner_receipt(raw)


def owner_matches(receipt: OwnerReceipt, context: ProfileContext) -> bool:
    """Return whether the receipt owns this profile's singleton installation."""

    return (
        receipt.schema == OWNER_SCHEMA
        and receipt.manifest_id == OFFICE_MANIFEST_ID
        and receipt.profile_name == context.profile_name
        and _windows_path_key(receipt.profile_home) == _windows_path_key(context.profile_home)
        and _windows_path_key(receipt.install_path) == _windows_path_key(context.install_path)
    )


def owner_fingerprint(receipt: OwnerReceipt) -> str:
    """Return a non-secret, case-insensitive fingerprint of the full receipt."""

    identity = {
        "bridge_port": receipt.bridge_port,
        "install_path": _windows_path_key(receipt.install_path),
        "manifest_id": receipt.manifest_id.casefold(),
        "plugin_version": receipt.plugin_version,
        "profile_home": _windows_path_key(receipt.profile_home),
        "profile_name": receipt.profile_name,
        "schema": receipt.schema,
    }
    encoded = json.dumps(identity, sort_keys=True, separators=(",", ":")).encode("utf-8")
    return "sha256:" + hashlib.sha256(encoded).hexdigest()
