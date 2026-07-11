"""Hermes Excel sidecar user plugin."""

from __future__ import annotations

from .cli import excel_sidecar_command, register_cli


def register(ctx) -> None:
    """Register the operator-facing Excel sidecar CLI."""
    ctx.register_cli_command(
        name="excel-sidecar",
        help="Install and operate the Hermes Excel sidecar",
        setup_fn=register_cli,
        handler_fn=excel_sidecar_command,
        description=(
            "Install, validate, inspect, or roll back the local Office.js "
            "task pane and authenticated Hermes bridge."
        ),
    )
