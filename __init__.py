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
    from .adapter import build_excel_adapter, check_excel_requirements
    from .excel_tool import EXCEL_RESPONSE_SCHEMA, excel_response_available, handle_excel_response
    from .excel_policy import excel_terminal_tool_policy

    ctx.register_middleware("llm_request", excel_terminal_tool_policy)

    ctx.register_tool(
        name="excel_response",
        toolset="hermes-excel-sidecar",
        schema=EXCEL_RESPONSE_SCHEMA,
        handler=handle_excel_response,
        check_fn=excel_response_available,
        description="Capture a typed Excel proposal for task-pane review.",
        emoji="📊",
    )
    ctx.register_platform(
        name="excel",
        label="Microsoft Excel",
        adapter_factory=build_excel_adapter,
        check_fn=check_excel_requirements,
        required_env=["HERMES_EXCEL_INGEST_TOKEN"],
        allowed_users_env="HERMES_EXCEL_ALLOWED_USERS",
        allow_all_env="HERMES_EXCEL_ALLOW_ALL_USERS",
        max_message_length=8000,
        pii_safe=False,
        emoji="📊",
    )
