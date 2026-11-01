"""Provider-boundary policy for the Excel terminal response tool."""

from __future__ import annotations

from copy import deepcopy
from typing import Any

from .excel_runtime import get_active_request


def _tool_name(tool: Any) -> str:
    if not isinstance(tool, dict):
        return ""
    function = tool.get("function")
    if isinstance(function, dict):
        return str(function.get("name") or "")
    return str(tool.get("name") or "")


def excel_terminal_tool_policy(*, request: dict[str, Any], platform: Any = "",
                               api_call_count: int = 0, model: Any = "",
                               provider: Any = "", api_mode: Any = "", **_: Any) -> dict[str, Any] | None:
    """Force one typed Excel response, then reserve the next call for prose."""
    platform_name = str(getattr(platform, "value", platform) or "").lower()
    if platform_name != "excel":
        return None

    def named_choice(name: str) -> dict[str, Any]:
        if str(api_mode or "") in {"codex_responses", "responses"} or "input" in request:
            return {"type": "function", "name": name}
        return {"type": "function", "function": {"name": name}}

    rewritten = deepcopy(request)
    tools = rewritten.get("tools")
    names = [_tool_name(item) for item in tools] if isinstance(tools, list) else []
    if "excel_response" not in names:
        # The middleware framework currently swallows callback exceptions.
        # An impossible named choice makes the provider fail closed instead
        # of silently producing prose without the typed terminal channel.
        rewritten["tool_choice"] = named_choice("__excel_response_unavailable__")
        return {"request": rewritten, "source": "hermes-excel-sidecar",
                "reason": "excel_response tool missing; fail closed"}

    pending = get_active_request()
    # Validation errors do not capture a proposal. Permit a bounded correction
    # instead of forcing prose after the first failed tool call. Success still
    # seals the request immediately; correlation errors remain fail-closed.
    force_proposal = (not pending.proposal_captured and not pending.protocol_error
                      and int(api_call_count or 0) <= 3) if pending else int(api_call_count or 0) == 1
    if force_proposal:
        model_name = str(model or request.get("model") or "").lower()
        provider_name = str(provider or "").lower()
        if provider_name == "nous" and model_name.startswith("qwen/"):
            # Nous-hosted Qwen 3.8 accepts tools and reliably follows the
            # terminal-tool instruction, but its upstream rejects both the
            # named-function object and "required" tool_choice forms with
            # HTTP 400.  "auto" preserves typed proposals; the adapter still
            # fails closed if the model returns prose before excel_response.
            rewritten["tool_choice"] = "auto"
        else:
            rewritten["tool_choice"] = named_choice("excel_response")
    else:
        rewritten["tool_choice"] = "none"
    return {"request": rewritten, "source": "hermes-excel-sidecar",
            "reason": "enforce one typed Excel proposal before final prose"}
