"""Provider-boundary policy for the Excel terminal response tool."""

from __future__ import annotations

from copy import deepcopy
from typing import Any


def _tool_name(tool: Any) -> str:
    if not isinstance(tool, dict):
        return ""
    function = tool.get("function")
    if isinstance(function, dict):
        return str(function.get("name") or "")
    return str(tool.get("name") or "")


def excel_terminal_tool_policy(*, request: dict[str, Any], platform: Any = "",
                               api_call_count: int = 0, model: Any = "",
                               provider: Any = "", **_: Any) -> dict[str, Any] | None:
    """Force one typed Excel response, then reserve the next call for prose."""
    platform_name = str(getattr(platform, "value", platform) or "").lower()
    if platform_name != "excel":
        return None

    rewritten = deepcopy(request)
    tools = rewritten.get("tools")
    names = [_tool_name(item) for item in tools] if isinstance(tools, list) else []
    if "excel_response" not in names:
        # The middleware framework currently swallows callback exceptions.
        # An impossible named choice makes the provider fail closed instead
        # of silently producing prose without the typed terminal channel.
        rewritten["tool_choice"] = {
            "type": "function", "function": {"name": "__excel_response_unavailable__"}
        }
        return {"request": rewritten, "source": "hermes-excel-sidecar",
                "reason": "excel_response tool missing; fail closed"}

    if int(api_call_count or 0) == 1:
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
            rewritten["tool_choice"] = {
                "type": "function", "function": {"name": "excel_response"}
            }
    else:
        rewritten["tool_choice"] = "none"
    return {"request": rewritten, "source": "hermes-excel-sidecar",
            "reason": "enforce one typed Excel proposal before final prose"}
