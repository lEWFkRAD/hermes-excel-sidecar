"""Capture-only terminal tool for typed Excel proposals."""

from __future__ import annotations

import json
import math
import re
from typing import Any

from .excel_runtime import capture_proposal

MAX_ACTIONS = 50
MAX_PAYLOAD_BYTES = 512_000
ID_PATTERN = r"^[A-Za-z0-9][A-Za-z0-9._:-]{7,127}$"
CELL = {"type": ["string", "number", "boolean", "null"]}
MATRIX = {
    "type": "array", "minItems": 1, "maxItems": 300,
    "items": {"type": "array", "minItems": 1, "maxItems": 30, "items": CELL},
}


def _action(kind: str, properties: dict[str, Any], required: list[str]) -> dict[str, Any]:
    return {
        "type": "object", "additionalProperties": False,
        "properties": {"type": {"const": kind}, **properties},
        "required": ["type", *required],
    }


RANGE = {"type": "string", "minLength": 1, "maxLength": 160}
ACTION_SCHEMAS = [
    _action("write_cells", {"start_cell": RANGE, "values": MATRIX, "allow_overwrite": {"type": "boolean"},
                             "auto_format": {"type": "boolean"}}, ["values"]),
    _action("create_sheet", {"name": {"type": "string", "minLength": 1, "maxLength": 31}, "values": MATRIX}, ["name", "values"]),
    _action("format_cells", {"range": RANGE, "style": {"oneOf": [{"type": "string", "maxLength": 40},
        {"type": "array", "maxItems": 8, "items": {"type": "string", "maxLength": 40}}]},
        "bold": {"type": "boolean"}, "italic": {"type": "boolean"}, "underline": {"type": "boolean"},
        "font_color": {"type": "string", "maxLength": 20}, "fill_color": {"type": "string", "maxLength": 20},
        "font_size": {"type": "number", "minimum": 1, "maximum": 72}, "font_name": {"type": "string", "maxLength": 80},
        "number_format": {"type": "string", "maxLength": 120}, "wrap_text": {"type": "boolean"},
        "number_format_dp": {"type": "integer", "minimum": 0, "maximum": 10}, "currency_symbol": {"type": "string", "maxLength": 8},
        "horizontal_alignment": {"type": "string", "maxLength": 30}, "vertical_alignment": {"type": "string", "maxLength": 30},
        "borders": {"type": "string", "maxLength": 30}, "border_color": {"type": "string", "maxLength": 20},
        "border_top": {"type": "string", "maxLength": 30}, "border_bottom": {"type": "string", "maxLength": 30},
        "border_left": {"type": "string", "maxLength": 30}, "border_right": {"type": "string", "maxLength": 30},
        "column_width": {"type": "number", "minimum": 1, "maximum": 255},
        "row_height": {"type": "number", "minimum": 1, "maximum": 409},
        "auto_fit": {"type": "boolean"}}, ["range"]),
    _action("conditional_format", {"range": RANGE, "operator": {"type": "string", "maxLength": 32},
        "value": CELL, "value2": CELL, "fill_color": {"type": "string", "maxLength": 20},
        "font_color": {"type": "string", "maxLength": 20}}, ["range", "operator", "value"]),
    _action("read_range", {"range": RANGE, "reason": {"type": "string", "maxLength": 240}}, ["range"]),
    _action("export", {"name": {"type": "string", "minLength": 1, "maxLength": 120}, "values": MATRIX}, ["name", "values"]),
]
ACTION_SCHEMAS.append(_action("merge_cells", {"range": RANGE, "across": {"type": "boolean"}}, ["range"]))
ACTION_SCHEMAS.append(_action("unmerge_cells", {"range": RANGE}, ["range"]))
ACTION_SCHEMAS.append(_action("autofit", {"range": RANGE, "columns": {"type": "boolean"}, "rows": {"type": "boolean"}}, ["range"]))
for _kind in ("insert_rows", "delete_rows", "insert_columns", "delete_columns"):
    at_schema = ({"type": "string", "pattern": "^[A-Za-z]{1,3}$", "maxLength": 3}
                 if "columns" in _kind else {"type": "integer", "minimum": 1, "maximum": 1_048_576})
    ACTION_SCHEMAS.append(_action(_kind, {"sheet": RANGE, "at": at_schema,
        "count": {"type": "integer", "minimum": 1, "maximum": 1000}}, ["at", "count"]))
ACTION_SCHEMAS.extend([
    _action("set_column_width", {"range": RANGE, "width": {"type": "number", "minimum": 1, "maximum": 255}}, ["range", "width"]),
    _action("set_row_height", {"range": RANGE, "height": {"type": "number", "minimum": 1, "maximum": 409}}, ["range", "height"]),
    _action("freeze_panes", {"rows": {"type": "integer", "minimum": 0, "maximum": 1000},
        "columns": {"type": "integer", "minimum": 0, "maximum": 1000}, "sheet": RANGE}, []),
    _action("unfreeze_panes", {"sheet": RANGE}, []),
    _action("rename_sheet", {"from": {"type": "string", "maxLength": 31}, "to": {"type": "string", "minLength": 1, "maxLength": 31}}, ["to"]),
    _action("delete_sheet", {"name": {"type": "string", "minLength": 1, "maxLength": 31}}, ["name"]),
    _action("sort_range", {"range": RANGE, "column": {"type": "integer", "minimum": 0, "maximum": 16383},
        "ascending": {"type": "boolean"}, "has_header": {"type": "boolean"}}, ["range", "column"]),
    _action("clear_range", {"range": RANGE, "target": {"type": "string", "enum": ["contents", "formats", "all"]}}, ["range"]),
])
ACTION_TYPES = {schema["properties"]["type"]["const"] for schema in ACTION_SCHEMAS}


def _validate_matrix(value: Any) -> bool:
    return (isinstance(value, list) and 1 <= len(value) <= 300 and
            all(isinstance(row, list) and 1 <= len(row) <= 30 and
                all(cell is None or isinstance(cell, (str, int, float, bool)) for cell in row) for row in value))


def _validate_action_values(action: dict[str, Any]) -> None:
    kind = action["type"]
    if "values" in action and not _validate_matrix(action["values"]):
        raise ValueError(f"Excel action {kind} has an invalid values matrix")
    for field in ("range", "start_cell", "sheet", "name", "from", "to"):
        if field in action and (not isinstance(action[field], str) or not action[field].strip() or len(action[field]) > 160):
            raise ValueError(f"Excel action {kind} has an invalid {field}")


def _matches_type(value: Any, expected: str) -> bool:
    if expected == "null": return value is None
    if expected == "boolean": return isinstance(value, bool)
    if expected == "integer": return isinstance(value, int) and not isinstance(value, bool)
    if expected == "number": return ((isinstance(value, int) and not isinstance(value, bool)) or
                                      (isinstance(value, float) and math.isfinite(value)))
    if expected == "string": return isinstance(value, str)
    if expected == "array": return isinstance(value, list)
    if expected == "object": return isinstance(value, dict)
    return False


def validate_schema(value: Any, schema: dict[str, Any], path: str = "value") -> None:
    if "oneOf" in schema:
        matches = 0
        for candidate in schema["oneOf"]:
            try:
                validate_schema(value, candidate, path)
                matches += 1
            except ValueError:
                pass
        if matches != 1: raise ValueError(f"{path} must match exactly one allowed schema")
        return
    if "const" in schema and value != schema["const"]: raise ValueError(f"{path} has the wrong constant value")
    expected = schema.get("type")
    if expected:
        choices = expected if isinstance(expected, list) else [expected]
        if not any(_matches_type(value, choice) for choice in choices): raise ValueError(f"{path} has the wrong type")
    if "enum" in schema and value not in schema["enum"]: raise ValueError(f"{path} is not an allowed value")
    if isinstance(value, str):
        if len(value) < schema.get("minLength", 0) or len(value) > schema.get("maxLength", len(value)): raise ValueError(f"{path} has invalid length")
        if schema.get("pattern") and not re.fullmatch(schema["pattern"], value): raise ValueError(f"{path} has invalid format")
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if (isinstance(value, float) and not math.isfinite(value)) or value < schema.get("minimum", value) or value > schema.get("maximum", value): raise ValueError(f"{path} is out of range")
    if isinstance(value, list):
        if len(value) < schema.get("minItems", 0) or len(value) > schema.get("maxItems", len(value)): raise ValueError(f"{path} has invalid item count")
        for index, item in enumerate(value): validate_schema(item, schema.get("items", {}), f"{path}[{index}]")
    if isinstance(value, dict):
        required = schema.get("required", [])
        missing = [field for field in required if field not in value]
        if missing: raise ValueError(f"{path} is missing {', '.join(missing)}")
        properties = schema.get("properties", {})
        if schema.get("additionalProperties") is False and set(value) - set(properties): raise ValueError(f"{path} has unknown fields")
        for key, item in value.items():
            if key in properties: validate_schema(item, properties[key], f"{path}.{key}")
EXCEL_RESPONSE_SCHEMA = {
    "name": "excel_response",
    "description": (
        "Finish an Excel turn by submitting a typed workbook-change proposal. "
        "This does not modify Excel. The task pane validates and previews it before Apply."
    ),
    "parameters": {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "request_id": {"type": "string", "pattern": ID_PATTERN, "maxLength": 128},
            "workbook_id": {"type": "string", "pattern": ID_PATTERN, "maxLength": 128},
            "conversation_id": {"type": "string", "pattern": ID_PATTERN, "maxLength": 128},
            "round": {"type": "integer", "minimum": 0, "maximum": 5},
            "message": {"type": "string", "maxLength": 4000},
            "actions": {
                "type": "array",
                "maxItems": MAX_ACTIONS,
                "items": {"oneOf": ACTION_SCHEMAS},
            },
        },
        "required": ["request_id", "workbook_id", "conversation_id", "round", "message", "actions"],
    },
}


def excel_response_available() -> bool:
    return True


def handle_excel_response(args: dict[str, Any], **_kwargs: Any) -> str:
    if not isinstance(args, dict):
        raise ValueError("excel_response arguments must be an object")
    validate_schema(args, EXCEL_RESPONSE_SCHEMA["parameters"], "excel_response")
    encoded = json.dumps(args, ensure_ascii=False, separators=(",", ":"), allow_nan=False).encode("utf-8")
    if len(encoded) > MAX_PAYLOAD_BYTES:
        raise ValueError("Excel proposal exceeds the 512000-byte limit")
    actions = args.get("actions")
    if not isinstance(actions, list) or len(actions) > MAX_ACTIONS:
        raise ValueError("Excel proposal actions are invalid or exceed the limit")
    for action in actions:
        if not isinstance(action, dict) or action.get("type") not in ACTION_TYPES:
            raise ValueError("Excel proposal contains an unknown action")
        schema = next(item for item in ACTION_SCHEMAS if item["properties"]["type"]["const"] == action["type"])
        allowed = set(schema["properties"])
        if set(action) - allowed or any(field not in action for field in schema["required"]):
            raise ValueError(f"Excel action {action.get('type')} has invalid or missing fields")
        _validate_action_values(action)
    capture_proposal(args)
    return "Proposal captured for task-pane validation. Do not claim it was applied; provide only the final summary."
