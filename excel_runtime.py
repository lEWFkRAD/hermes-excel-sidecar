"""Request correlation for the Excel platform and capture-only response tool."""

from __future__ import annotations

import asyncio
import threading
from dataclasses import dataclass
from typing import Any


@dataclass
class PendingExcelRequest:
    request_id: str
    workbook_id: str
    conversation_id: str
    round: int
    loop: asyncio.AbstractEventLoop
    proposal: asyncio.Future
    final_message: asyncio.Future
    proposal_captured: bool = False
    final_captured: bool = False
    protocol_error: str = ""


_pending: dict[str, PendingExcelRequest] = {}
_active_workbooks: dict[str, str] = {}
_lock = threading.RLock()


def _set_future(item: PendingExcelRequest, future: asyncio.Future, value: Any) -> None:
    try:
        same_loop = asyncio.get_running_loop() is item.loop
    except RuntimeError:
        same_loop = False
    if same_loop:
        future.set_result(value)
    else:
        item.loop.call_soon_threadsafe(future.set_result, value)


def open_request(*, request_id: str, workbook_id: str, conversation_id: str, round: int) -> PendingExcelRequest:
    loop = asyncio.get_running_loop()
    with _lock:
        if request_id in _pending:
            raise ValueError("duplicate request_id")
        if workbook_id in _active_workbooks:
            raise RuntimeError("workbook already has an active request")
        item = PendingExcelRequest(
            request_id=request_id,
            workbook_id=workbook_id,
            conversation_id=conversation_id,
            round=round,
            loop=loop,
            proposal=loop.create_future(),
            final_message=loop.create_future(),
        )
        _pending[request_id] = item
        _active_workbooks[workbook_id] = request_id
        return item


def get_request(request_id: str) -> PendingExcelRequest | None:
    with _lock:
        return _pending.get(request_id)


def capture_proposal(args: dict[str, Any]) -> None:
    request_id = str(args.get("request_id") or "")
    with _lock:
        item = _pending.get(request_id)
        if item is None:
            raise ValueError("unknown or expired request_id")
        if str(args.get("workbook_id") or "") != item.workbook_id:
            item.protocol_error = "workbook_id mismatch"
            raise ValueError("workbook_id mismatch")
        if str(args.get("conversation_id") or "") != item.conversation_id:
            item.protocol_error = "conversation_id mismatch"
            raise ValueError("conversation_id mismatch")
        if int(args.get("round", -1)) != item.round:
            item.protocol_error = "round mismatch"
            raise ValueError("round mismatch")
        if item.proposal_captured:
            item.protocol_error = "duplicate proposal capture"
            raise ValueError("proposal already captured")
        if item.final_captured:
            item.protocol_error = "proposal arrived after final delivery"
            raise ValueError("proposal arrived after final delivery")
        item.proposal_captured = True
    _set_future(item, item.proposal, args)


def capture_final(request_id: str, message: str) -> bool:
    with _lock:
        item = _pending.get(request_id)
        if item is None or item.final_captured or not item.proposal_captured:
            return False
        item.final_captured = True
    _set_future(item, item.final_message, str(message))
    return True


def capture_final_for_conversation(conversation_id: str, message: str) -> bool:
    with _lock:
        matches = [item.request_id for item in _pending.values()
                   if item.conversation_id == conversation_id]
    return len(matches) == 1 and capture_final(matches[0], message)


def close_request(request_id: str) -> None:
    with _lock:
        item = _pending.pop(request_id, None)
        if item and _active_workbooks.get(item.workbook_id) == request_id:
            _active_workbooks.pop(item.workbook_id, None)
    if item:
        for future in (item.proposal, item.final_message):
            if not future.done():
                item.loop.call_soon_threadsafe(future.cancel)
