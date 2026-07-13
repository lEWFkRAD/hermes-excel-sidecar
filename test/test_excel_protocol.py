from __future__ import annotations

import asyncio
import importlib.util
import pathlib
import sys
import threading
import types
import unittest

ROOT = pathlib.Path(__file__).resolve().parents[1]
PKG = "excel_sidecar_testpkg"
package = types.ModuleType(PKG)
package.__path__ = [str(ROOT)]
sys.modules.setdefault(PKG, package)


def load(name: str):
    full = f"{PKG}.{name}"
    spec = importlib.util.spec_from_file_location(full, ROOT / f"{name}.py")
    module = importlib.util.module_from_spec(spec)
    sys.modules[full] = module
    assert spec.loader
    spec.loader.exec_module(module)
    return module


try:
    runtime = load("excel_runtime")
    tool = load("excel_tool")
    adapter_mod = load("adapter")
except ModuleNotFoundError as error:  # pragma: no cover - bare checkout without the Hermes runtime
    raise unittest.SkipTest(f"hermes gateway runtime not importable here: {error}")
policy = load("excel_policy")


class PolicyTests(unittest.TestCase):
    def request(self):
        return {"messages": [], "tools": [{"type": "function", "function": {"name": "excel_response"}}]}

    def test_excel_first_call_forces_named_tool_then_disables_tools(self):
        original = self.request()
        first = policy.excel_terminal_tool_policy(request=original, platform="excel", api_call_count=1)["request"]
        second = policy.excel_terminal_tool_policy(request=original, platform="excel", api_call_count=2)["request"]
        self.assertEqual(first["tool_choice"]["function"]["name"], "excel_response")
        self.assertEqual(second["tool_choice"], "none")
        self.assertNotIn("tool_choice", original)

    def test_non_excel_unchanged_and_missing_tool_fails_closed(self):
        self.assertIsNone(policy.excel_terminal_tool_policy(request=self.request(), platform="kindle", api_call_count=1))
        result = policy.excel_terminal_tool_policy(request={"tools": []}, platform="excel", api_call_count=1)
        self.assertEqual(result["request"]["tool_choice"]["function"]["name"], "__excel_response_unavailable__")


class RuntimeTests(unittest.IsolatedAsyncioTestCase):
    async def asyncTearDown(self):
        for request_id in list(runtime._pending):
            runtime.close_request(request_id)
        await asyncio.sleep(0)

    async def test_correlation_capture_and_final(self):
        item = runtime.open_request(request_id="request-12345678", workbook_id="workbook-12345678",
                                    conversation_id="conversation-12345678", round=1)
        args = {"request_id": item.request_id, "workbook_id": item.workbook_id,
                "conversation_id": item.conversation_id, "round": 1, "message": "Ready", "actions": []}
        runtime.capture_proposal(args)
        self.assertTrue(runtime.capture_final(item.request_id, "Final"))
        self.assertEqual(await item.proposal, args)
        self.assertEqual(await item.final_message, "Final")

    async def test_wrong_correlation_and_final_before_proposal_rejected(self):
        item = runtime.open_request(request_id="request-abcdefgh", workbook_id="workbook-abcdefgh",
                                    conversation_id="conversation-abcdefgh", round=0)
        self.assertFalse(runtime.capture_final(item.request_id, "too soon"))
        with self.assertRaisesRegex(ValueError, "workbook_id mismatch"):
            runtime.capture_proposal({"request_id": item.request_id, "workbook_id": "wrong-workbook",
                "conversation_id": item.conversation_id, "round": 0})

    async def test_same_workbook_and_duplicate_request_rejected(self):
        runtime.open_request(request_id="request-one-123", workbook_id="workbook-one-123",
                             conversation_id="conversation-one-123", round=0)
        with self.assertRaisesRegex(RuntimeError, "active request"):
            runtime.open_request(request_id="request-two-123", workbook_id="workbook-one-123",
                                 conversation_id="conversation-two-123", round=0)
        with self.assertRaisesRegex(ValueError, "duplicate"):
            runtime.open_request(request_id="request-one-123", workbook_id="workbook-two-123",
                                 conversation_id="conversation-two-123", round=0)

    async def test_concurrent_double_capture_has_one_winner(self):
        item = runtime.open_request(request_id="request-race-123", workbook_id="workbook-race-123",
                                    conversation_id="conversation-race-123", round=0)
        args = {"request_id": item.request_id, "workbook_id": item.workbook_id,
                "conversation_id": item.conversation_id, "round": 0, "message": "x", "actions": []}
        outcomes = []
        def capture():
            try:
                runtime.capture_proposal(args)
                outcomes.append("ok")
            except ValueError:
                outcomes.append("rejected")
        threads = [threading.Thread(target=capture) for _ in range(2)]
        for thread in threads: thread.start()
        for thread in threads: thread.join()
        await asyncio.sleep(0)
        self.assertCountEqual(outcomes, ["ok", "rejected"])


class ToolValidationTests(unittest.TestCase):
    def test_schema_is_strict_union(self):
        actions = tool.EXCEL_RESPONSE_SCHEMA["parameters"]["properties"]["actions"]
        self.assertIn("oneOf", actions["items"])
        self.assertTrue(all(schema["additionalProperties"] is False for schema in actions["items"]["oneOf"]))

    def test_unknown_missing_and_oversize_actions_rejected_before_capture(self):
        base = {"request_id": "request-tool-123", "workbook_id": "workbook-tool-123",
                "conversation_id": "conversation-tool-123", "round": 0, "message": "x"}
        for actions in ([{"type": "arbitrary_code"}], [{"type": "write_cells"}],
                        [{"type": "write_cells", "values": [[1] * 31]}],
                        [{"type": "read_range", "range": "A1", "extra": True}]):
            with self.subTest(actions=actions), self.assertRaises(ValueError):
                tool.handle_excel_response({**base, "actions": actions})

    def test_numeric_enum_identity_and_message_constraints(self):
        base = {"request_id": "request-valid-123", "workbook_id": "workbook-valid-123",
                "conversation_id": "conversation-valid-123", "round": 0, "message": "x", "actions": []}
        invalid = [
            {**base, "round": 6},
            {**base, "request_id": "bad"},
            {**base, "message": "x" * 4001},
            {**base, "actions": [{"type": "delete_rows", "at": 1, "count": -1}]},
            {**base, "actions": [{"type": "set_column_width", "range": "A:A", "width": float("nan")}]},
            {**base, "actions": [{"type": "clear_range", "range": "A1", "target": "evil"}]},
            {**base, "actions": [{"type": "sort_range", "range": "A1:B2", "column": True}]},
            {**base, "actions": [{"type": "set_column_width", "range": "A:A", "width": 10**400}]},
        ]
        for payload in invalid:
            with self.subTest(payload=payload), self.assertRaises(ValueError):
                tool.handle_excel_response(payload)

    def test_positive_style_array_and_column_letter_contract(self):
        tool.validate_schema({"type": "format_cells", "range": "A1:B2", "style": ["header", "total-row"]},
                             next(s for s in tool.ACTION_SCHEMAS if s["properties"]["type"]["const"] == "format_cells"))
        tool.validate_schema({"type": "insert_columns", "at": "E", "count": 2},
                             next(s for s in tool.ACTION_SCHEMAS if s["properties"]["type"]["const"] == "insert_columns"))


class FakeRequest:
    def __init__(self, body, token="secret-token"):
        self._body = body
        self.headers = {"X-Excel-Token": token}

    async def json(self):
        return self._body


class AdapterTests(unittest.IsolatedAsyncioTestCase):
    def make_adapter(self, handler):
        adapter = object.__new__(adapter_mod.ExcelAdapter)
        adapter._token = "secret-token"
        adapter._timeout = 0.05
        adapter._background_tasks = set()
        adapter._request_tasks = {}
        adapter.build_source = lambda **kwargs: types.SimpleNamespace(**kwargs)
        adapter._message_handler = handler
        return adapter

    @staticmethod
    def body(suffix="success"):
        return {"request_id": f"request-{suffix}-12345678", "workbook_id": f"workbook-{suffix}-12345678",
                "conversation_id": f"conversation-{suffix}-12345678", "round": 0, "prompt": "Build a table", "context": {}}

    async def asyncTearDown(self):
        for request_id in list(runtime._pending): runtime.close_request(request_id)
        await asyncio.sleep(0)

    async def test_immediate_proposal_and_final_completion(self):
        async def handler(event):
            body = event.raw_message
            runtime.capture_proposal({"request_id": body["request_id"], "workbook_id": body["workbook_id"],
                "conversation_id": body["conversation_id"], "round": 0, "message": "ready", "actions": []})
            return "final"
        response = await self.make_adapter(handler)._handle_ingest(FakeRequest(self.body()))
        self.assertEqual(response.status, 200)
        self.assertIn(b'"source": "hermes-platform"', response.body)

    async def test_agent_finishes_without_protocol_is_immediate_error(self):
        async def handler(_event): return None
        response = await self.make_adapter(handler)._handle_ingest(FakeRequest(self.body("missing")))
        self.assertEqual(response.status, 502)

    async def test_wrong_correlation_then_correct_capture_fails_whole_request(self):
        async def handler(event):
            body = event.raw_message
            bad = {"request_id": body["request_id"], "workbook_id": "workbook-wrong-12345678",
                   "conversation_id": body["conversation_id"], "round": 0, "message": "bad", "actions": []}
            with self.assertRaisesRegex(ValueError, "workbook_id mismatch"):
                runtime.capture_proposal(bad)
            runtime.capture_proposal({"request_id": body["request_id"], "workbook_id": body["workbook_id"],
                "conversation_id": body["conversation_id"], "round": 0, "message": "ready", "actions": []})
            return "final"
        response = await self.make_adapter(handler)._handle_ingest(FakeRequest(self.body("correlation")))
        self.assertEqual(response.status, 502)
        self.assertIn(b"workbook_id mismatch", response.body)

    async def test_timeout_cleans_runtime_and_task(self):
        async def handler(_event): await asyncio.sleep(60)
        adapter = self.make_adapter(handler)
        body = self.body("timeout")
        response = await adapter._handle_ingest(FakeRequest(body))
        self.assertEqual(response.status, 504)
        self.assertNotIn(body["request_id"], runtime._pending)
        self.assertFalse(adapter._request_tasks)

    async def test_auth_and_round_validation(self):
        async def handler(_event): return None
        adapter = self.make_adapter(handler)
        self.assertEqual((await adapter._handle_ingest(FakeRequest(self.body("auth"), "wrong"))).status, 401)
        bad = self.body("round")
        bad["round"] = 6
        self.assertEqual((await adapter._handle_ingest(FakeRequest(bad))).status, 400)

    async def test_send_requires_final_notify_after_proposal(self):
        async def handler(_event): return None
        adapter = self.make_adapter(handler)
        item = runtime.open_request(request_id="request-send-123", workbook_id="workbook-send-123",
                                    conversation_id="conversation-send-123", round=0)
        preview = await adapter.send(item.request_id, "preview", metadata={"notify": False})
        self.assertFalse(preview.success)
        early = await adapter.send(item.request_id, "early", metadata={"notify": True})
        self.assertFalse(early.success)
        runtime.capture_proposal({"request_id": item.request_id, "workbook_id": item.workbook_id,
            "conversation_id": item.conversation_id, "round": 0, "message": "ready", "actions": []})
        final = await adapter.send(item.request_id, "final", metadata={"notify": True})
        duplicate = await adapter.send(item.request_id, "duplicate", metadata={"notify": True})
        self.assertTrue(final.success)
        self.assertFalse(duplicate.success)
        self.assertEqual(await item.final_message, "final")

    async def test_cancel_releases_workbook_and_is_idempotent(self):
        started = asyncio.Event()
        async def handler(_event):
            started.set()
            await asyncio.sleep(60)
        adapter = self.make_adapter(handler)
        body = self.body("cancel")
        ingest = asyncio.create_task(adapter._handle_ingest(FakeRequest(body)))
        await started.wait()
        canceled = await adapter._handle_cancel(FakeRequest({"request_id": body["request_id"],
                                                              "workbook_id": body["workbook_id"]}))
        self.assertEqual(canceled.status, 200)
        response = await ingest
        self.assertEqual(response.status, 499)
        self.assertNotIn(body["request_id"], runtime._pending)
        replacement = runtime.open_request(request_id="request-after-cancel", workbook_id=body["workbook_id"],
                                           conversation_id=body["conversation_id"], round=0)
        runtime.close_request(replacement.request_id)
        again = await adapter._handle_cancel(FakeRequest({"request_id": body["request_id"],
                                                           "workbook_id": body["workbook_id"]}))
        self.assertEqual(again.status, 200)

    async def test_two_rounds_use_stable_conversation_chat_id(self):
        seen = []
        async def handler(event):
            seen.append((event.source.chat_id, event.message_id))
            body = event.raw_message
            runtime.capture_proposal({"request_id": body["request_id"], "workbook_id": body["workbook_id"],
                "conversation_id": body["conversation_id"], "round": body["round"], "message": "ready", "actions": []})
            return "final"
        adapter = self.make_adapter(handler)
        first = self.body("round-one")
        second = dict(first, request_id="request-round-two-12345678", round=1)
        self.assertEqual((await adapter._handle_ingest(FakeRequest(first))).status, 200)
        self.assertEqual((await adapter._handle_ingest(FakeRequest(second))).status, 200)
        self.assertEqual(seen[0][0], seen[1][0])
        self.assertNotEqual(seen[0][1], seen[1][1])


if __name__ == "__main__":
    unittest.main()
