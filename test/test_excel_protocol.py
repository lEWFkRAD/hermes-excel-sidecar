from __future__ import annotations

import asyncio
import contextlib
import importlib.util
import json
import os
import pathlib
import sys
import tempfile
import threading
import types
import unittest
from dataclasses import dataclass
from unittest import mock

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


def install_gateway_stubs() -> None:
    """Install the tiny public gateway surface this plugin imports in bare CI."""

    gateway = types.ModuleType("gateway")
    gateway.__path__ = []
    config = types.ModuleType("gateway.config")
    platforms = types.ModuleType("gateway.platforms")
    platforms.__path__ = []
    base = types.ModuleType("gateway.platforms.base")

    class Platform(str):
        pass

    class PlatformConfig:
        pass

    class BasePlatformAdapter:
        pass

    class MessageEvent:
        def __init__(self, **values):
            self.__dict__.update(values)

    class MessageType:
        TEXT = "text"

    @dataclass
    class SendResult:
        success: bool
        message_id: str | None = None
        error: str | None = None

    config.Platform = Platform
    config.PlatformConfig = PlatformConfig
    base.BasePlatformAdapter = BasePlatformAdapter
    base.MessageEvent = MessageEvent
    base.MessageType = MessageType
    base.SendResult = SendResult
    sys.modules.update({
        "gateway": gateway,
        "gateway.config": config,
        "gateway.platforms": platforms,
        "gateway.platforms.base": base,
    })


def install_aiohttp_stub() -> None:
    """Provide json_response only when aiohttp is absent from a bare runner."""

    aiohttp = types.ModuleType("aiohttp")
    web = types.ModuleType("aiohttp.web")

    @dataclass
    class Response:
        body: bytes
        status: int

    def json_response(payload, status=200):
        return Response(json.dumps(payload).encode("utf-8"), status)

    web.json_response = json_response
    aiohttp.web = web
    sys.modules.update({"aiohttp": aiohttp, "aiohttp.web": web})


try:
    runtime = load("excel_runtime")
    tool = load("excel_tool")
    ownership = load("profile_ownership")
    adapter_mod = load("adapter")
except ModuleNotFoundError as error:  # pragma: no cover - exercised on bare CI runners
    if not (error.name or "").startswith("gateway"):
        raise
    install_gateway_stubs()
    sys.modules.pop(f"{PKG}.adapter", None)
    adapter_mod = load("adapter")

try:
    import aiohttp  # noqa: F401
except ModuleNotFoundError:  # pragma: no cover - exercised on bare CI runners
    install_aiohttp_stub()
policy = load("excel_policy")
plugin_entry = load("__init__")


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

    def test_nous_qwen_uses_auto_but_remains_typed_and_fail_closed(self):
        first = policy.excel_terminal_tool_policy(
            request=self.request(), platform="excel", api_call_count=1,
            provider="nous", model="qwen/qwen3.8-max",
        )["request"]
        self.assertEqual(first["tool_choice"], "auto")
        self.assertEqual(first["tools"][0]["function"]["name"], "excel_response")

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
    def __init__(self, body, token="secret-token", query=None):
        self._body = body
        self.headers = {"X-Excel-Token": token}
        self.query = query or {}

    async def json(self):
        return self._body


class AdapterTests(unittest.IsolatedAsyncioTestCase):
    def make_adapter(self, handler):
        adapter = object.__new__(adapter_mod.ExcelAdapter)
        owner = adapter_mod.AdapterOwnerIdentity(
            profile_name="finance",
            owner_fingerprint="sha256:" + "ab" * 32,
            bridge_port=8788,
        )
        adapter._bound_owner = owner
        adapter._load_active_owner_binding = lambda: (owner, "secret-token")
        adapter._timeout = 0.05
        adapter._background_tasks = set()
        adapter._request_tasks = {}
        adapter.build_source = lambda **kwargs: types.SimpleNamespace(**kwargs)
        adapter._message_handler = handler
        return adapter

    async def test_every_http_endpoint_revalidates_owner_binding(self):
        async def handler(_event): return None
        adapter = self.make_adapter(handler)
        original = adapter._load_active_owner_binding
        calls = 0

        def counted():
            nonlocal calls
            calls += 1
            return original()

        adapter._load_active_owner_binding = counted
        health = await adapter._handle_health(FakeRequest({}))
        activity = await adapter._handle_activity(FakeRequest(
            {}, query={"request_id": "request-activity-12345678"},
        ))
        cancel = await adapter._handle_cancel(FakeRequest({
            "request_id": "request-cancel-missing",
            "workbook_id": "workbook-cancel-missing",
        }))
        ingest = await adapter._handle_ingest(FakeRequest({}))
        self.assertEqual((health.status, activity.status, cancel.status, ingest.status), (200, 200, 200, 400))
        self.assertEqual(calls, 4)

    async def test_health_attests_owner_and_fails_closed_after_receipt_change(self):
        async def handler(_event): return None
        adapter = self.make_adapter(handler)
        healthy = await adapter._handle_health(FakeRequest({}))
        payload = json.loads(healthy.body)
        self.assertEqual(payload["profile_name"], "finance")
        self.assertEqual(payload["owner_fingerprint"], "sha256:" + "ab" * 32)
        self.assertEqual(payload["bridge_port"], 8788)

        adapter._load_active_owner_binding = lambda: (
            adapter_mod.AdapterOwnerIdentity(
                profile_name="personal",
                owner_fingerprint="sha256:" + "cd" * 32,
                bridge_port=8788,
            ),
            "secret-token",
        )
        rejected = await adapter._handle_health(FakeRequest({}))
        self.assertEqual(rejected.status, 503)
        self.assertNotIn(b"personal", rejected.body)
        self.assertNotIn(b"sha256", rejected.body)

    async def test_missing_or_corrupt_live_binding_precedes_token_auth(self):
        async def handler(_event): return None
        adapter = self.make_adapter(handler)
        adapter._load_active_owner_binding = mock.Mock(
            side_effect=adapter_mod.OwnerBindingError("receipt details must not leak")
        )
        response = await adapter._handle_ingest(FakeRequest(self.body("owner"), "wrong"))
        self.assertEqual(response.status, 503)
        self.assertNotIn(b"receipt details", response.body)

    async def test_connect_refuses_to_bind_without_current_owner(self):
        adapter = object.__new__(adapter_mod.ExcelAdapter)
        adapter._load_active_owner_binding = mock.Mock(
            side_effect=adapter_mod.OwnerBindingError("missing")
        )
        adapter._set_fatal_error = mock.Mock()
        self.assertFalse(await adapter.connect())
        adapter._set_fatal_error.assert_called_once_with(
            "excel_owner_binding_unavailable",
            "Excel sidecar ownership is unavailable for the active profile",
            retryable=False,
        )

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

    async def test_owner_swap_after_long_turn_withholds_proposal(self):
        async def handler(event):
            body = event.raw_message
            runtime.capture_proposal({"request_id": body["request_id"], "workbook_id": body["workbook_id"],
                "conversation_id": body["conversation_id"], "round": 0, "message": "ready", "actions": []})
            return "final"

        adapter = self.make_adapter(handler)
        valid = adapter._load_active_owner_binding
        adapter._load_active_owner_binding = mock.Mock(
            side_effect=[valid(), adapter_mod.OwnerBindingError("owner changed")]
        )
        response = await adapter._handle_ingest(FakeRequest(self.body("owner-swap")))
        self.assertEqual(response.status, 503)
        self.assertNotIn(b'"proposal"', response.body)
        self.assertEqual(adapter._load_active_owner_binding.call_count, 2)

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


class AdapterProfileBindingTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.base = pathlib.Path(self.temp.name)
        self.local_appdata = self.base / "Local"
        self.profile_home = self.local_appdata / "hermes" / "profiles" / "finance"
        self.environ = {
            "LOCALAPPDATA": str(self.local_appdata),
            "HERMES_HOME": str(self.profile_home),
            # A stale User-scope value must never authorize the adapter.
            "HERMES_EXCEL_INGEST_TOKEN": "f" * 64,
        }
        self.context = ownership.resolve_profile_context("finance", self.environ)
        self.receipt = ownership.OwnerReceipt.from_context(
            self.context,
            bridge_port=8788,
            plugin_version="0.2.0",
        )
        self.context.receipt_path.parent.mkdir(parents=True, exist_ok=True)
        self.context.receipt_path.write_text(json.dumps(self.receipt.to_dict()), encoding="utf-8")
        self.context.data_path.mkdir(parents=True, exist_ok=True)
        self.file_token = "a" * 64
        (self.context.data_path / ".ingest-token").write_bytes(self.file_token.encode("ascii"))

    def tearDown(self):
        self.temp.cleanup()

    def runtime_scope(self, *, name="finance", home=None):
        stack = contextlib.ExitStack()
        hermes_constants = types.ModuleType("hermes_constants")
        hermes_constants.get_hermes_home = lambda: home or self.profile_home
        hermes_cli = types.ModuleType("hermes_cli")
        hermes_cli.__path__ = []
        profiles = types.ModuleType("hermes_cli.profiles")
        profiles.get_active_profile_name = lambda: name
        stack.enter_context(mock.patch.dict(os.environ, self.environ, clear=False))
        stack.enter_context(mock.patch.dict(sys.modules, {
            "hermes_constants": hermes_constants,
            "hermes_cli": hermes_cli,
            "hermes_cli.profiles": profiles,
        }))
        return stack

    def test_call_time_official_profile_selects_file_token_and_ignores_user_env(self):
        with self.runtime_scope():
            owner, token = adapter_mod.load_active_owner_binding("finance")
        self.assertEqual(token, self.file_token)
        self.assertNotEqual(token, self.environ["HERMES_EXCEL_INGEST_TOKEN"])
        self.assertEqual(owner.profile_name, "finance")
        self.assertEqual(owner.bridge_port, 8788)
        self.assertEqual(owner.owner_fingerprint, ownership.owner_fingerprint(self.receipt))

    def test_token_file_is_exact_lowercase_hex_and_bounded(self):
        token_path = self.context.data_path / ".ingest-token"
        for value in (b"a" * 63, b"a" * 65, b"A" * 64, b"a" * 64 + b"\n"):
            with self.subTest(length=len(value), prefix=value[:1]):
                token_path.write_bytes(value)
                with self.runtime_scope(), self.assertRaises(adapter_mod.OwnerBindingError):
                    adapter_mod.load_active_owner_binding("finance")

    def test_missing_corrupt_and_other_owner_receipts_fail_before_token_read(self):
        receipt_path = self.context.receipt_path
        cases = (None, b"{not-json", b"{}")
        for raw in cases:
            with self.subTest(raw=raw):
                if raw is None:
                    receipt_path.unlink(missing_ok=True)
                else:
                    receipt_path.write_bytes(raw)
                with self.runtime_scope(), mock.patch.object(
                    adapter_mod, "_read_ingest_token"
                ) as read_token, self.assertRaises(adapter_mod.OwnerBindingError):
                    adapter_mod.load_active_owner_binding("finance")
                read_token.assert_not_called()

        other_home = self.local_appdata / "hermes" / "profiles" / "personal"
        other_env = dict(self.environ, HERMES_HOME=str(other_home))
        other_context = ownership.resolve_profile_context("personal", other_env)
        other_receipt = ownership.OwnerReceipt.from_context(
            other_context, bridge_port=8788, plugin_version="0.2.0"
        )
        receipt_path.write_text(json.dumps(other_receipt.to_dict()), encoding="utf-8")
        with self.runtime_scope(), mock.patch.object(
            adapter_mod, "_read_ingest_token"
        ) as read_token, self.assertRaises(adapter_mod.OwnerBindingError):
            adapter_mod.load_active_owner_binding("finance")
        read_token.assert_not_called()

    def test_registration_hint_and_call_time_home_must_still_agree(self):
        personal_home = self.local_appdata / "hermes" / "profiles" / "personal"
        for name, home, hint in (
            ("personal", personal_home, "finance"),
            ("finance", personal_home, "finance"),
        ):
            with self.subTest(name=name, home=home), self.runtime_scope(name=name, home=home), self.assertRaises(
                adapter_mod.OwnerBindingError
            ):
                adapter_mod.load_active_owner_binding(hint)

    def test_receipt_swap_during_token_read_fails_closed(self):
        changed = ownership.OwnerReceipt.from_context(
            self.context, bridge_port=8789, plugin_version="0.2.0"
        )
        with self.runtime_scope(), mock.patch.object(
            adapter_mod.profile_ownership,
            "load_owner_receipt",
            side_effect=[self.receipt, changed],
        ), mock.patch.object(adapter_mod, "_read_ingest_token", return_value=self.file_token):
            with self.assertRaises(adapter_mod.OwnerBindingError):
                adapter_mod.load_active_owner_binding("finance")

    def test_dependency_check_includes_the_bound_owner_probe(self):
        owner = adapter_mod.AdapterOwnerIdentity(
            profile_name="finance",
            owner_fingerprint="sha256:" + "ab" * 32,
            bridge_port=8788,
        )
        with mock.patch.object(
            adapter_mod, "load_active_owner_binding", return_value=(owner, self.file_token)
        ) as load_binding:
            self.assertTrue(adapter_mod.check_excel_requirements(profile_name="finance"))
        load_binding.assert_called_once_with("finance")

        with mock.patch.object(
            adapter_mod,
            "load_active_owner_binding",
            side_effect=adapter_mod.OwnerBindingError("other owner"),
        ):
            self.assertFalse(adapter_mod.check_excel_requirements(profile_name="finance"))

    def test_plugin_registration_binds_one_profile_into_factory_and_check(self):
        class Context:
            def __init__(self):
                self.profile_reads = 0
                self.platform = None

            @property
            def profile_name(self):
                self.profile_reads += 1
                return "finance"

            def register_cli_command(self, **_kwargs): pass
            def register_middleware(self, *_args, **_kwargs): pass
            def register_tool(self, **_kwargs): pass

            def register_platform(self, **kwargs):
                self.platform = kwargs

        context = Context()
        plugin_entry.register(context)
        self.assertEqual(context.profile_reads, 1)
        self.assertEqual(context.platform["adapter_factory"].keywords, {"profile_name": "finance"})
        self.assertEqual(context.platform["check_fn"].keywords, {"profile_name": "finance"})
        self.assertNotIn("required_env", context.platform)


if __name__ == "__main__":
    unittest.main()
