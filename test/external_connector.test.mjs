import test from "node:test";
import assert from "node:assert/strict";
import { ExternalConnector, normalizeConnectorRequest, isPrivateHost } from "../external_connector.mjs";

test("connector rejects non-HTTPS, private hosts, and query credentials", () => {
  assert.throws(() => normalizeConnectorRequest({ url: "http://example.com" }));
  assert.throws(() => normalizeConnectorRequest({ url: "https://127.0.0.1/x" }));
  assert.throws(() => normalizeConnectorRequest({ url: "https://example.com/x?token=secret" }));
  assert.equal(isPrivateHost("10.0.0.1"), true);
  assert.equal(isPrivateHost("example.com"), false);
});

test("connector is disabled by default and approval is mandatory", async () => {
  const connector = new ExternalConnector({ allowedOrigins: ["https://example.com"] });
  await assert.rejects(() => connector.request({ url: "https://example.com/catalog", purpose: "read", operation: "api_read" }), /disabled/);
  const enabled = new ExternalConnector({ enabled: true, allowedOrigins: ["https://example.com"] });
  await assert.rejects(() => enabled.request({ url: "https://example.com/catalog", purpose: "read" }), /approval/);
});

test("credential approval is separate and token is resolved at request time", async () => {
  const calls = [];
  const connector = new ExternalConnector({ enabled: true, allowedOrigins: ["https://example.com"], credentialResolver: async (scope) => { calls.push(scope); return "runtime-token"; } });
  await assert.rejects(() => connector.request({ url: "https://example.com/catalog", purpose: "read", credential_scope: "catalog.read" }, { originApproved: true }), /credential approval/);
  const originalFetch = globalThis.fetch;
  globalThis.fetch = async (_url, options) => ({ ok: true, status: 200, headers: new Headers({ "content-type": "application/json" }), text: async () => JSON.stringify({ result: true, token: "should-not-leak" }), options });
  try {
    const result = await connector.request({ url: "https://example.com/catalog", purpose: "read", credential_scope: "catalog.read" }, { originApproved: true, credentialApproved: true });
    assert.deepEqual(calls, ["catalog.read"]);
    assert.equal(result.data.token, "[REDACTED]");
  } finally { globalThis.fetch = originalFetch; }
});
