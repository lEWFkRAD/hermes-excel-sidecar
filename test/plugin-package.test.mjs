import assert from "node:assert/strict";
import fs from "node:fs/promises";
import test from "node:test";

test("repository is an installable Hermes standalone plugin", async () => {
  const manifest = await fs.readFile("plugin.yaml", "utf8");
  assert.match(manifest, /^manifest_version: 1$/m);
  assert.match(manifest, /^name: hermes-excel-sidecar$/m);
  assert.match(manifest, /^kind: standalone$/m);
  const entrypoint = await fs.readFile("__init__.py", "utf8");
  assert.match(entrypoint, /def register\(ctx\)/);
  assert.match(entrypoint, /register_cli_command/);
  await fs.access("after-install.md");
});
