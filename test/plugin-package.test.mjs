import assert from "node:assert/strict";
import fs from "node:fs/promises";
import test from "node:test";

test("repository is an installable Hermes standalone plugin", async () => {
  const manifest = await fs.readFile("plugin.yaml", "utf8");
  assert.match(manifest, /^manifest_version: 1$/m);
  assert.match(manifest, /^name: hermes-excel-sidecar$/m);
  assert.match(manifest, /^kind: standalone$/m);
  const pluginVersion = manifest.match(/^version:\s*(\S+)$/m)?.[1];
  const packageMetadata = JSON.parse(await fs.readFile("package.json", "utf8"));
  assert.match(pluginVersion, /^\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?(?:\+[0-9A-Za-z.-]+)?$/);
  assert.equal(pluginVersion, packageMetadata.version);
  const changelog = await fs.readFile("CHANGELOG.md", "utf8");
  assert.match(changelog, new RegExp(`^## \\[${pluginVersion}\\]`, "m"));
  const entrypoint = await fs.readFile("__init__.py", "utf8");
  assert.match(entrypoint, /def register\(ctx\)/);
  assert.match(entrypoint, /register_cli_command/);
  await fs.access("after-install.md");
});
