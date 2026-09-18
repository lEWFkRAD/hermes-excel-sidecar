import test from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { execFile } from "node:child_process";
import { promisify } from "node:util";

const run = promisify(execFile);
const root = path.resolve(import.meta.dirname, "..");

test("release manifest emits canonical non-secret component metadata", async () => {
  const dir = await mkdtemp(path.join(tmpdir(), "hermes-release-manifest-"));
  const output = path.join(dir, "release-manifest.json");
  try {
    await run(process.execPath, [path.join(root, "scripts/build-release-manifest.mjs"), "--output", output, "--release", "test-release"], { cwd: root });
    const manifest = JSON.parse(await readFile(output, "utf8"));
    assert.equal(manifest.release, "test-release");
    assert.match(manifest.commit, /^[0-9a-f]{40}$/);
    for (const key of ["client_hash", "bridge_hash", "adapter_hash", "protocol_version", "policy_version"]) assert.ok(manifest[key]);
    assert.doesNotMatch(JSON.stringify(manifest), /token|password|secret|authorization/i);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});
