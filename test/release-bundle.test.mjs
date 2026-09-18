import test from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { execFile } from "node:child_process";
import { promisify } from "node:util";

const run = promisify(execFile);
const root = path.resolve(import.meta.dirname, "..");

test("release bundle is reproducible and excludes runtime secrets", async () => {
  const dir = await mkdtemp(path.join(tmpdir(), "hermes-release-bundle-"));
  const first = path.join(dir, "one.zip");
  const second = path.join(dir, "two.zip");
  try {
    await run(process.execPath, [path.join(root, "scripts/build-release-bundle.mjs"), "--output", first, "--release", "test-release"], { cwd: root });
    await run(process.execPath, [path.join(root, "scripts/build-release-bundle.mjs"), "--output", second, "--release", "test-release"], { cwd: root });
    assert.deepEqual(await readFile(first), await readFile(second));
    const listing = (await run("unzip", ["-Z1", first])).stdout.split("\n").filter(Boolean);
    assert.ok(listing.includes("release-manifest.json"));
    assert.ok(listing.includes("ROLLBACK.md"));
    assert.equal(listing.some((entry) => /(^|\/)(?:data|uploads|exports|\.git|node_modules)(?:\/|$)|(?:\.env|\.bridge-token|\.ingest-token|\.key|\.pem|\.crt)$/.test(entry)), false);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});
