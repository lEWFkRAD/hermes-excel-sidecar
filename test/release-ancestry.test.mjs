import assert from "node:assert/strict";
import test from "node:test";

import { verifyReleaseAncestry } from "../scripts/verify-release-ancestry.mjs";

test("release ancestry dereferences the tag and accepts a main ancestor", () => {
  const calls = [];
  const runGit = (args) => {
    calls.push(args);
    if (args[0] === "rev-parse") return "abc123";
    if (args[0] === "merge-base") return "";
    throw new Error(`unexpected git command: ${args.join(" ")}`);
  };

  assert.equal(verifyReleaseAncestry("v0.2.0", "origin/main", runGit), "abc123");
  assert.deepEqual(calls, [
    ["rev-parse", "v0.2.0^{}"],
    ["merge-base", "--is-ancestor", "abc123", "origin/main"],
  ]);
});

test("release ancestry rejects a tag commit outside main", () => {
  const runGit = (args) => {
    if (args[0] === "rev-parse") return "def456";
    throw new Error("not an ancestor");
  };

  assert.throws(
    () => verifyReleaseAncestry("v0.2.0", "origin/main", runGit),
    /not reachable from origin\/main/,
  );
});
