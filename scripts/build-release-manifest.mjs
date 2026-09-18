#!/usr/bin/env node
import { createHash } from "node:crypto";
import { readFile, writeFile } from "node:fs/promises";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import path from "node:path";

const exec = promisify(execFile);
const root = path.resolve(import.meta.dirname, "..");
const args = process.argv.slice(2);
const outputIndex = args.indexOf("--output");
const output = outputIndex >= 0 ? args[outputIndex + 1] : path.join(root, "release-manifest.json");
const releaseIndex = args.indexOf("--release");
const release = releaseIndex >= 0 ? args[releaseIndex + 1] : JSON.parse(await readFile(path.join(root, "package.json"), "utf8")).version;
if (!output || !release) throw new Error("Usage: build-release-manifest.mjs --output PATH [--release VALUE]");

async function gitCommit() {
  const { stdout } = await exec("git", ["rev-parse", "HEAD"], { cwd: root });
  return stdout.trim();
}
async function hashFiles(files) {
  const hash = createHash("sha256");
  for (const file of files) hash.update(file + "\0").update(await readFile(path.join(root, file)));
  return hash.digest("hex");
}

const manifest = {
  schema: 1,
  release: String(release),
  commit: await gitCommit(),
  client_hash: await hashFiles(["manifest.xml", "taskpane.html", "taskpane.css", "taskpane.js"]),
  bridge_hash: await hashFiles(["broker/server.mjs", "broker/formula-rebase.mjs", "broker/local_extract.py"]),
  adapter_hash: await hashFiles(["adapter.py", "excel_runtime.py", "excel_tool.py"]),
  protocol_version: "1",
  policy_version: await hashFiles(["excel_policy.py"]),
};
await writeFile(output, JSON.stringify(manifest, null, 2) + "\n", "utf8");
console.log(JSON.stringify({ output, release: manifest.release, commit: manifest.commit }));
