#!/usr/bin/env node
import { createHash } from "node:crypto";
import { readFile, readdir, stat, writeFile, mkdir, rm } from "node:fs/promises";
import path from "node:path";
import { execFile } from "node:child_process";
import { promisify } from "node:util";

const exec = promisify(execFile);
const root = path.resolve(import.meta.dirname, "..");
const args = process.argv.slice(2);
const outIndex = args.indexOf("--output");
const output = outIndex >= 0 ? path.resolve(args[outIndex + 1]) : path.resolve(root, "dist/hermes-excel-sidecar.zip");
const releaseIndex = args.indexOf("--release");
const release = releaseIndex >= 0 ? args[releaseIndex + 1] : JSON.parse(await readFile(path.join(root, "package.json"), "utf8")).version;
if (!output || !release) throw new Error("Usage: build-release-bundle.mjs --output PATH [--release VALUE]");

const includedRoots = ["manifest.xml", "plugin.yaml", "package.json", "package-lock.json", "README.md", "PROTOCOL.md", "SECURITY.md", "RELEASING.md", "CHANGELOG.md", "LICENSE", "taskpane.html", "taskpane.css", "taskpane.js", "broker", "install", "scripts", "test", "assets", "adapter.py", "excel_runtime.py", "excel_policy.py", "excel_tool.py", "profile_ownership.py", "remote_mode.py", "cli.py", "__init__.py"];
const forbidden = /(^|\/)(?:\.git|node_modules|data|uploads|exports|__pycache__|\.pytest_cache|dist)(?:\/|$)|(?:^|\/)(?:\.env|\.bridge-token|\.ingest-token|.*\.(?:key|pem|pfx|crt))$/i;
function dosDateTime() { return { date: 0x0021, time: 0 }; }
function u16(n) { const b = Buffer.alloc(2); b.writeUInt16LE(n); return b; }
function u32(n) { const b = Buffer.alloc(4); b.writeUInt32LE(n >>> 0); return b; }
function crc32(buffer) { let c = 0xffffffff; for (const byte of buffer) { c ^= byte; for (let i = 0; i < 8; i++) c = (c >>> 1) ^ (c & 1 ? 0xedb88320 : 0); } return (c ^ 0xffffffff) >>> 0; }
function zipEntry(name, data, offset) {
  const nameBytes = Buffer.from(name);
  const { date, time } = dosDateTime();
  const crc = crc32(data);
  const local = Buffer.concat([Buffer.from([0x50, 0x4b, 3, 4]), u16(20), u16(0), u16(0), u16(time), u16(date), u32(crc), u32(data.length), u32(data.length), u16(nameBytes.length), u16(0), nameBytes]);
  const central = Buffer.concat([Buffer.from([0x50, 0x4b, 1, 2]), u16(20), u16(20), u16(0), u16(0), u16(time), u16(date), u32(crc), u32(data.length), u32(data.length), u16(nameBytes.length), u16(0), u16(0), u16(0), u16(0), u32(0), u32(offset), nameBytes]);
  return { local, central, data };
}
async function walk(relative) {
  const absolute = path.join(root, relative);
  const info = await stat(absolute);
  if (info.isFile()) return [relative.replaceAll(path.sep, "/")];
  const names = (await readdir(absolute)).sort();
  const files = [];
  for (const name of names) { const child = path.posix.join(relative.replaceAll(path.sep, "/"), name); if (!forbidden.test(child)) files.push(...await walk(child)); }
  return files;
}
const files = [...new Set((await Promise.all(includedRoots.filter((entry) => !forbidden.test(entry)).map(walk))).flat())].sort();
const manifestPath = path.join(root, ".release-manifest.tmp.json");
try {
  await exec(process.execPath, [path.join(root, "scripts/build-release-manifest.mjs"), "--output", manifestPath, "--release", String(release)], { cwd: root });
  files.push("release-manifest.json", "ROLLBACK.md");
  const entries = [];
  let offset = 0;
  for (const name of files) {
    const data = name === "release-manifest.json" ? await readFile(manifestPath) : name === "ROLLBACK.md" ? Buffer.from("Rollback: retain the previous verified bundle and restore it with the normal installer.\n") : await readFile(path.join(root, name));
    const entry = zipEntry(name, data, offset); entries.push(entry); offset += entry.local.length + entry.data.length;
  }
  const central = Buffer.concat(entries.map((entry) => entry.central));
  const archive = Buffer.concat([...entries.flatMap((entry) => [entry.local, entry.data]), central, Buffer.from([0x50, 0x4b, 5, 6]), u16(0), u16(0), u16(entries.length), u16(entries.length), u32(central.length), u32(offset), u16(0)]);
  await mkdir(path.dirname(output), { recursive: true });
  await writeFile(output, archive);
  console.log(JSON.stringify({ output, release: String(release), entries: files.length, sha256: createHash("sha256").update(archive).digest("hex") }));
} finally { await rm(manifestPath, { force: true }); }
