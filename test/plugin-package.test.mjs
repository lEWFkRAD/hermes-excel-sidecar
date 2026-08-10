import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
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
  assert.doesNotMatch(
    manifest,
    /HERMES_EXCEL_(?:INGEST|BRIDGE)_TOKEN/,
    "Generated bridge secrets are injected from profile-owned files, never declared as plugin env inputs",
  );
  await fs.access("after-install.md");
});

test("git archive excludes ownership, secrets, and runtime artifacts", (context) => {
  // Canonical deployment and release archives deliberately share Git's
  // tracked-file manifest. Rejecting these paths here protects both surfaces.
  let trackedManifest;
  try {
    trackedManifest = execFileSync("git", ["ls-files", "-z"], {
      encoding: "utf8",
      stdio: ["ignore", "pipe", "pipe"],
    });
  } catch (error) {
    const stderr = error?.stderr?.toString() || "";
    if (/not a git repository/i.test(stderr)) {
      context.skip("archive extraction has no Git metadata; source validation guards its manifest");
      return;
    }
    throw error;
  }

  const trackedPaths = trackedManifest
    .split("\0")
    .filter(Boolean)
    .map((entry) => entry.replaceAll("\\", "/"));

  const forbidden = [
    {
      kind: "owner or installation receipt",
      pattern: /(^|\/)(?:owner\.pid|(?:[^/]+[-_.])?receipt(?:\.[^/]+)?|(?:[^/]+[-_.])?owner(?:ship)?(?:[-_.](?:receipt|v\d+))?\.(?:json|ya?ml|txt))$/i,
    },
    {
      kind: "token file",
      pattern: /(^|\/)(?:\.?(?:bridge|ingest|api|auth|access|refresh)[-_.]?token|\.?tokens?)(?:\.[^/]*)?$/i,
    },
    {
      kind: "environment file",
      pattern: /(^|\/)\.env(?:\.[^/]*)?$/i,
    },
    {
      kind: "TLS key or certificate material",
      pattern: /\.(?:key|pem|pfx|p12|crt|cer|jks)$/i,
    },
    {
      kind: "upload, export, or log artifact",
      pattern: /(^|\/)(?:uploads?|exports?|logs?)(?:\/|$)|\.log$/i,
    },
    {
      kind: "workbook or data artifact",
      pattern: /\.(?:xlsx|xlsm|xlsb|xls|ods|csv|tsv|parquet|sqlite3?|db|pdf|docx?|pptx?)$/i,
    },
    {
      kind: "runtime artifact",
      pattern: /(^|\/)(?:data|runtime|state|cache|tmp|temp|node_modules|__pycache__|\.pytest_cache|dist)(?:\/|$)|^(?:hermes-)?config\.ya?ml$|(^|\/)(?:run-bridge\.(?:cmd|vbs)|bridge-service\.vbs|bridge-supervisor\.lock|[^/]+\.(?:pid|sock))$/i,
    },
  ];

  const representativeRuntimePaths = [
    "shared/profile-owner.json",
    "shared/excel-sidecar/owner-v1.json",
    "data/.bridge-token",
    ".env.local",
    "tls/localhost.key",
    "uploads/customer-workbook.pdf",
    "exports/customer-workbook.csv",
    "logs/bridge.log",
    "workbooks/customer.xlsx",
    "data/state.json",
    "runtime/bridge.pid",
    "run-bridge.cmd",
  ];
  for (const entry of representativeRuntimePaths) {
    assert.ok(
      forbidden.some(({ pattern }) => pattern.test(entry)),
      `Archive hygiene policy must reject ${entry}`,
    );
  }

  const violations = trackedPaths.flatMap((entry) =>
    forbidden
      .filter(({ pattern }) => pattern.test(entry))
      .map(({ kind }) => `${kind}: ${entry}`),
  );

  assert.deepEqual(
    violations,
    [],
    `Sensitive files would enter git archive/canonical deploy:\n${violations.join("\n")}`,
  );
});
