#!/usr/bin/env node

import fs from "node:fs";

function fail(message) {
  console.error(message);
  process.exit(1);
}

const packageMetadata = JSON.parse(fs.readFileSync("package.json", "utf8"));
const packageVersion = packageMetadata.version;
const manifest = fs.readFileSync("plugin.yaml", "utf8");
const manifestMatch = manifest.match(/^version:\s*(\S+)\s*$/m);
const semver = /^\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?(?:\+[0-9A-Za-z.-]+)?$/;

if (!semver.test(packageVersion)) {
  fail(`package.json has an invalid SemVer version: ${packageVersion}`);
}
if (!manifestMatch) {
  fail("plugin.yaml does not declare a version");
}
if (manifestMatch[1] !== packageVersion) {
  fail(
    `Version mismatch: plugin.yaml=${manifestMatch[1]} package.json=${packageVersion}`,
  );
}

const changelog = fs.readFileSync("CHANGELOG.md", "utf8");
const escapedVersion = packageVersion.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
const releaseHeading = changelog.match(
  new RegExp(`^## \\[${escapedVersion}\\](?: - (Unreleased|\\d{4}-\\d{2}-\\d{2}))?$`, "m"),
);
if (!releaseHeading) {
  fail(`CHANGELOG.md does not contain a ${packageVersion} release section`);
}

const tagIndex = process.argv.indexOf("--tag");
if (tagIndex !== -1) {
  const tag = process.argv[tagIndex + 1];
  if (!tag) {
    fail("--tag requires a value");
  }
  if (tag !== `v${packageVersion}`) {
    fail(`Tag ${tag} does not match package version v${packageVersion}`);
  }
  if (!/^\d{4}-\d{2}-\d{2}$/.test(releaseHeading[1] || "")) {
    fail(
      `CHANGELOG.md must date the ${packageVersion} section before tag ${tag} can be released`,
    );
  }
}

process.stdout.write(`${packageVersion}\n`);
