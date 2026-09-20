#!/usr/bin/env node

import { readFileSync, readdirSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const TEXT_EXTENSIONS = new Set(['.js', '.mjs', '.cjs', '.py', '.json', '.yaml', '.yml', '.md', '.txt', '.ps1', '.psm1', '.xml', '.html', '.css', '.sh', '.cmd', '.vbs']);
const TEXT_SUFFIXES = ['.cmd.template', '.vbs.template'];
const TEXT_FILENAMES = new Set(['.env']);
const IGNORED_DIRECTORIES = new Set(['.git', '.hermes', 'dist', 'node_modules', 'test']);
const IGNORED_PATHS = new Set(['package-lock.json']);

export function isTextFile(relativePath) {
  const normalized = String(relativePath).replaceAll('\\', '/');
  if (normalized.split('/').some((part) => IGNORED_DIRECTORIES.has(part)) || IGNORED_PATHS.has(normalized)) return false;
  return TEXT_FILENAMES.has(path.basename(normalized)) || TEXT_EXTENSIONS.has(path.extname(normalized).toLowerCase()) || TEXT_SUFFIXES.some((suffix) => normalized.endsWith(suffix));
}

const SAFE_DYNAMIC_POWERSHELL_COMMANDS = new Set(['Ensure-BridgeToken', 'Ensure-IngestToken']);

function isRedactedOrEnvironmentReference(value) {
  return value === '[REDACTED]' || value === '<present>' || /^(?:process\.env\.|env:|%[A-Za-z_][A-Za-z0-9_]*%|\$\{?[A-Za-z_][A-Za-z0-9_]*)/i.test(value);
}

export function scanText(text, file = '') {
  const findings = [];
  const lines = String(text).split(/\r?\n/);
  for (let index = 0; index < lines.length; index += 1) {
    const line = lines[index];
    const assignment = line.match(/(?:\b(?:api[_-]?key|secret|password|passwd|token)\b\s*[:=]\s*(?:"([^"]{8,})"|'([^']{8,})'|([A-Za-z0-9_-]{8,}))\s*;?\s*(?:[#;].*)?$|["'](?:api[_-]?key|secret|password|passwd|token)["']\s*:\s*(?:"([^"]{8,})"|'([^']{8,})')\s*[,}])/i);
    const value = assignment?.[1] || assignment?.[2] || assignment?.[3] || assignment?.[4] || assignment?.[5] || "";
    const isDynamicPowerShellCommand = /^\s*\$/.test(line) && Boolean(assignment?.[3]) && SAFE_DYNAMIC_POWERSHELL_COMMANDS.has(value);
    if (assignment && !isDynamicPowerShellCommand && !isRedactedOrEnvironmentReference(value)) {
      findings.push({ file, line: index + 1, kind: 'credential-assignment' });
    }
    if (/\bgh[pousr]_[A-Za-z0-9_]{20,}\b/.test(line)) findings.push({ file, line: index + 1, kind: 'github-token' });
    if (/\bsk-(?:proj-)?[A-Za-z0-9_-]{20,}\b/.test(line)) findings.push({ file, line: index + 1, kind: 'openai-token' });
  }
  return findings;
}

function textFiles(directory = ROOT, prefix = '') {
  const files = [];
  for (const entry of readdirSync(directory, { withFileTypes: true })) {
    if (IGNORED_DIRECTORIES.has(entry.name)) continue;
    const relativePath = prefix ? `${prefix}/${entry.name}` : entry.name;
    const absolutePath = path.join(directory, entry.name);
    if (entry.isDirectory()) files.push(...textFiles(absolutePath, relativePath));
    else if (entry.isFile() && isTextFile(relativePath)) files.push(relativePath);
  }
  return files;
}

function main() {
  const findings = [];
  for (const relativePath of textFiles()) {
    findings.push(...scanText(readFileSync(path.join(ROOT, relativePath), 'utf8'), relativePath));
  }
  if (findings.length) {
    for (const finding of findings) console.error(`${finding.file}:${finding.line}: ${finding.kind}`);
    process.exitCode = 1;
    return;
  }
  console.log('Secret scan passed.');
}

if (process.argv[1] && import.meta.url === pathToFileURL(path.resolve(process.argv[1])).href) main();
