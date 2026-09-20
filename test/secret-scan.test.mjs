import test from 'node:test';
import assert from 'node:assert/strict';
import { scanText, isTextFile } from '../scripts/check-secret-scan.mjs';

test('secret scan flags credential assignments but ignores redacted placeholders', () => {
  assert.deepEqual(scanText('const api_key = "live-secret-value";', 'sample.js'), [
    { file: 'sample.js', line: 1, kind: 'credential-assignment' },
  ]);
  assert.deepEqual(scanText('API_KEY=live-secret-value # test configuration', 'sample.env'), [
    { file: 'sample.env', line: 1, kind: 'credential-assignment' },
  ]);
  assert.deepEqual(scanText('$token = "live-secret-value"', 'sample.ps1'), [
    { file: 'sample.ps1', line: 1, kind: 'credential-assignment' },
  ]);
  assert.deepEqual(scanText('$token = live-secret-value', 'sample.ps1'), [
    { file: 'sample.ps1', line: 1, kind: 'credential-assignment' },
  ]);
  assert.deepEqual(scanText('$token = SECRET-LEAK', 'sample.ps1'), [
    { file: 'sample.ps1', line: 1, kind: 'credential-assignment' },
  ]);
  assert.deepEqual(scanText('const token = "[REDACTED]live-secret-value";', 'sample.js'), [
    { file: 'sample.js', line: 1, kind: 'credential-assignment' },
  ]);
  assert.deepEqual(scanText('const token = "<present>live-secret-value";', 'sample.js'), [
    { file: 'sample.js', line: 1, kind: 'credential-assignment' },
  ]);
  assert.deepEqual(scanText('$token = Ensure-BridgeToken', 'sample.ps1'), []);
  assert.deepEqual(scanText('"api_key": "live-secret-value",', 'sample.json'), [
    { file: 'sample.json', line: 1, kind: 'credential-assignment' },
  ]);
  assert.deepEqual(scanText('const token = "[REDACTED]";', 'sample.js'), []);
  assert.deepEqual(scanText('const token = process.env.TOKEN;', 'sample.js'), []);
});

test('secret scan includes Windows launcher and PowerShell text formats', () => {
  assert.equal(isTextFile('install/profile-ownership.psm1'), true);
  assert.equal(isTextFile('install/run-bridge.cmd.template'), true);
  assert.equal(isTextFile('service/bridge-service.cmd'), true);
  assert.equal(isTextFile('.env'), true);
  assert.equal(isTextFile('node_modules/ignore.js'), false);
});

test('secret scan flags GitHub and OpenAI token shapes', () => {
  const text = 'ghp_abcdefghijklmnopqrstuvwxyz1234567890\nsk-proj-abcdefghijklmnopqrstuvwxyz1234567890';
  const findings = scanText(text, 'sample.txt');
  assert.equal(findings.length, 2);
  assert.deepEqual(findings.map((item) => item.kind), ['github-token', 'openai-token']);
});
