import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';

const source = fs.readFileSync(new URL('../taskpane.js', import.meta.url), 'utf8');
const sandbox = { window: { location: { origin: 'https://localhost:8788' } }, document: { querySelector: () => null, getElementById: () => ({}) }, Office: { onReady() {} }, Excel: {}, URL, crypto: { getRandomValues: (bytes) => bytes.fill(1) }, Uint8Array, Blob };
vm.createContext(sandbox);
vm.runInContext(source, sandbox);

test('external approval keys are session-scoped by origin, operation, and credential scope', () => {
  assert.equal(
    sandbox.externalApprovalKey({ origin: 'https://example.com', operation: 'read', credential_scope: '' }),
    'https://example.com|read|',
  );
  assert.notEqual(
    sandbox.externalApprovalKey({ origin: 'https://example.com', operation: 'read', credential_scope: '' }),
    sandbox.externalApprovalKey({ origin: 'https://other.example', operation: 'read', credential_scope: '' }),
  );
  assert.notEqual(
    sandbox.externalApprovalKey({ origin: 'https://example.com', operation: 'read', credential_scope: '' }),
    sandbox.externalApprovalKey({ origin: 'https://example.com', operation: 'read', credential_scope: 'cynteka.read' }),
  );
});

test('external approval requests reject non-HTTPS origins and non-read operations', () => {
  assert.throws(() => sandbox.normalizeExternalAccessRequest({ url: 'http://example.com', purpose: 'read' }), /HTTPS/);
  assert.throws(() => sandbox.normalizeExternalAccessRequest({ url: 'https://example.com', purpose: 'write', operation: 'write' }), /read-only/);
  assert.throws(() => sandbox.normalizeExternalAccessRequest({ url: 'https://0.0.0.0', purpose: 'read' }), /private/);
});
