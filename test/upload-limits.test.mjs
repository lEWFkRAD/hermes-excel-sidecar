import test from 'node:test';
import assert from 'node:assert/strict';
import http from 'node:http';
import { Readable } from 'node:stream';
import { readJson, validateUploadBatch } from '../broker/server.mjs';

test('chat reader accepts a synthetic request above the former 50 MiB ceiling', async () => {
  const block = 'x'.repeat(1024 * 1024);
  function* chunks() { yield Buffer.from('{"data":"'); for (let i = 0; i < 51; i++) yield Buffer.from(block); yield Buffer.from('"}'); }
  const req = Readable.from(chunks());
  const body = await readJson(req, 210 * 1024 * 1024);
  assert.equal(body.data.length, 51 * 1024 * 1024);
});

test('attachment validation uses encoded contents rather than claimed sizes', () => {
  const file = (n) => ({ size: 0, base64: Buffer.alloc(n).toString('base64') });
  const limits = { perFile: 100, totalLimit: 150 };
  assert.doesNotThrow(() => validateUploadBatch([file(29), file(71), file(29)], limits));
  assert.doesNotThrow(() => validateUploadBatch([file(100), file(50)], limits));
  assert.throws(() => validateUploadBatch([file(101)], limits), { statusCode: 413 });
  assert.throws(() => validateUploadBatch([file(80), file(80)], limits), { statusCode: 413 });
});

test('oversized fixed-length and chunked HTTP bodies return 413 without destroying the connection', async () => {
  const server = http.createServer(async (req, res) => {
    try { await readJson(req, 16); res.end('ok'); }
    catch (error) { res.writeHead(error.statusCode || 400); res.end(error.message); }
  });
  await new Promise((resolve) => server.listen(0, '127.0.0.1', resolve));
  const send = (headers, chunks) => new Promise((resolve, reject) => {
    const req = http.request({ hostname: '127.0.0.1', port: server.address().port, method: 'POST', headers }, (res) => {
      let body = ''; res.on('data', (chunk) => { body += chunk; });
      res.on('end', () => resolve({ status: res.statusCode, body }));
    });
    req.on('error', reject);
    for (const chunk of chunks) req.write(chunk);
    req.end();
  });
  try {
    for (const headers of [{ 'content-length': '32' }, {}]) {
      const result = await send(headers, ['x'.repeat(16), 'x'.repeat(16)]);
      assert.equal(result.status, 413);
      assert.match(result.body, /Request exceeds/);
    }
    assert.equal((await send({}, ['{}'])).status, 200);
  } finally {
    server.closeAllConnections();
    await new Promise((resolve) => server.close(resolve));
  }
});
