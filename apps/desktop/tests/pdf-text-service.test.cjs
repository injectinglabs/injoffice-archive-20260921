const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const os = require('node:os');
const path = require('node:path');
const { createHash } = require('node:crypto');
const { replacePdfText } = require('../electron/pdf-text-service.cjs');

function request(bytes, extra = {}) {
  return {
    bytes,
    revision: createHash('sha256').update(bytes).digest('hex'),
    page: 1,
    rect: [0, 0, 10, 10],
    oldText: 'old',
    newText: 'new',
    ...extra,
  };
}

test('PDF text host validates input, serializes, times out, and returns worker bytes', async t => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'injoffice-pdf-text-service-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const workerPath = path.join(directory, 'worker.cjs');
  await fs.writeFile(workerPath, `const {parentPort,workerData}=require('node:worker_threads');const mode=Buffer.from(workerData.bytes).toString();if(mode==='hang')setInterval(()=>{},1000);else if(mode==='empty')parentPort.postMessage({bytes:new Uint8Array()});else parentPort.postMessage({bytes:Buffer.from('%PDF-out')});`);
  const bytes = Buffer.from('okay');
  await assert.rejects(replacePdfText({ ...request(bytes), revision: 'not-a-hash' }, workerPath), /Invalid PDF text selection/);
  await assert.rejects(replacePdfText({ ...request(bytes), rect: [0, 0, 10] }, workerPath), /Invalid PDF text selection/);
  const pending = replacePdfText(request(Buffer.from('hang')), workerPath, { timeoutMs: 40 });
  await assert.rejects(replacePdfText(request(bytes), workerPath), /Another PDF text edit is running/);
  await assert.rejects(pending, /timed out/);
  await assert.rejects(replacePdfText(request(Buffer.from('empty')), workerPath), /no document/);
  const output = await replacePdfText(request(bytes), workerPath);
  assert.equal(Buffer.from(output).toString(), '%PDF-out');
});
