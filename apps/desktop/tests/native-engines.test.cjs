// Proves the native engines that will be packaged are present and functional:
// each built worker (packages/*-wasm/dist) is driven over its own message
// protocol with the built wasm_exec.js and .wasm, and extracts a blank document.
// No browser, no bundler: this is the tree electron-builder packs.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { ENGINE_BUILD, nativeEngineAssets, missingNativeEngineAssets } = require('../scripts/native-assets.cjs');
const { createBlankDocument } = require('../electron/new-document.cjs');

const PROTOCOL = 'injoffice.native-wasm-worker';
const FORMATS = ['docx', 'xlsx', 'pptx'];

function engineFiles(format) {
  const engine = nativeEngineAssets().find(candidate => candidate.package === `@injoffice/${format}-wasm`);
  const file = name => engine.assets.find(asset => asset.file === name)?.path;
  return { worker: file(`${format}native.worker.js`), wasm: file(`${format}native.wasm`), goRuntime: file('wasm_exec.js') };
}

// A classic-worker host: importScripts and fetch read the built files, postMessage collects replies.
function startWorker(files) {
  const replies = [];
  const sandbox = {
    ArrayBuffer, Uint8Array, TextEncoder, TextDecoder, WebAssembly, crypto: globalThis.crypto, performance,
    setTimeout, clearTimeout, setInterval, clearInterval, queueMicrotask, console,
    postMessage: message => replies.push(message),
    close: () => { sandbox.closed = true; },
    importScripts: url => { assert.equal(url, files.goRuntime); vm.runInContext(fs.readFileSync(url, 'utf8'), context, { filename: url }); },
    fetch: async url => {
      assert.equal(url, files.wasm);
      const bytes = fs.readFileSync(url);
      return { ok: true, status: 200, arrayBuffer: async () => bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) };
    },
  };
  sandbox.self = sandbox;
  const context = vm.createContext(sandbox);
  vm.runInContext(fs.readFileSync(files.worker, 'utf8'), context, { filename: files.worker });
  assert.equal(typeof sandbox.onmessage, 'function', `${files.worker} installs onmessage`);
  return {
    async send(request) {
      await sandbox.onmessage({ data: request });
      const reply = replies.at(-1);
      assert.equal(reply?.id, request.id, `reply to ${request.op}`);
      return reply;
    },
  };
}

test(`the engine assets that electron-builder packs exist (${ENGINE_BUILD})`, () => {
  assert.deepEqual(missingNativeEngineAssets(), []);
});

for (const format of FORMATS) {
  test(`${format}: the built worker boots the built wasm_exec.js + .wasm and extracts a blank ${format}`, async () => {
    const files = engineFiles(format);
    for (const [name, file] of Object.entries(files)) assert.ok(file && fs.existsSync(file), `${format} ${name}`);
    const worker = startWorker(files);
    const envelope = { protocol: PROTOCOL, version: 1, format };
    const init = await worker.send({ ...envelope, id: 'init', op: 'init', assets: { wasmUrl: files.wasm, goRuntimeUrl: files.goRuntime } });
    assert.equal(init.ok, true, `init: ${JSON.stringify(init.error)}`);
    const bytes = await createBlankDocument(format);
    const buffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
    const extract = await worker.send({ ...envelope, id: 'extract', op: 'extract', bytes: buffer });
    assert.equal(extract.ok, true, `extract: ${JSON.stringify(extract.error)}`);
    const contract = JSON.parse(extract.result.contractJson);
    const expected = { docx: ['protocol', 'revision', 'sections'], xlsx: ['protocol', 'revision', 'sheets'], pptx: ['contractVersion', 'sourceRevision', 'slides'] }[format];
    for (const key of expected) assert.ok(key in contract, `${format} contract has ${key}: ${Object.keys(contract).join(', ')}`);
    const units = contract.sections ?? contract.sheets ?? contract.slides;
    assert.ok(Array.isArray(units) && units.length === 1, `a blank ${format} extracts one section/sheet/slide`);
  });
}
