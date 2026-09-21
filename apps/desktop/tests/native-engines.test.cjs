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

// Run formatting has to work in the tree electron-builder packs, not only in
// the Go tests: this drives the built DOCX worker end to end, bolding part of
// one run in a real package and re-extracting the result.
test('docx: the built worker patches run properties over a selection and re-extracts them', async () => {
  const files = engineFiles('docx');
  const worker = startWorker(files);
  const envelope = { protocol: PROTOCOL, version: 1, format: 'docx' };
  const init = await worker.send({ ...envelope, id: 'init', op: 'init', assets: { wasmUrl: files.wasm, goRuntimeUrl: files.goRuntime } });
  assert.equal(init.ok, true, `init: ${JSON.stringify(init.error)}`);

  const source = fs.readFileSync(path.join(__dirname, '../../../go/officecompat/corpus/generated/packages/docx-inline-png-page-paint.docx'));
  const buffer = source.buffer.slice(source.byteOffset, source.byteOffset + source.byteLength);
  const extract = await worker.send({ ...envelope, id: 'extract', op: 'extract', bytes: buffer });
  assert.equal(extract.ok, true, `extract: ${JSON.stringify(extract.error)}`);
  const contract = JSON.parse(extract.result.contractJson);
  const paragraph = contract.body.blocks.find(block => block.paragraph?.runs.some(run => run.kind === 'text'))?.paragraph;
  const run = paragraph.runs.find(candidate => candidate.kind === 'text');
  const sourceText = paragraph.runs.filter(candidate => candidate.kind === 'text').map(candidate => candidate.text).join('');
  assert.ok(paragraph.edit_policy.allowed_operations.includes('properties.patch'), `paragraph policy: ${JSON.stringify(paragraph.edit_policy)}`);
  assert.ok(run.text.length > 4, `fixture run text: ${JSON.stringify(run.text)}`);

  const payload = JSON.stringify({ mutations: [{
    target_kind: 'run',
    target_id: run.id,
    expected_xml_sha256: run.anchor.xml_sha256,
    properties: { bold: true, color: 'FF0000' },
    range: { start_utf16: 0, end_utf16: 4 },
  }] });
  const applied = await worker.send({ ...envelope, id: 'apply', op: 'apply', original: buffer, payload, expectedRevision: contract.source.package_sha256 });
  assert.equal(applied.ok, true, `apply: ${JSON.stringify(applied.error)}`);

  const after = await worker.send({ ...envelope, id: 'reextract', op: 'extract', bytes: applied.result.bytes });
  assert.equal(after.ok, true, `re-extract: ${JSON.stringify(after.error)}`);
  const reread = JSON.parse(after.result.contractJson).body.blocks.find(block => block.paragraph?.anchor.path === paragraph.anchor.path).paragraph;
  const formatted = reread.runs.filter(candidate => candidate.kind === 'text');
  assert.equal(formatted.map(candidate => candidate.text).join(''), sourceText, 'the paragraph text is unchanged');
  assert.equal(formatted[0].text, run.text.slice(0, 4), 'the run split at the selection boundary');
  assert.equal(formatted[0].properties.bold, true);
  assert.equal(formatted[0].properties.color, 'FF0000');
  assert.equal(formatted[1].properties?.bold, undefined, 'the unselected remainder is untouched');
});

test('docx: packaged worker inserts an editable paragraph and splits text at Enter', async () => {
  const files = engineFiles('docx'), worker = startWorker(files);
  const envelope = { protocol: PROTOCOL, version: 1, format: 'docx' };
  const init = await worker.send({ ...envelope, id: 'init', op: 'init', assets: { wasmUrl: files.wasm, goRuntimeUrl: files.goRuntime } });
  assert.equal(init.ok, true);
  let bytes = await createBlankDocument('docx');
  const extract = async () => {
    const reply = await worker.send({ ...envelope, id: 'extract', op: 'extract', bytes });
    assert.equal(reply.ok, true, JSON.stringify(reply.error));
    return JSON.parse(reply.result.contractJson);
  };
  const apply = async mutation => {
    const document = await extract();
    const reply = await worker.send({ ...envelope, id: 'apply', op: 'apply', original: bytes, expectedRevision: document.source.package_sha256, payload: JSON.stringify({ mutations: [mutation] }) });
    assert.equal(reply.ok, true, JSON.stringify(reply.error)); bytes = reply.result.bytes;
  };
  let document = await extract(), p = document.body.blocks[0].paragraph;
  assert.ok(p.edit_policy.allowed_operations.includes('block.insert_after'));
  await apply({ target_kind: 'paragraph', target_id: p.id, expected_xml_sha256: p.anchor.xml_sha256, operation: 'block.insert_after', text: '' });
  document = await extract(); assert.equal(document.body.blocks.length, 2);
  let run = document.body.blocks[1].paragraph.runs[0];
  await apply({ target_kind: 'run', target_id: run.id, expected_xml_sha256: run.anchor.xml_sha256, text: 'hello world' });
  document = await extract(); p = document.body.blocks[1].paragraph; run = p.runs[0];
  await apply({ target_kind: 'paragraph', target_id: p.id, expected_xml_sha256: p.anchor.xml_sha256, operation: 'paragraph.split', split: { run_id: run.id, offset_utf16: 5 } });
  document = await extract();
  assert.deepEqual(document.body.blocks.map(block => block.paragraph.runs.map(run => run.text ?? '').join('')), ['', 'hello', ' world']);
});
