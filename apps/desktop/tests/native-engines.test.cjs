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

test('docx: packaged worker inserts, updates and removes a selected hyperlink', async () => {
  const files = engineFiles('docx'), worker = startWorker(files);
  const envelope = { protocol: PROTOCOL, version: 1, format: 'docx' };
  assert.equal((await worker.send({ ...envelope, id: 'init', op: 'init', assets: { wasmUrl: files.wasm, goRuntimeUrl: files.goRuntime } })).ok, true);
  let bytes = await createBlankDocument('docx');
  const extract = async () => {
    const reply = await worker.send({ ...envelope, id: 'extract', op: 'extract', bytes });
    assert.equal(reply.ok, true, JSON.stringify(reply.error)); return JSON.parse(reply.result.contractJson);
  };
  const apply = async (document, mutation) => {
    const reply = await worker.send({ ...envelope, id: 'apply', op: 'apply', original: bytes, expectedRevision: document.source.package_sha256, payload: JSON.stringify({mutations: [mutation]}) });
    assert.equal(reply.ok, true, JSON.stringify(reply.error)); bytes = reply.result.bytes;
  };
  let document = await extract(), run = document.body.blocks[0].paragraph.runs[0];
  await apply(document, {target_kind: 'run', target_id: run.id, expected_xml_sha256: run.anchor.xml_sha256, text: 'Visit example today'});
  document = await extract(); run = document.body.blocks[0].paragraph.runs[0];
  assert.equal(run.can_edit_hyperlink, true);
  await apply(document, {target_kind: 'run', target_id: run.id, expected_xml_sha256: run.anchor.xml_sha256, operation: 'hyperlink.set', hyperlink: {url: 'https://example.com'}, range: {start_utf16: 6, end_utf16: 13}});
  document = await extract();
  assert.deepEqual(document.body.blocks[0].paragraph.runs.map(run => [run.text, run.hyperlink?.url]), [['Visit ', undefined], ['example', 'https://example.com'], [' today', undefined]]);
  for (const url of ['mailto:editor@example.com', null]) {
    run = document.body.blocks[0].paragraph.runs[1];
    await apply(document, {target_kind: 'run', target_id: run.id, expected_xml_sha256: run.anchor.xml_sha256, operation: 'hyperlink.set', hyperlink: {url, expected_xml_sha256: run.hyperlink.anchor.xml_sha256}});
    document = await extract();
    assert.equal(document.body.blocks[0].paragraph.runs[1].hyperlink?.url, url ?? undefined);
    assert.equal(document.body.blocks[0].paragraph.runs.map(run => run.text).join(''), 'Visit example today');
  }
});

test('docx: packaged worker replaces inline-image text with edge spaces and empty text', async () => {
  const files = engineFiles('docx'), worker = startWorker(files);
  const envelope = { protocol: PROTOCOL, version: 1, format: 'docx' };
  assert.equal((await worker.send({ ...envelope, id: 'init', op: 'init', assets: { wasmUrl: files.wasm, goRuntimeUrl: files.goRuntime } })).ok, true);
  const source = fs.readFileSync(path.join(__dirname, '../../../go/officecompat/corpus/generated/packages/docx-inline-png-page-paint.docx'));
  for (const text of [' After edited', 'After edited ', ' After edited ', '']) {
    const bytes = source.buffer.slice(source.byteOffset, source.byteOffset + source.byteLength);
    const extracted = await worker.send({ ...envelope, id: 'extract', op: 'extract', bytes });
    assert.equal(extracted.ok, true);
    const doc = JSON.parse(extracted.result.contractJson);
    const run = doc.body.blocks.flatMap(block => block.paragraph?.runs ?? []).find(run => run.text === ' After');
    assert.ok(run, 'fixture contains the failing leading-space run');
    const applied = await worker.send({ ...envelope, id: 'apply', op: 'apply', original: bytes, expectedRevision: doc.source.package_sha256, payload: JSON.stringify({ mutations: [{ target_kind: 'run', target_id: run.id, expected_xml_sha256: run.anchor.xml_sha256, text }] }) });
    assert.equal(applied.ok, true, JSON.stringify(applied.error));
    const reopened = await worker.send({ ...envelope, id: 'reopen', op: 'extract', bytes: applied.result.bytes });
    assert.equal(reopened.ok, true);
    const next = JSON.parse(reopened.result.contractJson).body.blocks.flatMap(block => block.paragraph?.runs ?? []).find(candidate => candidate.anchor.path === run.anchor.path);
    assert.equal(next.text, text);
  }
});

test('docx: packaged worker round-trips paragraph spacing and indentation on empty text', async () => {
  const files = engineFiles('docx'), worker = startWorker(files);
  const envelope = { protocol: PROTOCOL, version: 1, format: 'docx' };
  assert.equal((await worker.send({ ...envelope, id: 'init', op: 'init', assets: { wasmUrl: files.wasm, goRuntimeUrl: files.goRuntime } })).ok, true);
  const source = await createBlankDocument('docx');
  let bytes = source.buffer.slice(source.byteOffset, source.byteOffset + source.byteLength);
  for (const properties of [
    { spacing_before_twips: 120, spacing_after_twips: 240, indent_left_twips: 720, indent_right_twips: 360, first_line_twips: 240, line_spacing: 360, line_rule: 'auto' },
    { spacing_before_twips: null, first_line_twips: null, hanging_twips: 360, line_spacing: 480, line_rule: 'exact' },
  ]) {
    const extracted = await worker.send({ ...envelope, id: 'extract', op: 'extract', bytes });
    assert.equal(extracted.ok, true, JSON.stringify(extracted.error));
    const document = JSON.parse(extracted.result.contractJson), paragraph = document.body.blocks[0].paragraph;
    const payload = JSON.stringify({ mutations: [{ target_kind: 'paragraph', target_id: paragraph.id, expected_xml_sha256: paragraph.anchor.xml_sha256, properties }] });
    const applied = await worker.send({ ...envelope, id: 'apply', op: 'apply', original: bytes, expectedRevision: document.source.package_sha256, payload });
    assert.equal(applied.ok, true, JSON.stringify(applied.error));
    bytes = applied.result.bytes;
    const reread = await worker.send({ ...envelope, id: 'reread', op: 'extract', bytes });
    assert.equal(reread.ok, true, JSON.stringify(reread.error));
    const result = JSON.parse(reread.result.contractJson).body.blocks[0].paragraph;
    for (const [key, value] of Object.entries(properties)) assert.equal(result.properties[key], value ?? undefined, key);
    assert.equal(result.runs.map(run => run.text ?? '').join(''), '');
    assert.ok(result.edit_policy.allowed_operations.includes('properties.patch'));
  }
});

test('docx: packaged worker applies margins, orientation and paper size to a section', async () => {
  const files = engineFiles('docx'), worker = startWorker(files);
  const envelope = { protocol: PROTOCOL, version: 1, format: 'docx' };
  assert.equal((await worker.send({ ...envelope, id: 'init', op: 'init', assets: { wasmUrl: files.wasm, goRuntimeUrl: files.goRuntime } })).ok, true);
  const source = await createBlankDocument('docx');
  let bytes = source.buffer.slice(source.byteOffset, source.byteOffset + source.byteLength);
  let page = { width_twips: 12240, height_twips: 15840, orientation: 'portrait', margin_top_twips: 720, margin_right_twips: 720, margin_bottom_twips: 720, margin_left_twips: 720 };
  for (const change of [{}, { width_twips: 15840, height_twips: 12240, orientation: 'landscape' }, { width_twips: 16838, height_twips: 11906 }]) {
    page = { ...page, ...change };
    const extract = await worker.send({ ...envelope, id: 'extract', op: 'extract', bytes });
    assert.equal(extract.ok, true, JSON.stringify(extract.error));
    const document = JSON.parse(extract.result.contractJson), section = document.sections[0];
    assert.ok(section.edit_policy.allowed_operations.includes('section.page.patch'));
    const payload = JSON.stringify({ mutations: [{ target_kind: 'section', target_id: section.id, expected_xml_sha256: section.anchor.xml_sha256, operation: 'section.page.patch', page }] });
    const applied = await worker.send({ ...envelope, id: 'apply', op: 'apply', original: bytes, expectedRevision: document.source.package_sha256, payload });
    assert.equal(applied.ok, true, JSON.stringify(applied.error));
    bytes = applied.result.bytes;
    const reread = await worker.send({ ...envelope, id: 'reread', op: 'extract', bytes });
    assert.equal(reread.ok, true, JSON.stringify(reread.error));
    const geometry = JSON.parse(reread.result.contractJson).sections[0].page;
    assert.equal(geometry.width_twips, page.width_twips);
    assert.equal(geometry.height_twips, page.height_twips);
    assert.equal(geometry.orientation, page.orientation);
    for (const side of ['top', 'right', 'bottom', 'left']) assert.equal(geometry.margins[`${side}_twips`], page[`margin_${side}_twips`]);
    assert.equal(geometry.margins.header_twips, section.page.margins.header_twips);
    assert.equal(geometry.margins.footer_twips, section.page.margins.footer_twips);
    assert.equal(geometry.margins.gutter_twips, section.page.margins.gutter_twips);
  }
});

// List numbering has to work in the packaged worker: turn a blank paragraph
// into a bullet, prove the numbering catalog and w:numPr read back, then
// remove the list.
test('docx: the built worker attaches and removes a bullet list', async () => {
  const files = engineFiles('docx');
  const worker = startWorker(files);
  const envelope = { protocol: PROTOCOL, version: 1, format: 'docx' };
  const init = await worker.send({ ...envelope, id: 'init', op: 'init', assets: { wasmUrl: files.wasm, goRuntimeUrl: files.goRuntime } });
  assert.equal(init.ok, true, `init: ${JSON.stringify(init.error)}`);

  const source = await createBlankDocument('docx');
  const buffer = source.buffer.slice(source.byteOffset, source.byteOffset + source.byteLength);
  const extract = await worker.send({ ...envelope, id: 'extract', op: 'extract', bytes: buffer });
  assert.equal(extract.ok, true, `extract: ${JSON.stringify(extract.error)}`);
  const contract = JSON.parse(extract.result.contractJson);
  const paragraph = contract.body.blocks[0].paragraph;
  assert.ok(paragraph.edit_policy.allowed_operations.includes('properties.patch'), `paragraph policy: ${JSON.stringify(paragraph.edit_policy)}`);
  assert.ok(Array.isArray(contract.numbering_definitions) && contract.numbering_definitions.some(item => item.levels[0].format === 'bullet'),
    `blank document numbering catalog: ${JSON.stringify(contract.numbering_definitions)}`);

  const payload = JSON.stringify({ mutations: [{
    target_kind: 'paragraph',
    target_id: paragraph.id,
    expected_xml_sha256: paragraph.anchor.xml_sha256,
    properties: { numbering_kind: 'bullet', numbering_level: 0 },
  }] });
  const applied = await worker.send({ ...envelope, id: 'apply', op: 'apply', original: buffer, payload, expectedRevision: contract.source.package_sha256 });
  assert.equal(applied.ok, true, `apply: ${JSON.stringify(applied.error)}`);

  const after = await worker.send({ ...envelope, id: 'reextract', op: 'extract', bytes: applied.result.bytes });
  assert.equal(after.ok, true, `re-extract: ${JSON.stringify(after.error)}`);
  const reread = JSON.parse(after.result.contractJson);
  const listed = reread.body.blocks[0].paragraph;
  assert.equal(listed.properties.numbering.num_id, '1');
  assert.equal(listed.properties.numbering.level, 0);

  const remove = JSON.stringify({ mutations: [{
    target_kind: 'paragraph',
    target_id: listed.id,
    expected_xml_sha256: listed.anchor.xml_sha256,
    properties: { numbering_num_id: null, numbering_level: null },
  }] });
  const cleared = await worker.send({ ...envelope, id: 'apply', op: 'apply', original: applied.result.bytes, payload: remove, expectedRevision: reread.source.package_sha256 });
  assert.equal(cleared.ok, true, `remove: ${JSON.stringify(cleared.error)}`);
  const gone = JSON.parse((await worker.send({ ...envelope, id: 'final', op: 'extract', bytes: cleared.result.bytes })).result.contractJson).body.blocks[0].paragraph;
  assert.equal(gone.properties.numbering, undefined);
});
