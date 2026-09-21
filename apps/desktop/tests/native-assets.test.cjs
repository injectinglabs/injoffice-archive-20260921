const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { ENGINE_PACKAGES, nativeEngineAssets, missingNativeEngineAssets, missingRendererAssets } = require('../scripts/native-assets.cjs');
const { STUB_ENV, nativeRendererModules } = require('../scripts/renderer-modules.cjs');
const { PDF_FONT, pdfResourceFiles, pdfAssets } = require('../scripts/pdf-assets.cjs');

const root = path.resolve(__dirname, '..');
const wasm = /\.wasm$/;

test('desktop Vite emits the PDF font and support resources into the packaged renderer', () => {
  const files = pdfResourceFiles();
  assert.equal(files.get(PDF_FONT).readUInt32BE(0), 0x00010000, 'valid TrueType font');
  assert.ok(files.has('pdf-assets/standard_fonts/LICENSE_LIBERATION'));
  const emitted = new Map();
  pdfAssets().generateBundle.call({ emitFile: asset => emitted.set(asset.fileName, asset.source) });
  assert.deepEqual(emitted, files);
  const config = fs.readFileSync(path.join(root, 'vite.config.ts'), 'utf8');
  assert.match(config, /plugins:.*pdfAssets\(\)/);
  const pkg = JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8'));
  assert.ok(pkg.build.files.includes('renderer/**/*'), 'electron-builder includes renderer/pdf-assets');
});

test('every engine package declares its wasm, wasm_exec.js and worker assets', () => {
  const engines = nativeEngineAssets();
  assert.deepEqual(engines.map(engine => engine.package), ENGINE_PACKAGES.map(name => `@injoffice/${name}`));
  for (const engine of engines) {
    const files = engine.assets.map(asset => asset.file);
    assert.ok(files.some(file => wasm.test(file)), `${engine.package} wasm`);
    assert.ok(files.includes('wasm_exec.js'), `${engine.package} wasm_exec.js`);
    assert.ok(files.some(file => file.endsWith('.worker.js')), `${engine.package} worker`);
    for (const asset of engine.assets) assert.match(asset.path, /packages[\\/](docx|xlsx|pptx)-wasm[\\/]dist[\\/]/);
  }
});

test('missing dist assets and missing hashed renderer copies are both reported by file', () => {
  const [docx] = nativeEngineAssets();
  const missingWasm = docx.assets.find(asset => wasm.test(asset.file)).path;
  assert.deepEqual(missingNativeEngineAssets({ exists: file => file !== missingWasm }), [missingWasm]);
  assert.deepEqual(missingNativeEngineAssets({ exists: () => true }), []);
  const emitted = ['renderer/index.html', 'assets/docxnative-E-jJdecv.wasm', 'assets/docxnative.worker-xMgFAj-H.js', 'assets/wasm_exec-vVX9mWnE.js',
    'assets/xlsxnative-C2eZhFuA.wasm', 'assets/xlsxrichsource-BEesRfty.wasm', 'assets/xlsxnative.worker-jzN5YjB9.js', 'assets/xlsxsource.worker-ByGA6KWF.js',
    'assets/xlsxsource2.worker-S8j5gXdi.js', 'assets/xlsxrichsource.worker-C-dHYZaU.js', 'assets/pptxnative-BIGkBVsI.wasm', 'assets/pptxnative.worker-Ry9sAili.js'];
  assert.deepEqual(missingRendererAssets(emitted), []);
  assert.deepEqual(missingRendererAssets(emitted.filter(name => !name.includes('pptxnative-'))), ['pptxnative.wasm']);
  assert.deepEqual(missingRendererAssets(['assets/docxnative.wasm']).includes('docxnative.wasm'), true, 'an unhashed name is not a Vite asset');
});

function plugin(overrides = {}) {
  const warnings = [];
  const instance = nativeRendererModules({ allowStubs: false, exists: () => true, read: () => 'export async function exportNativePptxSlideSvg() {}', warn: message => warnings.push(message), ...overrides });
  return { instance, warnings };
}

test('a build without the engines fails before packaging can hide it', () => {
  const [docx] = nativeEngineAssets();
  const missingWasm = docx.assets.find(asset => wasm.test(asset.file)).path;
  const { instance, warnings } = plugin({ exists: file => file !== missingWasm });
  assert.throws(() => instance.buildStart(), error => error.message.includes(missingWasm) && error.message.includes(STUB_ENV) && error.message.includes('npm run build -w @injoffice/docx-wasm'));
  assert.throws(() => instance.generateBundle({}, { 'assets/index-abc.js': {} }), error => /no hashed copy of these engine assets/.test(error.message) && error.message.includes("docxnative.wasm"));
  assert.deepEqual(warnings, []);
  const complete = plugin();
  complete.instance.buildStart();
  assert.deepEqual(complete.warnings, []);
});

test(`${STUB_ENV}=1 downgrades missing engines to warnings for local UI work`, () => {
  const { instance, warnings } = plugin({ allowStubs: true, exists: () => false });
  instance.buildStart();
  instance.generateBundle({}, {});
  assert.equal(warnings.length, 2);
  for (const warning of warnings) assert.match(warning, /must not be packaged/);
});

test('the spreadsheet editor is required; its stub exists only for local UI work', () => {
  const editor = path.join(root, 'src/SpreadsheetEditor.tsx');
  assert.equal(fs.existsSync(editor), true);
  assert.equal(plugin().instance.resolveId('./SpreadsheetEditor'), undefined);
  const missing = plugin({ exists: file => file !== editor });
  assert.throws(() => missing.instance.resolveId('./SpreadsheetEditor'), /spreadsheet editor is missing/);
  const stubbed = plugin({ allowStubs: true, exists: file => file !== editor });
  const id = stubbed.instance.resolveId('./SpreadsheetEditor');
  assert.match(stubbed.instance.load(id), /return null/);
});

test('slide SVG export is stubbed with a build-time warning until pptx-render exports it', () => {
  const { instance, warnings } = plugin({ read: () => 'export const nothing = 1' });
  const id = instance.resolveId('@injoffice/pptx-render');
  assert.match(instance.load(id), /SVG export is unavailable in this app build/);
  assert.equal(warnings.length, 1);
  assert.match(warnings[0], /exportNativePptxSlideSvg/);
  assert.equal(plugin().instance.resolveId('@injoffice/pptx-render'), undefined);
});

test('vite resolves the engine packages from their built dist and the desktop build compiles them first', () => {
  const config = fs.readFileSync(path.join(root, 'vite.config.ts'), 'utf8');
  assert.match(config, /nativeRendererModules\(\)/);
  assert.doesNotMatch(config, /@injoffice\/(?:docx|xlsx|pptx)-wasm/, 'engine packages must resolve to dist/, where their assets live');
  const pkg = JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8'));
  for (const name of ENGINE_PACKAGES) assert.match(pkg.scripts.prebuild, new RegExp(`-w @injoffice/${name}`));
  assert.equal(pkg.scripts.build, 'vite build');
});
