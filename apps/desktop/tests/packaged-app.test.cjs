// Structural check of every application electron-builder unpacked under
// apps/desktop/release: the Electron host, the renderer entry and every native
// engine asset must be inside app.asar, byte-identical to the built engines.
// Runs on every OS runner without launching a GUI. INJOFFICE_RELEASE_DIR overrides the directory.
const test = require('node:test');
const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const fs = require('node:fs');
const path = require('node:path');
const { defaultReleaseDir, findPackagedApps, readAsar } = require('../scripts/packaged-app.cjs');
const { nativeEngineAssets, missingRendererAssets, rendererAssetPatterns } = require('../scripts/native-assets.cjs');

const releaseDir = process.env.INJOFFICE_RELEASE_DIR ? path.resolve(process.env.INJOFFICE_RELEASE_DIR) : defaultReleaseDir;
const sha256 = bytes => crypto.createHash('sha256').update(bytes).digest('hex');
const apps = findPackagedApps(releaseDir);

test(`electron-builder unpacked at least one application under ${releaseDir}`, () => {
  const listing = fs.existsSync(releaseDir) ? fs.readdirSync(releaseDir).join(', ') : '(missing)';
  assert.ok(apps.length > 0, `no mac*/ *.app, win*-unpacked/ or linux*-unpacked/ found; release/ has: ${listing}`);
  for (const app of apps) assert.ok(fs.existsSync(app.asar), `${app.name}: ${app.asar} is missing`);
});

for (const app of apps) {
  test(`${app.name}: app.asar carries the host, the renderer and every native engine asset`, () => {
    const asar = readAsar(app.asar);
    try {
      const entries = asar.list();
      const metadata = JSON.parse(asar.read('package.json').toString('utf8'));
      assert.equal(metadata.main, 'electron/main.cjs');
      assert.equal(metadata.name, '@injoffice/desktop');
      for (const host of ['electron/main.cjs', 'electron/preload.cjs', 'electron/new-document.cjs', 'electron/updates.cjs']) assert.ok(asar.has(host), host);

      const html = asar.read('renderer/index.html').toString('utf8');
      const [, entry] = /src="\.\/(assets\/index-[\w-]+\.js)"/.exec(html) ?? [];
      assert.ok(entry, `renderer/index.html references a hashed ./assets/index-*.js: ${html}`);
      assert.ok(asar.has(`renderer/${entry}`), `renderer/${entry}`);

      const rendererAssets = entries.filter(name => name.startsWith('renderer/assets/'));
      assert.deepEqual(missingRendererAssets(rendererAssets), [], `engine assets missing from ${app.asar}`);

      // The packaged engine bytes are the built engine bytes.
      const patterns = rendererAssetPatterns();
      for (const engine of nativeEngineAssets()) {
        for (const asset of engine.assets) {
          const packaged = rendererAssets.find(name => patterns.get(asset.file).test(name));
          assert.ok(packaged, `${engine.package} ${asset.file}`);
          assert.equal(asar.size(packaged), fs.statSync(asset.path).size, `${packaged} size matches ${asset.path}`);
          assert.equal(sha256(asar.read(packaged)), sha256(fs.readFileSync(asset.path)), `${packaged} bytes match ${asset.path}`);
        }
      }
      const wasm = rendererAssets.filter(name => name.endsWith('.wasm'));
      console.log(`${app.name}: ${entries.length} asar entries, ${rendererAssets.length} renderer assets, ${wasm.length} wasm modules (${wasm.map(name => `${path.basename(name)} ${asar.size(name)} B`).join(', ')})`);
    } finally { asar.close(); }
  });
}
