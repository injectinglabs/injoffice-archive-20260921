// The native engines the desktop renderer must ship: the Go WASM modules,
// their wasm_exec.js runtime and the classic workers, as declared by each
// engine package's `injoffice.assets`. Vite copies them from each engine dist directory
// into renderer/assets with a content hash; electron-builder packs renderer/.
const fs = require('node:fs');
const path = require('node:path');

const repoRoot = path.resolve(__dirname, '../../..');
const ENGINE_PACKAGES = ['docx-wasm', 'xlsx-wasm', 'pptx-wasm'];
const ENGINE_BUILD = 'npm run build -w @injoffice/docx-wasm -w @injoffice/xlsx-wasm -w @injoffice/pptx-wasm';

function nativeEngineAssets(root = repoRoot) {
  return ENGINE_PACKAGES.map(name => {
    const directory = path.join(root, 'packages', name);
    const manifest = JSON.parse(fs.readFileSync(path.join(directory, 'package.json'), 'utf8'));
    const assets = manifest.injoffice?.assets ?? [];
    if (assets.length === 0) throw new Error(`${manifest.name} declares no injoffice.assets`);
    return { package: manifest.name, assets: assets.map(file => ({ file, path: path.join(directory, 'dist', file) })) };
  });
}

/** Engine asset files absent from the engine dist directories (the engines were not built). */
function missingNativeEngineAssets({ root = repoRoot, exists = fs.existsSync } = {}) {
  return nativeEngineAssets(root).flatMap(engine => engine.assets.filter(asset => !exists(asset.path)).map(asset => asset.path));
}

const escape = text => text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

/** One pattern per distinct asset file name; wasm_exec.js is shared by all engines and emitted once. */
function rendererAssetPatterns(root = repoRoot) {
  const patterns = new Map();
  for (const engine of nativeEngineAssets(root)) {
    for (const { file } of engine.assets) {
      if (patterns.has(file)) continue;
      const extension = path.extname(file);
      const stem = file.slice(0, -extension.length);
      patterns.set(file, new RegExp(`^(?:.*/)?${escape(stem)}-[A-Za-z0-9_-]+${escape(extension)}$`));
    }
  }
  return patterns;
}

/** Declared engine assets with no hashed copy among the emitted renderer file names. */
function missingRendererAssets(fileNames, root = repoRoot) {
  const names = Array.from(fileNames, name => String(name).replace(/\\/g, '/'));
  return [...rendererAssetPatterns(root)].filter(([, pattern]) => !names.some(name => pattern.test(name))).map(([file]) => file);
}

module.exports = { ENGINE_BUILD, ENGINE_PACKAGES, nativeEngineAssets, missingNativeEngineAssets, rendererAssetPatterns, missingRendererAssets };
