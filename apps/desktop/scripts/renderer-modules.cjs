// Vite policy for the desktop renderer's native engines and optional modules.
//
// A renderer built without its engines still bundles: `new URL('./x.wasm',
// import.meta.url)` is left for runtime when the file is absent, so the packaged
// app would launch and show "Native worker failed". This plugin fails the build
// instead, before packaging can hide it. INJOFFICE_ALLOW_STUBS=1 opts into
// warnings and empty stubs for local UI work only; CI and packaging never set it.
const fs = require('node:fs');
const path = require('node:path');
const { ENGINE_BUILD, missingNativeEngineAssets, missingRendererAssets } = require('./native-assets.cjs');

const STUB_ENV = 'INJOFFICE_ALLOW_STUBS';
const SPREADSHEET_STUB = '\0injoffice-spreadsheet-editor';
const PPTX_RENDER_STUB = '\0injoffice-pptx-render';
const desktopRoot = path.resolve(__dirname, '..');

function nativeRendererModules({
  allowStubs = process.env[STUB_ENV] === '1',
  root = desktopRoot,
  repoRoot = path.resolve(root, '../..'),
  pptxRenderEntry = path.join(repoRoot, 'packages/pptx-render/src/index.ts'),
  exists = fs.existsSync,
  read = file => fs.readFileSync(file, 'utf8'),
  warn = message => console.warn(message),
} = {}) {
  const spreadsheetEditor = path.join(root, 'src/SpreadsheetEditor.tsx');
  const advice = `Set ${STUB_ENV}=1 to build a stubbed renderer for local UI work only; a packaged app must never ship without its engines.`;
  const refuse = (problem, stub) => {
    if (!allowStubs) throw new Error(`[injoffice-desktop] ${problem} ${advice}`);
    warn(`[injoffice-desktop] ${problem} Continuing because ${STUB_ENV}=1; this renderer must not be packaged.`);
    return stub;
  };
  return {
    name: 'injoffice-desktop-native-renderer-modules',
    enforce: 'pre',
    buildStart() {
      const missing = missingNativeEngineAssets({ root: repoRoot, exists });
      if (missing.length > 0) refuse(`Native engine assets are missing:\n  ${missing.join('\n  ')}\nBuild the engines first: ${ENGINE_BUILD} (needs Go).`);
    },
    resolveId(id) {
      if ((id === './SpreadsheetEditor' || id === './SpreadsheetEditor.tsx' || id.endsWith('/SpreadsheetEditor')) && !exists(spreadsheetEditor)) {
        return refuse(`The spreadsheet editor is missing (${spreadsheetEditor}); workbooks would open as an empty pane.`, SPREADSHEET_STUB);
      }
      if (id === '@injoffice/pptx-render' && (!exists(pptxRenderEntry) || !read(pptxRenderEntry).includes('exportNativePptxSlideSvg'))) {
        // Slide SVG export is a renderer feature, not an engine: @injoffice/pptx-render does
        // not export exportNativePptxSlideSvg yet, so this stub is expected and announced at
        // build time. The PPTX engine itself (@injoffice/pptx-wasm) is covered by buildStart.
        warn(`[injoffice-desktop] ${pptxRenderEntry} does not export exportNativePptxSlideSvg; slide SVG export will report "SVG export is unavailable in this app build."`);
        return PPTX_RENDER_STUB;
      }
      return undefined;
    },
    load(id) {
      if (id === SPREADSHEET_STUB) return 'export function SpreadsheetEditor() { return null }\n';
      if (id === PPTX_RENDER_STUB) return 'export async function exportNativePptxSlideSvg() { throw new Error("SVG export is unavailable in this app build."); }\n';
      return undefined;
    },
    generateBundle(_options, bundle) {
      const missing = missingRendererAssets(Object.keys(bundle), repoRoot);
      if (missing.length > 0) refuse(`The renderer bundle has no hashed copy of these engine assets: ${missing.join(', ')}. Vite left their URLs to be resolved at runtime, where they do not exist.`);
    },
  };
}

module.exports = { STUB_ENV, nativeRendererModules };
