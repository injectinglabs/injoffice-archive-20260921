import { spawnSync } from 'node:child_process'
import { existsSync, mkdtempSync, readdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { join, resolve } from 'node:path'
import { tmpdir } from 'node:os'

const root = resolve(import.meta.dirname, '..')
const packageRoot = resolve(root, 'packages')
const temporary = mkdtempSync(join(tmpdir(), 'injoffice-consumer-'))
const cache = process.env.INJOFFICE_SMOKE_CACHE || resolve(temporary, 'npm-cache')
const npmEnvironment = { ...process.env, npm_config_cache: cache }
const requestedPackages = process.env.INJOFFICE_SMOKE_PACKAGES
  ? new Set(process.env.INJOFFICE_SMOKE_PACKAGES.split(',').map((name) => name.trim()).filter(Boolean))
  : undefined
const packageNames = []
const packageEntrypoints = []
const tarballs = []
const docsAdapterDocument = JSON.parse(readFileSync(resolve(root, 'go/officecompat/corpus/generated/expected/docx-strict-relocated.json'), 'utf8')).native
const docsAdapterParagraph = docsAdapterDocument.body.blocks[0].paragraph
const docsAdapterRun = docsAdapterParagraph.runs[0]
const docsAdapterBefore = {
  doc_size: docsAdapterRun.text.length + 2,
  paragraphs: [{
    paragraph_id: docsAdapterParagraph.id,
    node_from: 0,
    node_to: docsAdapterRun.text.length + 2,
    runs: [{
      run_id: docsAdapterRun.id,
      text_from: 1,
      text_to: docsAdapterRun.text.length + 1,
      text: docsAdapterRun.text,
      marks: ['font-size:24', 'italic'],
      native_properties: docsAdapterRun.properties,
    }],
  }],
}
const docsAdapterAfter = structuredClone(docsAdapterBefore)
docsAdapterAfter.paragraphs[0].runs[0].text = 'Strict native'
const sheetsAdapterWorkbook = JSON.parse(readFileSync(resolve(root, 'go/xlsxpatch/testdata/native-xlsx-v1/valid/lexical-render.json'), 'utf8'))
sheetsAdapterWorkbook.normal_style = {
  style_xf_id: 0, font_id: 0, font_name: 'Fixture', font_size_points: 11, font_bold: false, font_italic: false,
  font_record_sha256: `sha256:${'c'.repeat(64)}`,
}
const docsAdapterTransaction = {
  protocol: 'injoffice.docx.prosemirror-transaction',
  version: 1,
  schema_id: 'injoffice.docx.prosemirror-text',
  schema_version: 1,
  mutation_id: 'consumer-smoke',
  document_id: docsAdapterDocument.document_id,
  expected_revision: docsAdapterDocument.source.package_sha256,
  before: docsAdapterBefore,
  steps: [{
    step_type: 'replace', step_index: 0, from: 8, to: 14,
    paragraph_id: docsAdapterParagraph.id, run_id: docsAdapterRun.id,
    expected_xml_sha256: docsAdapterRun.anchor.xml_sha256,
    deleted_text: 'routed', before_text: 'Strict routed', after_text: 'Strict native',
    slice: { open_start: 0, open_end: 0, text: 'native', marks: ['font-size:24', 'italic'] },
  }],
  after: docsAdapterAfter,
}

const run = (command, args, cwd, env = process.env) => {
  const result = spawnSync(command, args, { cwd, env, encoding: 'utf8' })
  if (result.status !== 0) {
    throw new Error([result.stdout, result.stderr].filter(Boolean).join('\n'))
  }
  return result.stdout
}

const listFiles = (directory) => readdirSync(directory, { withFileTypes: true }).flatMap((entry) => {
  const path = resolve(directory, entry.name)
  return entry.isDirectory() ? listFiles(path) : [path]
})

try {
  for (const entry of readdirSync(packageRoot, { withFileTypes: true }).sort((a, b) => a.name < b.name ? -1 : a.name > b.name ? 1 : 0)) {
    if (!entry.isDirectory()) continue
    const directory = resolve(packageRoot, entry.name)
    const manifestPath = resolve(directory, 'package.json')
    if (!existsSync(manifestPath)) continue
    const manifest = JSON.parse(readFileSync(manifestPath, 'utf8'))
    if (manifest.private === true) continue
    if (requestedPackages && !requestedPackages.has(manifest.name)) continue
    packageNames.push(manifest.name)
    for (const [subpath, target] of Object.entries(manifest.exports ?? { '.': manifest.main })) {
      const importTarget = typeof target === 'string' ? target : target?.import
      if (typeof importTarget !== 'string' || !importTarget.endsWith('.js')) continue
      packageEntrypoints.push(subpath === '.' ? manifest.name : `${manifest.name}/${subpath.slice(2)}`)
    }
    const report = JSON.parse(run(process.platform === 'win32' ? 'npm.cmd' : 'npm', ['pack', '--json', '--pack-destination', temporary, directory], root, npmEnvironment))[0]
    tarballs.push(resolve(temporary, report.filename))
  }

  writeFileSync(resolve(temporary, 'package.json'), JSON.stringify({
    name: 'injoffice-consumer-smoke',
    private: true,
    type: 'module',
  }, null, 2))
  writeFileSync(resolve(temporary, 'smoke.mjs'), [
    ...packageNames.map((name) => `await import(${JSON.stringify(name)})`),
    ...(packageNames.includes('@injoffice/docs') ? [
      `const docs = await import('@injoffice/docs')`,
      `const docxCompiler = await import('@injoffice/docs/native-page-paint-compiler')`,
      `if (typeof docs.layoutNativeDocxHeadersFootersV1 !== 'function' || typeof docs.nativeDocxHeaderFooterLayoutSha256V1 !== 'function') throw new Error('DOCX native header/footer consumer exports are missing')`,
      `if (typeof docxCompiler.prepareNativeDocxPagePaintV1 !== 'function' || typeof docxCompiler.qualifyNativeDocxTablesV1 !== 'function') throw new Error('DOCX table page-paint compiler exports are unavailable')`,
      `if (typeof docs.nativeDocxPagePaintPaginatedLayoutSha256V1 !== 'function' || typeof docxCompiler.completeNativeDocxPagePaintV1 !== 'function') throw new Error('DOCX note page-paint consumer exports are incomplete')`,
      `const adapterResult = docs.adaptNativeDocxProseMirrorTransactionV1(${JSON.stringify(docsAdapterDocument)}, ${JSON.stringify(docsAdapterTransaction)})`,
      `if (!adapterResult.ok || adapterResult.value.envelope.payload.mutations[0]?.text !== 'Strict native') throw new Error('DOCX transaction adapter tarball smoke failed')`,
    ] : []),
    ...(packageNames.includes('@injoffice/pdf') ? [`await import('@injoffice/pdf/browser')`] : []),
    ...(packageNames.includes('@injoffice/font-metrics') ? [
      `await import('@injoffice/font-metrics/layout')`,
      `await import('@injoffice/font-metrics/harfbuzz')`,
      `const bidi = await import('@injoffice/font-metrics/bidi')`,
      `const bidiResult = bidi.resolveNativeBidiParagraphV1({ text: 'abc אב 12', baseDirection: 'ltr' })`,
      `if (!bidiResult.ok || bidiResult.value.providerId !== 'injoffice.bidi-js' || JSON.stringify(bidiResult.value.levels) !== JSON.stringify([0,0,0,0,1,1,1,2,2])) throw new Error('bidi subpath did not execute its exact pinned resolver')`,
      `if (bidi.resolveNativeBidiParagraphV1({ text: 'א😀ב', baseDirection: 'rtl' }).code !== 'unsupported-scalar') throw new Error('bidi subpath did not atomically refuse astral input')`,
      `const unicode13 = await import('@injoffice/font-metrics/unicode13')`,
      `if (!unicode13.UNICODE_13_TABLES_RUNTIME_MATCH || unicode13.unicode13Script(0x16FE2) !== 'Zyyy' || unicode13.unicode13Script(0x16FE3) !== 'Zyyy' || unicode13.unicode13Punctuation(0x16FE2) !== 'close' || unicode13.unicode13Punctuation(0x16FE3) !== 'none') throw new Error('Unicode 13 classifier vectors drifted')`,
    ] : []),
    ...(packageNames.includes('@injoffice/slides') && packageNames.includes('@injoffice/pptx-authored') ? [
      `const { compileDeckToWire } = await import('@injoffice/slides/authoring')`,
      `const { compileWireDeckToNativeV1 } = await import('@injoffice/pptx-authored')`,
      `if (typeof compileDeckToWire !== 'function') throw new Error('slides authoring subpath is unavailable')`,
      `const authored = compileWireDeckToNativeV1({ slides: [{ background: '#FFFFFF', shapes: [] }] })`,
      `if (!authored.ok || authored.deck.contractVersion !== 'pptx-native/v1') throw new Error('native authored compiler consumer smoke failed')`,
    ] : []),
    ...(packageNames.includes('@injoffice/sheets') ? [
      `const sheets = await import('@injoffice/sheets')`,
      `if (typeof sheets.projectNativeWorkbookV1 !== 'function' || typeof sheets.compileNativeSheetDecorationsV1 !== 'function' || typeof sheets.emitNativeSheetDecorationCommandsV1 !== 'function') throw new Error('XLSX native projection/decoration consumer exports are missing')`,
      `const nativeWorkbook = ${JSON.stringify(sheetsAdapterWorkbook)}`,
      `const packageDigest = nativeWorkbook.source.package_sha256`,
      `const revision = nativeWorkbook.revision`,
      `const fontRecordDigest = 'sha256:' + 'c'.repeat(64)`,
      `const model = sheets.projectNativeWorkbookV1(nativeWorkbook)`,
      `const metricAuthority = { source_revision: revision, source_package_sha256: packageDigest, normal_style_xf_id: 0, normal_style_font_id: 0, font_name: 'Fixture', font_size_points: 11, font_bold: false, font_italic: false, normal_font_record_sha256: fontRecordDigest, font_sha256: 'sha256:' + 'd'.repeat(64), provider_id: 'consumer-smoke', provider_revision: '1', measurement_dpi: 96, maximum_digit_width_pixels: 7 }`,
      `const geometry = sheets.compileNativeSheetGeometryV1(model, '7', { row: 0, column: 0, end_row: 0, end_column: 0 }, metricAuthority)`,
      `const decorations = sheets.compileNativeSheetDecorationsV1(model, geometry)`,
      `const decorationSurface = sheets.createNativeSheetDecorationRecordingSurfaceV1()`,
      `sheets.emitNativeSheetDecorationCommandsV1(decorations, decorationSurface)`,
      `if (!Object.isFrozen(model) || decorations.border_segments.length !== 0 || decorationSurface.finish().at(-1)?.kind !== 'endDecorations') throw new Error('XLSX native decoration consumer smoke failed')`,
    ] : []),
    ...(packageNames.includes('@injoffice/xlsx-wasm') ? [
      `const xlsxWasm = await import('@injoffice/xlsx-wasm')`,
      `if (typeof xlsxWasm.createXlsxWasmClient !== 'function' || typeof xlsxWasm.resolveXlsxWasmAssetUrls !== 'function') throw new Error('XLSX WASM consumer exports are missing')`,
      `const { existsSync: xlsxAssetExists } = await import('node:fs')`,
      `const { fileURLToPath: xlsxAssetPath } = await import('node:url')`,
      `const xlsxAssets = xlsxWasm.resolveXlsxWasmAssetUrls()`,
      `for (const assetUrl of Object.values(xlsxAssets)) if (!xlsxAssetExists(xlsxAssetPath(assetUrl))) throw new Error('XLSX WASM packaged asset is missing: ' + assetUrl)`,
    ] : []),
    ...(packageNames.includes('@injoffice/pptx-wasm') ? [
      `const pptxWasm = await import('@injoffice/pptx-wasm')`,
      `if (typeof pptxWasm.createPptxWasmClient !== 'function' || typeof pptxWasm.resolvePptxWasmAssetUrls !== 'function') throw new Error('PPTX WASM consumer exports are missing')`,
      `const { existsSync: pptxAssetExists } = await import('node:fs')`,
      `const { fileURLToPath: pptxAssetPath } = await import('node:url')`,
      `for (const assetUrl of Object.values(pptxWasm.resolvePptxWasmAssetUrls())) if (!pptxAssetExists(pptxAssetPath(assetUrl))) throw new Error('PPTX WASM packaged asset is missing: ' + assetUrl)`,
    ] : []),
    ...(packageNames.includes('@injoffice/docx-wasm') ? [
      `const docxWasm = await import('@injoffice/docx-wasm')`,
      `if (typeof docxWasm.createDocxWasmClient !== 'function' || typeof docxWasm.resolveDocxWasmAssetUrls !== 'function') throw new Error('DOCX WASM consumer exports are missing')`,
      `const { existsSync: docxAssetExists } = await import('node:fs')`,
      `const { fileURLToPath: docxAssetPath } = await import('node:url')`,
      `const docxAssets = docxWasm.resolveDocxWasmAssetUrls()`,
      `for (const assetUrl of Object.values(docxAssets)) if (!docxAssetExists(docxAssetPath(assetUrl))) throw new Error('DOCX WASM packaged asset is missing: ' + assetUrl)`,
    ] : []),
    `console.log('Imported ${packageNames.length} packages from tarballs.')`,
    '',
  ].join('\n'))
  writeFileSync(resolve(temporary, 'smoke-types.mts'), [
    ...packageEntrypoints.map((name, index) => `import * as entry${index} from ${JSON.stringify(name)}; void entry${index}`),
    '',
  ].join('\n'))
  writeFileSync(resolve(temporary, 'tsconfig.json'), JSON.stringify({
    compilerOptions: {
      target: 'ES2022',
      module: 'NodeNext',
      moduleResolution: 'NodeNext',
      lib: ['ES2022', 'DOM', 'DOM.Iterable'],
      strict: true,
      noEmit: true,
      // Third-party Univer/PDF declarations currently require consumer-specific
      // ambient CSS and Emscripten types. Internal InjOffice declaration links
      // are checked separately by check-packages.
      skipLibCheck: true,
    },
    files: ['smoke-types.mts'],
  }, null, 2))

  run(process.platform === 'win32' ? 'npm.cmd' : 'npm', [
    'install',
    '--ignore-scripts',
    '--no-package-lock',
    '--cache',
    cache,
    ...tarballs,
  ], temporary, npmEnvironment)
  process.stdout.write(run(process.execPath, ['smoke.mjs'], temporary))
  run(process.execPath, [resolve(root, 'node_modules/typescript/bin/tsc'), '-p', 'tsconfig.json'], temporary)
  console.log(`Type-checked ${packageEntrypoints.length} package entrypoints from tarballs with TypeScript NodeNext resolution.`)

  const browserPackages = [
    { packageName: '@injoffice/xlsx-wasm', entry: 'xlsx-browser.mjs', resolver: 'resolveXlsxWasmAssetUrls', globalName: '__xlsxWasmAssets', assetName: 'xlsxnative', label: 'XLSX' },
    { packageName: '@injoffice/pptx-wasm', entry: 'pptx-browser.mjs', resolver: 'resolvePptxWasmAssetUrls', globalName: '__pptxWasmAssets', assetName: 'pptxnative', label: 'PPTX' },
    { packageName: '@injoffice/docx-wasm', entry: 'docx-browser.mjs', resolver: 'resolveDocxWasmAssetUrls', globalName: '__docxWasmAssets', assetName: 'docxnative', label: 'DOCX' },
  ].filter(({ packageName }) => packageNames.includes(packageName))
  for (const browserPackage of browserPackages) {
    writeFileSync(resolve(temporary, 'index.html'), `<script type="module" src="/${browserPackage.entry}"></script>\n`)
    writeFileSync(resolve(temporary, browserPackage.entry), [
      `import { ${browserPackage.resolver} } from '${browserPackage.packageName}'`,
      `globalThis.${browserPackage.globalName} = ${browserPackage.resolver}()`,
      '',
    ].join('\n'))
    const browserOut = resolve(temporary, 'browser-dist')
    run(process.execPath, [resolve(root, 'node_modules/vite/bin/vite.js'), 'build', '--outDir', browserOut], temporary)
    const browserFiles = listFiles(browserOut)
    if (!browserFiles.some((path) => path.endsWith('.wasm') && path.includes(browserPackage.assetName))) throw new Error(`Vite consumer did not emit ${browserPackage.assetName}.wasm`)
    const emittedJavaScript = browserFiles
      .filter((path) => path.endsWith('.js'))
      .map((path) => readFileSync(path, 'utf8'))
    if (!emittedJavaScript.some((source) => source.includes('injoffice.native-wasm-worker'))) throw new Error(`Vite consumer did not emit the ${browserPackage.label} worker`)
    if (!emittedJavaScript.some((source) => source.includes('globalThis.Go'))) throw new Error('Vite consumer did not emit the matching Go runtime')
    if (emittedJavaScript.some((source) => /from ["']node:/.test(source))) throw new Error(`Vite ${browserPackage.label} WASM consumer emitted a Node builtin import`)
    console.log(`Bundled ${browserPackage.label} WASM assets from the installed tarball with Vite.`)
  }
} finally {
  rmSync(temporary, { recursive: true, force: true })
}
