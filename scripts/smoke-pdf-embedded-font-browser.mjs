import assert from 'node:assert/strict'
import { existsSync, mkdtempSync, readFileSync, readdirSync, writeFileSync } from 'node:fs'
import { createRequire } from 'node:module'
import { tmpdir } from 'node:os'
import { resolve, basename } from 'node:path'
import { decodePDFRawStream, PDFDocument, PDFName, PDFHexString } from 'pdf-lib'
import { launchChromeForCDP, terminateProcess } from './chrome-cdp-startup.mjs'
import { startShowcaseDevServer } from './showcase-smoke-dev-server.mjs'

// Exercises the real Vite/browser fontkit entry: Node unit tests cannot catch
// its default-only ESM export disagreeing with the package's named typings.
const require = createRequire(import.meta.url)
const output = mkdtempSync(resolve(tmpdir(), 'injoffice-pdf-font-smoke-'))
const fontPath = require.resolve('dejavu-fonts-ttf/ttf/DejaVuSans.ttf')
const sourcePath = resolve(output, 'source.pdf')
const doc = await PDFDocument.create()
const field = doc.getForm().createTextField('Unicode sample')
field.setText('BEFORE')
field.addToPage(doc.addPage([400, 300]), { x: 20, y: 180, width: 350, height: 80 })
field.setFontSize(24)
const source = await doc.save()
writeFileSync(sourcePath, source)

let server, chrome, socket, sequence = 0
const pending = new Map()
const errors = []
const postRequests = []
let replacingDownloadedSource = false
let expectedFontName
const panel = '[data-demo-surface="pdf"] #pdf-operation-panel'
const textInput = `${panel} [aria-label="Form value: Unicode sample"]`
const appearanceSelect = `${panel} [aria-label="Saved text appearance"]`
const pause = milliseconds => new Promise(resolve => setTimeout(resolve, milliseconds))
function send(method, params = {}) {
  return new Promise((resolve, reject) => {
    const id = ++sequence
    const timer = setTimeout(() => { pending.delete(id); reject(new Error(`CDP timed out: ${method}`)) }, 30_000)
    pending.set(id, { resolve, reject, timer })
    socket.send(JSON.stringify({ id, method, params }))
  })
}
async function evaluate(expression) {
  const result = await send('Runtime.evaluate', { expression, returnByValue: true, awaitPromise: true, userGesture: true })
  if (result.exceptionDetails) throw new Error(result.exceptionDetails.exception?.description ?? result.exceptionDetails.text)
  return result.result.value
}
async function until(expression, label) {
  const deadline = Date.now() + 60_000
  do {
    if (await evaluate(expression)) return
    const error = await evaluate(`document.querySelector('[data-demo-surface="pdf"] [role="alert"]')?.textContent`)
    if (error) throw new Error(`${label}: ${error}`)
    await pause(100)
  } while (Date.now() < deadline)
  throw new Error(`Timed out: ${label}\n${JSON.stringify(await evaluate(`(() => {
    const panel = document.querySelector(${JSON.stringify(panel)});
    const button = [...(panel?.querySelectorAll('button') ?? [])].find(b => b.textContent.trim() === 'Apply form values');
    return {text: panel?.textContent, applyDisabled: button?.disabled,
      input: document.querySelector(${JSON.stringify(textInput)})?.value,
      inputDisabled: document.querySelector(${JSON.stringify(textInput)})?.disabled,
      appearance: document.querySelector(${JSON.stringify(appearanceSelect)})?.value,
      appearanceDisabled: document.querySelector(${JSON.stringify(appearanceSelect)})?.disabled,
      operation: window.__pdfFontApply};
  })()`))}`)
}
async function upload(selector, file) {
  if (selector === '[aria-label="Embedded appearance font"]') expectedFontName = basename(file)
  const { result } = await send('Runtime.evaluate', { expression: `document.querySelector(${JSON.stringify(selector)})` })
  assert.ok(result.objectId, `file input exists: ${selector}`)
  await send('DOM.setFileInputFiles', { objectId: result.objectId, files: [file] })
  await send('Runtime.releaseObject', { objectId: result.objectId })
}
const setValue = (selector, value, select = false) => evaluate(`(() => {
  const input = document.querySelector(${JSON.stringify(selector)});
  Object.getOwnPropertyDescriptor(${select ? 'HTMLSelectElement' : 'HTMLInputElement'}.prototype, 'value').set.call(input, ${JSON.stringify(value)});
  input.dispatchEvent(new Event(${JSON.stringify(select ? 'change' : 'input')}, { bubbles: true }));
})()`)
async function apply(expectedValue) {
  // Readiness and the one click share a browser task: a render cannot disable
  // the button between our check and dispatch. Never retry an issued click.
  await until(`(() => {
    const panel = document.querySelector(${JSON.stringify(panel)});
    const button = [...(panel?.querySelectorAll('button') ?? [])].find(b => b.textContent.trim() === 'Apply form values');
    const input = document.querySelector(${JSON.stringify(textInput)});
    const appearance = document.querySelector(${JSON.stringify(appearanceSelect)});
    if (!button || button.disabled || !input || input.disabled || input.value !== ${JSON.stringify(expectedValue)} ||
        !appearance || appearance.disabled || appearance.value !== 'embedded' ||
        !panel.textContent.includes(${JSON.stringify('Selected: ' + expectedFontName)})) return false;
    const previous = new Set(panel.querySelectorAll('[role="status"]'));
    window.__pdfFontApply = {clicked: true, freshResult: false, expectedValue: ${JSON.stringify(expectedValue)}};
    const observer = new MutationObserver(() => {
      const result = [...panel.querySelectorAll('[role="status"]')].find(node =>
        !previous.has(node) && /Generated 1 widget appearance|No form values applied[.]/.test(node.textContent));
      if (result) {
        window.__pdfFontApply.freshResult = true;
        window.__pdfFontApply.result = result.textContent;
        observer.disconnect();
      }
    });
    observer.observe(panel, {childList: true, subtree: true, characterData: true});
    button.click();
    return true;
  })()`, 'expected draft/font and enabled Apply');
  // The old summary may have identical text. Require a newly mounted result
  // notice from this operation before any caller checks its contents.
  await until('window.__pdfFontApply?.freshResult === true', 'fresh Apply result');
}

try {
  server = await startShowcaseDevServer(resolve(import.meta.dirname, '../apps/playground'))
  chrome = await launchChromeForCDP({
    executable: process.env.CHROME_BIN ?? ['/Applications/Google Chrome.app/Contents/MacOS/Google Chrome', '/usr/bin/google-chrome', '/usr/bin/chromium', '/usr/bin/chromium-browser'].find(existsSync) ?? 'google-chrome',
    createProfile: () => mkdtempSync(resolve(tmpdir(), 'injoffice-pdf-font-chrome-')),
  })
  socket = new WebSocket(chrome.target.webSocketDebuggerUrl)
  await new Promise((resolve, reject) => { socket.addEventListener('open', resolve, { once: true }); socket.addEventListener('error', reject, { once: true }) })
  socket.addEventListener('message', ({ data }) => {
    const message = JSON.parse(data)
    if (message.id) {
      const task = pending.get(message.id)
      if (!task) return
      pending.delete(message.id)
      clearTimeout(task.timer)
      if (message.error) task.reject(new Error(message.error.message))
      else task.resolve(message.result)
    } else if (message.method === 'Runtime.exceptionThrown') {
      errors.push(message.params.exceptionDetails.exception?.description ?? message.params.exceptionDetails.text)
    } else if (message.method === 'Page.javascriptDialogOpening') {
      const expected = replacingDownloadedSource && message.params.type === 'confirm' && message.params.message === 'Open another PDF and discard your current edits? Download your edited PDF first if you want to keep it.'
      if (!expected) errors.push('Unexpected browser dialog: ' + message.params.message)
      void send('Page.handleJavaScriptDialog', { accept: expected }).catch(error => errors.push(error.message))
    } else if (message.method === 'Fetch.requestPaused') {
      const { requestId, request } = message.params
      const url = new URL(request.url)
      if (request.method === 'POST') postRequests.push(request.url)
      const allowed = request.method !== 'POST' && (url.origin === new URL(server.url).origin || ['blob:', 'data:'].includes(url.protocol))
      void send(allowed ? 'Fetch.continueRequest' : 'Fetch.failRequest', { requestId, ...(!allowed ? { errorReason: 'BlockedByClient' } : {}) }).catch(error => errors.push(error.message))
    }
  })
  await send('Runtime.enable')
  await send('Page.enable')
  await send('Fetch.enable', { patterns: [{ urlPattern: '*' }] })
  await send('Browser.setDownloadBehavior', { behavior: 'allow', downloadPath: output })
  await send('Page.navigate', { url: `${server.url}#/pdf?feature=editor` })
  await until(`!!document.querySelector('[aria-label="Open a PDF file"]')`, 'PDF editor loads')
  await upload('[aria-label="Open a PDF file"]', sourcePath)
  await evaluate(`Array.from(document.querySelectorAll('[data-demo-surface="pdf"] [role="tab"]')).find(tab => tab.textContent.trim() === 'Forms').click()`)
  await until(`document.querySelector(${JSON.stringify(textInput)})?.value === 'BEFORE' && !document.querySelector(${JSON.stringify(appearanceSelect)})?.disabled`, 'uploaded source replaces initial sample')
  assert.equal(await evaluate(`document.querySelector(${JSON.stringify(appearanceSelect)}).value`), 'viewer')
  await setValue(appearanceSelect, 'embedded', true)
  await until(`!!document.querySelector('[aria-label="Embedded appearance font"]')`, 'embedded font input')
  await upload('[aria-label="Embedded appearance font"]', fontPath)
  await until(`document.querySelector(${JSON.stringify(panel)}).textContent.includes('Selected: DejaVuSans.ttf')`, 'local font read')
  await setValue(textInput, 'AV café Ω Ж 😀')
  await apply('AV café Ω Ж 😀')
  await until(`document.querySelector(${JSON.stringify(panel)}).textContent.includes('Generated 1 widget appearance')`, 'browser fontkit creates embedded appearance')
  await until(`document.querySelector(${JSON.stringify(textInput)})?.value === 'AV café Ω Ж 😀' && !document.querySelector(${JSON.stringify(appearanceSelect)})?.disabled`, 'saved document finishes reloading')
  await evaluate(`Array.from(document.querySelectorAll('[data-demo-surface="pdf"] button')).find(button => button.textContent.trim() === 'Download edited PDF').click()`)
  let downloaded
  for (let attempt = 0; attempt < 300; attempt++) {
    downloaded = readdirSync(output).find(name => name.endsWith('.pdf') && name !== 'source.pdf')
    if (downloaded) break
    await pause(100)
  }
  assert.ok(downloaded, 'UI produces downloaded PDF bytes')
  const saved = await PDFDocument.load(readFileSync(resolve(output, downloaded)))
  const savedField = saved.getForm().getTextField('Unicode sample')
  assert.equal(savedField.getText(), 'AV café Ω Ж 😀')
  const ap = saved.context.lookup(savedField.acroField.getWidgets()[0].getNormalAppearance())
  const artwork = Buffer.from(decodePDFRawStream(ap).decode()).toString('latin1')
  assert.match(artwork, /63\.96484375/, 'AV pair kerning is saved as a PDF spacing adjustment')
  assert.match(artwork, /\] TJ/, 'browser emits positioned text rather than nominal glyph widths')
  const fonts = ap.dict.lookup(PDFName.of('Resources')).lookup(PDFName.of('Font'))
  const font = fonts.lookup(fonts.keys()[0])
  assert.equal(String(font.get(PDFName.of('Subtype'))), '/Type0')
  const cid = font.lookup(PDFName.of('DescendantFonts')).lookup(0)
  const fontFile = cid.lookup(PDFName.of('FontDescriptor')).lookup(PDFName.of('FontFile2'))
  const embeddedBytes = decodePDFRawStream(fontFile).decode()
  assert.deepEqual(Buffer.from(embeddedBytes), readFileSync(fontPath), 'saved font retains the complete fixed TrueType face')
  const cmap = Buffer.from(decodePDFRawStream(font.lookup(PDFName.of('ToUnicode'))).decode()).toString('latin1')
  assert.match(cmap, /<D83DDE00>/, 'supplementary scalar survives ToUnicode')
  const shapedValue = 'e\u0301 ffi a\u0301\u0323'
  await setValue(textInput, shapedValue)
  await apply(shapedValue)
  await until(`document.querySelector(${JSON.stringify(panel)}).textContent.includes('Generated 1 widget appearance')`, 'browser HarfBuzz creates positioned cluster appearance')
  await until(`document.querySelector(${JSON.stringify(textInput)})?.value === ${JSON.stringify(shapedValue)} && !document.querySelector(${JSON.stringify(appearanceSelect)})?.disabled`, 'shaped saved document finishes reloading')
  const shapedOutput = mkdtempSync(resolve(output, 'shaped-'))
  await send('Browser.setDownloadBehavior', { behavior: 'allow', downloadPath: shapedOutput })
  await evaluate(`Array.from(document.querySelectorAll('[data-demo-surface="pdf"] button')).find(button => button.textContent.trim() === 'Download edited PDF').click()`)
  let shapedDownload
  for (let attempt = 0; attempt < 300; attempt++) {
    shapedDownload = readdirSync(shapedOutput).find(name => name.endsWith('.pdf'))
    if (shapedDownload) break
    await pause(100)
  }
  assert.ok(shapedDownload, 'UI downloads shaped source text')
  const shapedDoc = await PDFDocument.load(readFileSync(resolve(shapedOutput, shapedDownload)))
  const shapedField = shapedDoc.getForm().getTextField('Unicode sample')
  assert.equal(shapedField.getText(), shapedValue, 'public browser form retains exact logical source')
  const shapedAp = shapedDoc.context.lookup(shapedField.acroField.getWidgets()[0].getNormalAppearance())
  const shapedArtwork = Buffer.from(decodePDFRawStream(shapedAp).decode()).toString('latin1')
  assert.match(shapedArtwork, /ActualText/, 'unpartitionable continuation has exact replacement text')
  shapedDoc.getForm().flatten({ updateFieldAppearances: false })
  const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs')
  const extraction = pdfjs.getDocument({ data: await shapedDoc.save() })
  try {
    const page = await (await extraction.promise).getPage(1)
    const content = await page.getTextContent({ disableNormalization: true })
    assert.ok(content.items.map(item => 'str' in item ? item.str : '').join('').includes(shapedValue), 'saved appearance extracts exact decomposed clusters and ligatures')
  } finally { await extraction.destroy() }
  const rtlValue = 'abc (السَّلَام 123) xyz'
  await setValue(textInput, rtlValue)
  await apply(rtlValue)
  await until(`document.querySelector(${JSON.stringify(panel)}).textContent.includes('Generated 1 widget appearance')`, 'browser creates contextual RTL appearance')
  await until(`document.querySelector(${JSON.stringify(textInput)})?.value === ${JSON.stringify(rtlValue)} && !document.querySelector(${JSON.stringify(appearanceSelect)})?.disabled`, 'RTL saved document finishes reloading')
  const rtlOutput = mkdtempSync(resolve(output, 'rtl-'))
  await send('Browser.setDownloadBehavior', { behavior: 'allow', downloadPath: rtlOutput })
  await evaluate(`Array.from(document.querySelectorAll('[data-demo-surface="pdf"] button')).find(button => button.textContent.trim() === 'Download edited PDF').click()`)
  let rtlDownload
  for (let attempt = 0; attempt < 300; attempt++) {
    rtlDownload = readdirSync(rtlOutput).find(name => name.endsWith('.pdf'))
    if (rtlDownload) break
    await pause(100)
  }
  assert.ok(rtlDownload, 'UI downloads contextual mixed RTL text')
  const rtlDoc = await PDFDocument.load(readFileSync(resolve(rtlOutput, rtlDownload)))
  const rtlField = rtlDoc.getForm().getTextField('Unicode sample')
  assert.equal(rtlField.getText(), rtlValue, 'browser export retains exact logical RTL form value')
  const rtlAp = rtlDoc.context.lookup(rtlField.acroField.getWidgets()[0].getNormalAppearance())
  assert.match(Buffer.from(decodePDFRawStream(rtlAp).decode()).toString('latin1'), /ActualText/, 'RTL appearance carries standard logical replacement text')
  const scriptCases = [
    ['NotoSansDevanagari', 'क्षि नमस्ते'], ['NotoSansBengali', 'ক্ষি বাংলা'],
    ['NotoSansThai', 'น้ำ ภาษาไทย'], ['NotoSansKhmer', 'ខ្មែរ'],
    ['NotoSansMyanmar', 'မြန်မာ'], ['NotoSansSyriac', 'ܫܠܡܐ'],
    ['NotoSansArabic', '\u0600بب'],
  ]
  for (const [fontName, value] of scriptCases) {
    const fixtureFont = resolve(import.meta.dirname, `../packages/pdf/testdata/fonts/${fontName}-Regular.ttf`)
    await upload('[aria-label="Embedded appearance font"]', fixtureFont)
    await until(`document.querySelector(${JSON.stringify(panel)}).textContent.includes(${JSON.stringify(`Selected: ${fontName}-Regular.ttf`)})`, 'local script font read')
    await setValue(textInput, value)
    await apply(value)
    await until(`document.querySelector(${JSON.stringify(panel)}).textContent.includes('Generated 1 widget appearance')`, 'browser creates contextual script appearance')
    await until(`document.querySelector(${JSON.stringify(textInput)})?.value === ${JSON.stringify(value)} && !document.querySelector(${JSON.stringify(appearanceSelect)})?.disabled`, 'script document finishes reloading')
    const scriptOutput = mkdtempSync(resolve(output, `${fontName}-`))
    await send('Browser.setDownloadBehavior', { behavior: 'allow', downloadPath: scriptOutput })
    await evaluate(`Array.from(document.querySelectorAll('[data-demo-surface="pdf"] button')).find(button => button.textContent.trim() === 'Download edited PDF').click()`)
    let scriptDownload
    for (let attempt = 0; attempt < 300; attempt++) {
      scriptDownload = readdirSync(scriptOutput).find(name => name.endsWith('.pdf'))
      if (scriptDownload) break
      await pause(100)
    }
    assert.ok(scriptDownload, `UI downloads ${fontName} appearance`)
    const scriptDoc = await PDFDocument.load(readFileSync(resolve(scriptOutput, scriptDownload)))
    assert.equal(scriptDoc.getForm().getTextField('Unicode sample').getText(), value, 'script export retains exact logical form value')
  }
  for (const [filename, value] of [['NotoSansDevanagari-Regular.otf', 'क्षि नमस्ते'], ['NotoSansJP-CID-subset.otf', 'Aé Ω 日本語 侮侮\uFE00 倦倦\u{E0100}'], ['NotoSansKR-CID-subset.otf', '한글 한글']]) {
    const cffPath = resolve(import.meta.dirname, '../packages/pdf/testdata/fonts', filename)
    const originalFont = readFileSync(cffPath)
    await upload('[aria-label="Embedded appearance font"]', cffPath)
    await until(`document.querySelector(${JSON.stringify(panel)}).textContent.includes(${JSON.stringify('Selected: ' + filename)})`, 'CFF font upload')
    await setValue(textInput, value)
    await apply(value)
    await until(`document.querySelector(${JSON.stringify(panel)}).textContent.includes('Generated 1 widget appearance')`, 'CFF font saves contextual appearance')
    await until(`document.querySelector(${JSON.stringify(textInput)})?.value === ${JSON.stringify(value)} && !document.querySelector(${JSON.stringify(appearanceSelect)})?.disabled`, 'CFF document finishes reloading')
    const cffOutput = mkdtempSync(resolve(output, filename + '-'))
    await send('Browser.setDownloadBehavior', { behavior: 'allow', downloadPath: cffOutput })
    await evaluate(`Array.from(document.querySelectorAll('[data-demo-surface="pdf"] button')).find(button => button.textContent.trim() === 'Download edited PDF').click()`)
    let cffDownload
    for (let attempt = 0; attempt < 300; attempt++) {
      cffDownload = readdirSync(cffOutput).find(name => name.endsWith('.pdf'))
      if (cffDownload) break
      await pause(100)
    }
    assert.ok(cffDownload, 'UI downloads CFF appearance')
    const cffDoc = await PDFDocument.load(readFileSync(resolve(cffOutput, cffDownload)))
    const cffField = cffDoc.getForm().getTextField('Unicode sample')
    assert.equal(cffField.getText(), value)
    const cffAp = cffDoc.context.lookup(cffField.acroField.getWidgets()[0].getNormalAppearance())
    const cffFonts = cffAp.dict.lookup(PDFName.of('Resources')).lookup(PDFName.of('Font'))
    const cffFont = cffFonts.lookup(cffFonts.keys()[0])
    const cffCid = cffFont.lookup(PDFName.of('DescendantFonts')).lookup(0)
    assert.equal(String(cffCid.get(PDFName.of('Subtype'))), '/CIDFontType0')
    assert.equal(cffCid.has(PDFName.of('CIDToGIDMap')), false)
    const cffDescriptor = cffCid.lookup(PDFName.of('FontDescriptor'))
    assert.equal(cffDescriptor.has(PDFName.of('FontFile2')), false)
    const program = cffDescriptor.lookup(PDFName.of('FontFile3'))
    assert.equal(String(program.dict.get(PDFName.of('Subtype'))), '/CIDFontType0C')
    assert.equal(decodePDFRawStream(program).decode()[0], 1)
    assert.match(Buffer.from(decodePDFRawStream(cffFont.lookup(PDFName.of('Encoding'))).decode()).toString(), /begincidchar/)
    if (filename === 'NotoSansKR-CID-subset.otf') {
      assert.equal(cffFonts.keys().length, 2, 'composed/decomposed aliases have separate Type0 font views')
      const views = cffFonts.keys().map(name => cffFonts.lookup(name))
      assert.equal(views[0].get(PDFName.of('DescendantFonts')).toString(), views[1].get(PDFName.of('DescendantFonts')).toString(), 'views share the same embedded font program')
    }
    assert.deepEqual(readFileSync(cffPath), originalFont)
  }
  const collectionPath = resolve(import.meta.dirname, '../packages/pdf/testdata/fonts/NotoSans-Devanagari-Bengali.ttc')
  const originalCollection = readFileSync(collectionPath)
  await upload('[aria-label="Embedded appearance font"]', collectionPath)
  const faceInput = '[aria-label="Collection face index"]'
  await until(`document.querySelector(${JSON.stringify(faceInput)})?.value === '0'`, 'collection face selection defaults to zero')
  await setValue(faceInput, '2')
  assert.equal(await evaluate(`Array.from(document.querySelectorAll(${JSON.stringify(`${panel} button`)})).find(button => button.textContent.trim() === 'Apply form values').disabled`), true, 'out-of-range collection face cannot apply')
  await setValue(faceInput, '1')
  const collectionValue = 'ক্ষি বাংলা'
  await setValue(textInput, collectionValue)
  await apply(collectionValue)
  await until(`document.querySelector(${JSON.stringify(panel)}).textContent.includes('Generated 1 widget appearance')`, 'browser embeds selected second collection face')
  await until(`document.querySelector(${JSON.stringify(textInput)})?.value === ${JSON.stringify(collectionValue)} && !document.querySelector(${JSON.stringify(appearanceSelect)})?.disabled`, 'collection document finishes reloading')
  const collectionOutput = mkdtempSync(resolve(output, 'collection-'))
  await send('Browser.setDownloadBehavior', { behavior: 'allow', downloadPath: collectionOutput })
  await evaluate(`Array.from(document.querySelectorAll('[data-demo-surface="pdf"] button')).find(button => button.textContent.trim() === 'Download edited PDF').click()`)
  let collectionDownload
  for (let attempt = 0; attempt < 300; attempt++) {
    collectionDownload = readdirSync(collectionOutput).find(name => name.endsWith('.pdf'))
    if (collectionDownload) break
    await pause(100)
  }
  assert.ok(collectionDownload, 'UI downloads selected collection face')
  const collectionDoc = await PDFDocument.load(readFileSync(resolve(collectionOutput, collectionDownload)))
  const collectionField = collectionDoc.getForm().getTextField('Unicode sample')
  assert.equal(collectionField.getText(), collectionValue)
  const collectionAp = collectionDoc.context.lookup(collectionField.acroField.getWidgets()[0].getNormalAppearance())
  const collectionFonts = collectionAp.dict.lookup(PDFName.of('Resources')).lookup(PDFName.of('Font'))
  const collectionFont = collectionFonts.lookup(collectionFonts.keys()[0])
  const collectionCid = collectionFont.lookup(PDFName.of('DescendantFonts')).lookup(0)
  const collectionProgram = decodePDFRawStream(collectionCid.lookup(PDFName.of('FontDescriptor')).lookup(PDFName.of('FontFile2'))).decode()
  assert.equal(new DataView(collectionProgram.buffer, collectionProgram.byteOffset, collectionProgram.byteLength).getUint32(0), 0x10000, 'PDF embeds standalone TrueType, not a TTC wrapper')
  assert.deepEqual(readFileSync(collectionPath), originalCollection)
  await upload('[aria-label="Embedded appearance font"]', fontPath)
  await until(`!document.querySelector(${JSON.stringify(faceInput)}) && document.querySelector(${JSON.stringify(panel)}).textContent.includes('Selected: DejaVuSans.ttf')`, 'standalone font resets collection selection')
  await upload('[aria-label="Embedded appearance font"]', collectionPath)
  await until(`document.querySelector(${JSON.stringify(faceInput)})?.value === '0'`, 'new font upload resets face index')
  await setValue(textInput, '\u{10FFFF}')
  await apply('\u{10FFFF}')
  await until(`document.querySelector(${JSON.stringify(panel)}).textContent.includes('No form values applied.')`, 'missing glyph is refused')
  // True automatic-size source: setting DA0 after widget creation avoids
  // pdf-lib's initial default-appearance generation replacing the authored zero.
  const autoDoc = await PDFDocument.create(), autoField = autoDoc.getForm().createTextField('Unicode sample')
  autoField.setText('BEFORE'); autoField.addToPage(autoDoc.addPage([400, 300]), { x: 20, y: 180, width: 350, height: 80 }); autoField.setFontSize(0)
  const autoPath = resolve(output, 'automatic-source.pdf')
  writeFileSync(autoPath, await autoDoc.save({ updateFieldAppearances: false }))
  replacingDownloadedSource = true
  await upload('[aria-label="Open a PDF file"]', autoPath)
  replacingDownloadedSource = false
  await until(`document.querySelector(${JSON.stringify(textInput)})?.value === 'BEFORE' && !document.querySelector(${JSON.stringify(appearanceSelect)})?.disabled`, 'automatic source loads and clears draft')
  await setValue(appearanceSelect, 'embedded', true)
  await until(`!!document.querySelector('[aria-label="Embedded appearance font"]')`, 'automatic embedded input')
  await upload('[aria-label="Embedded appearance font"]', require.resolve('dejavu-fonts-ttf/ttf/DejaVuSans-Oblique.ttf'))
  await until(`document.querySelector(${JSON.stringify(panel)}).textContent.includes('Selected: DejaVuSans-Oblique.ttf')`, 'italic font upload')
  for (const [label, value] of [['italic-auto', 'j'], ['glyphless', '\u200B\u00AD\u{E0100}']]) {
    await setValue(textInput, value)
    await apply(value)
    await until(`document.querySelector(${JSON.stringify(panel)}).textContent.includes('Generated 1 widget appearance')`, 'automatic appearance generated')
    await until(`document.querySelector(${JSON.stringify(textInput)})?.value === ${JSON.stringify(value)} && !document.querySelector(${JSON.stringify(appearanceSelect)})?.disabled`, 'automatic saved document reloads')
    const destination = mkdtempSync(resolve(output, label + '-'))
    await send('Browser.setDownloadBehavior', { behavior: 'allow', downloadPath: destination })
    await evaluate(`Array.from(document.querySelectorAll('[data-demo-surface="pdf"] button')).find(button => button.textContent.trim() === 'Download edited PDF').click()`)
    let file
    for (let attempt = 0; attempt < 300; attempt++) { file = readdirSync(destination).find(name => name.endsWith('.pdf')); if (file) break; await pause(100) }
    assert.ok(file, 'UI downloads completed appearance')
    const saved = await PDFDocument.load(readFileSync(resolve(destination, file))), field = saved.getForm().getTextField('Unicode sample')
    assert.equal(field.getText(), value)
    const ap = saved.context.lookup(field.acroField.getWidgets()[0].getNormalAppearance())
    const content = Buffer.from(decodePDFRawStream(ap).decode()).toString('latin1')
    if (label === 'glyphless') {
      assert.ok(content.includes(`/ActualText ${PDFHexString.fromText(value).toString()}`))
      assert.ok(content.includes('<> Tj')); assert.doesNotMatch(content, /<[0-9A-F]+> Tj/)
    } else {
      const scale = Number(content.match(/\/DejaVuSans-Oblique ([\d.]+) Tf/)[1]) / 2048
      const matrix = [...content.matchAll(/1 0 0 1 ([\d.-]+) ([\d.-]+) Tm/g)][1]
      // Independently pinned DejaVu italic j bounds, not production ink output.
      assert.ok(Number(matrix[1]) - 231 * scale >= 2 - 1e-7)
      assert.ok(Number(matrix[2]) - 426 * scale >= 2 - 1e-7)
    }
  }
  assert.deepEqual(readFileSync(sourcePath), Buffer.from(source), 'source bytes stay unchanged')
  assert.deepEqual(errors, [])
  assert.deepEqual(postRequests, [])
  console.log('PDF embedded font browser smoke: PASS (UI upload/apply/download, Type0 fixed font, exact clusters/ligatures, continuation outlines, contextual RTL/multiscript, name/CID-keyed CFF resources, selected TTC face/reset, variation selectors/Hangul, glyphless source semantics, italic automatic ink fit, missing-glyph refusal)')
} finally {
  socket?.close()
  for (const task of pending.values()) clearTimeout(task.timer)
  if (chrome) await terminateProcess(chrome.child)
  await server?.close()
}
