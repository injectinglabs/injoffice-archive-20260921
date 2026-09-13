import assert from 'node:assert/strict'
import { existsSync, mkdtempSync, readFileSync, readdirSync, writeFileSync } from 'node:fs'
import { createRequire } from 'node:module'
import { tmpdir } from 'node:os'
import { resolve } from 'node:path'
import { decodePDFRawStream, PDFDocument, PDFName } from 'pdf-lib'
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
field.addToPage(doc.addPage([400, 300]), { x: 20, y: 200, width: 350, height: 30 })
const source = await doc.save()
writeFileSync(sourcePath, source)

let server, chrome, socket, sequence = 0
const pending = new Map()
const errors = []
const postRequests = []
const panel = '[data-demo-surface="pdf"] #pdf-operation-panel'
const textInput = `${panel} input:not([type="file"])`
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
  throw new Error(`Timed out: ${label}\n${await evaluate(`document.querySelector(${JSON.stringify(panel)})?.textContent`)}`)
}
async function upload(selector, file) {
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
const apply = () => evaluate(`Array.from(document.querySelectorAll(${JSON.stringify(`${panel} button`)})).find(button => button.textContent.trim() === 'Apply form values').click()`)

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
  await until(`!!document.querySelector('[aria-label="TrueType appearance font"]')`, 'embedded font input')
  await upload('[aria-label="TrueType appearance font"]', fontPath)
  await until(`document.querySelector(${JSON.stringify(panel)}).textContent.includes('Selected: DejaVuSans.ttf')`, 'local font read')
  await setValue(textInput, 'AV café Ω Ж 😀')
  await apply()
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
  await apply()
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
  await setValue(textInput, 'مرحبا')
  await apply()
  await until(`document.querySelector(${JSON.stringify(panel)}).textContent.includes('No form values applied.')`, 'unsupported shaping is refused')
  assert.deepEqual(readFileSync(sourcePath), Buffer.from(source), 'source bytes stay unchanged')
  assert.deepEqual(errors, [])
  assert.deepEqual(postRequests, [])
  console.log('PDF embedded font browser smoke: PASS (UI upload/apply/download, Type0 fixed font, exact clusters/ligatures, continuation outlines, bidi refusal)')
} finally {
  socket?.close()
  for (const task of pending.values()) clearTimeout(task.timer)
  if (chrome) await terminateProcess(chrome.child)
  await server?.close()
}
