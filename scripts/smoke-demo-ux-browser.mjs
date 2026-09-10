// Run with Node 22+. DEMO_UX_URL can target an already-running integrated demo.
import { spawnSync } from 'node:child_process'
import { existsSync, mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { resolve, sep } from 'node:path'
import { PDFDocument, PDFName, PDFNumber, StandardFonts } from 'pdf-lib'
import { launchChromeForCDP, terminateProcess } from './chrome-cdp-startup.mjs'
import { startShowcaseServer } from './showcase-smoke-server.mjs'
import { startShowcaseDevServer } from './showcase-smoke-dev-server.mjs'

const root = resolve(import.meta.dirname, '..')
const profiles = []
const artifacts = process.env.SHOWCASE_OUTPUT ? resolve(process.env.SHOWCASE_OUTPUT) : mkdtempSync(resolve(tmpdir(), 'injoffice-ux-screenshots-'))
mkdirSync(artifacts, { recursive: true })
const errors = []
let chrome, server, cdp
const delay = ms => new Promise(resolvePromise => setTimeout(resolvePromise, ms))
const pdf = `document.querySelector('[data-demo-surface="pdf"]')`

try {
  let url = process.env.DEMO_UX_URL
  if (process.argv.includes('--built')) {
    server = await startShowcaseServer(resolve(root, 'apps/playground/dist'), '/injoffice-smoke/')
    url = server.url
  } else if (!url) {
    server = await startShowcaseDevServer(resolve(root, 'apps/playground'))
    url = server.url
  }
  chrome = await launchChromeForCDP({ executable: findChrome(), createProfile: () => {
    const profile = mkdtempSync(resolve(tmpdir(), 'injoffice-ux-chrome-'))
    profiles.push(profile)
    return profile
  } })
  cdp = await connectCDP(chrome.target.webSocketDebuggerUrl)
  cdp.on('Runtime.exceptionThrown', ({ exceptionDetails }) => errors.push(exceptionDetails.exception?.description ?? exceptionDetails.text))
  await cdp.send('Runtime.enable')
  await cdp.send('Page.enable')
  await cdp.send('Emulation.setDeviceMetricsOverride', { width: 1440, height: 1100, deviceScaleFactor: 1, mobile: false })
  await cdp.send('Page.navigate', { url: `${url.replace(/#.*$/, '')}#/pdf?feature=editor` })
  await ready()
  await evaluate(`${pdf}.scrollIntoView({ block: 'start' })`)
  await textLayerReady()
  await assert(`(() => {
    const span = ${pdf}.querySelector('.pdf-selectable-text span');
    const selection = window.getSelection(); selection.selectAllChildren(span);
    const copied = selection.toString(); selection.removeAllRanges();
    return copied.length > 0 && copied === span.textContent;
  })()`, 'native PDF text selection supports copying')
  await button('Rotate 90°')
  await ready()
  await textLayerReady()
  await assert(`${pdf}.querySelector('.pdf-selectable-text').dataset.mainRotation === '90'`, 'text layer follows page rotation')
  await button('Undo')
  await ready()
  await evaluate(`(() => { const select = ${pdf}.querySelector('[aria-label="Zoom"]'); select.value = '2'; select.dispatchEvent(new Event('change', { bubbles: true })); })()`)
  await ready()
  await textLayerReady()
  await screenshot('pdf-selectable-text-zoom.png')
  await evaluate(`(() => { const select = ${pdf}.querySelector('[aria-label="Zoom"]'); select.value = '1'; select.dispatchEvent(new Event('change', { bubbles: true })); })()`)
  await ready()

  // Tabs support a single tab stop, arrow keys, Home/End and linked panels.
  await evaluate(`document.querySelector('#pdf-tab-pages').focus()`)
  await key('ArrowRight')
  await assert(`document.activeElement.id === 'pdf-tab-mark' && document.querySelector('#pdf-operation-panel').getAttribute('aria-labelledby') === 'pdf-tab-mark'`, 'arrow-key tab activation')
  await key('End')
  await assert(`document.activeElement.id === 'pdf-tab-host'`, 'End selects last tab')
  await key('Home')
  await assert(`document.activeElement.id === 'pdf-tab-inspect'`, 'Home selects first tab')
  await assert(`${pdf}.querySelectorAll('[role="tab"][tabindex="0"]').length === 1 && ${pdf}.querySelector('[aria-label="Page number"]') !== null`, 'tab stops and page label')

  await tab('mark')
  await poll(() => evaluate(`Boolean(${pdf}.querySelector('.pdf-text-target'))`), 'selectable PDF passages')
  await clickSelector('[data-demo-surface="pdf"] .pdf-text-target')
  await assert(`Boolean(${pdf}.querySelector('.pdf-text-target[aria-pressed="true"]'))`, 'passage selection')
  await button('highlight')
  await ready()
  await assert(`${pdf}.dataset.demoDirty === 'true' && ${pdf}.querySelector('.pdf-result-list')?.textContent.includes('highlight')`, 'selected passage markup persisted')
  await screenshot('pdf-markup.png')
  await button('Undo')
  await ready()
  await assert(`${pdf}.dataset.demoDirty === 'false' && !${pdf}.querySelector('.pdf-result-list')?.textContent.includes('highlight')`, 'undo restores original PDF')
  await button('Redo')
  await ready()
  await assert(`${pdf}.dataset.demoDirty === 'true' && ${pdf}.querySelector('.pdf-result-list')?.textContent.includes('highlight')`, 'redo restores markup')

  await tab('draw')
  await button('Rectangle')
  await evaluate(`${pdf}.querySelector('.pdf-drawing-surface').scrollIntoView({ block: 'center' })`)
  await delay(150)
  // The page is taller than its scrolling viewport. Draw inside the visible
  // intersection, not at a fixed offset from the potentially clipped SVG top.
  const drawing = await evaluate(`(() => {
    const el = ${pdf}.querySelector('.pdf-drawing-surface');
    const r = el.getBoundingClientRect();
    const viewport = el.closest('.pdf-pages').getBoundingClientRect();
    const x = Math.max(r.left, viewport.left, 0) + 40;
    const y = Math.max(r.top, viewport.top, 80) + 40;
    if (x + 120 >= Math.min(r.right, viewport.right, innerWidth) || y + 60 >= Math.min(r.bottom, viewport.bottom, innerHeight)) throw new Error('PDF drawing viewport is too small');
    if (!el.contains(document.elementFromPoint(x, y)) || !el.contains(document.elementFromPoint(x + 120, y + 60))) throw new Error('PDF drawing coordinates are occluded');
    return { x, y };
  })()`)
  await cdp.send('Input.dispatchMouseEvent', { type: 'mousePressed', x: drawing.x, y: drawing.y, button: 'left', clickCount: 1 })
  await cdp.send('Input.dispatchMouseEvent', { type: 'mouseMoved', x: drawing.x + 120, y: drawing.y + 60, button: 'left', buttons: 1 })
  await assert(`Boolean(${pdf}.querySelector('.pdf-drawing-surface rect'))`, 'live drawing preview')
  await cdp.send('Input.dispatchMouseEvent', { type: 'mouseReleased', x: drawing.x + 120, y: drawing.y + 60, button: 'left', clickCount: 1 })
  await ready()
  await assert(`!${pdf}.querySelector('.pdf-drawing-surface')`, 'pointer drawing completes')
  await button('Line')
  await evaluate(`${pdf}.querySelector('.pdf-drawing-surface').focus()`)
  await key('Enter')
  await key('ArrowRight')
  await key('ArrowDown')
  await key('Enter')
  await ready()
  await assert(`!${pdf}.querySelector('.pdf-drawing-surface')`, 'keyboard drawing completes')
  await screenshot('pdf-drawings.png')

  await tab('inspect')
  await evaluate(`(() => { const input = ${pdf}.querySelector('[aria-label="Search PDF text"]'); Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(input, 'Launch'); input.dispatchEvent(new Event('input', { bubbles: true })); })()`)
  await button('Find')
  await poll(() => evaluate(`${pdf}.querySelectorAll('.pdf-search-highlight').length > 0`), 'visible search highlights')
  await screenshot('pdf-search.png')
  await assert(`Boolean(${pdf}.querySelector('.pdf-search-highlight').getBoundingClientRect().width)`, 'search highlight has visible bounds')

  // Verify real page deletion is recoverable, including the previous page count.
  await tab('pages')
  const pageCount = await evaluate(`Number(${pdf}.querySelector('[aria-label="Page number"]').max)`)
  await button('Delete page')
  await ready()
  await assert(`Number(${pdf}.querySelector('[aria-label="Page number"]').max) === ${pageCount - 1}`, 'page deletion')
  await button('Undo')
  await ready()
  await assert(`Number(${pdf}.querySelector('[aria-label="Page number"]').max) === ${pageCount}`, 'page deletion undo')
  cdp.on('Page.javascriptDialogOpening', () => { void cdp.send('Page.handleJavaScriptDialog', { accept: true }) })
  await evaluate(`(() => { const transfer = new DataTransfer(); transfer.items.add(new File(['not a PDF'], 'broken.pdf', { type: 'application/pdf' })); const input = ${pdf}.querySelector('[aria-label="Open a PDF file"]'); input.files = transfer.files; input.dispatchEvent(new Event('change', { bubbles: true })); })()`)
  await poll(() => evaluate(`${pdf}.innerText.includes('Could not read the selected PDF')`), 'invalid upload error')
  await assert(`Number(${pdf}.querySelector('[aria-label="Page number"]').max) === ${pageCount} && ${pdf}.dataset.demoDirty === 'true' && [...${pdf}.querySelectorAll('button')].some(b => b.textContent.trim() === 'Undo' && !b.disabled)`, 'invalid upload preserves working document and undo')
  // /UserUnit changes physical page geometry independently of viewer zoom.
  const unitDoc = await PDFDocument.create()
  const unitPage = unitDoc.addPage([200, 300])
  unitPage.node.set(PDFName.of('UserUnit'), PDFNumber.of(2))
  unitPage.drawText('Scaled PDF coordinates', { x: 20, y: 240, size: 12, font: await unitDoc.embedFont(StandardFonts.Helvetica) })
  const unitBytes = [...await unitDoc.save()]
  // Pass fixture bytes as a CDP value, never interpolate them into source code.
  const documentHandle = await cdp.send('Runtime.evaluate', { expression: 'document' })
  const objectId = documentHandle.result.objectId
  if (!objectId) throw new Error('Browser document handle unavailable')
  try {
    const uploaded = await cdp.send('Runtime.callFunctionOn', {
      objectId,
      functionDeclaration: `function(bytes) {
        const transfer = new DataTransfer();
        transfer.items.add(new File([new Uint8Array(bytes)], 'user-unit.pdf', { type: 'application/pdf' }));
        const input = this.querySelector('[data-demo-surface="pdf"] [aria-label="Open a PDF file"]');
        input.files = transfer.files;
        input.dispatchEvent(new Event('change', { bubbles: true }));
      }`,
      arguments: [{ value: unitBytes }],
      returnByValue: true,
    })
    if (uploaded.exceptionDetails) throw new Error(uploaded.exceptionDetails.text)
  } finally {
    await cdp.send('Runtime.releaseObject', { objectId })
  }
  await poll(() => evaluate(`${pdf}.querySelector('.native-status').textContent.includes('user-unit.pdf')`), 'UserUnit PDF opened')
  await ready()
  await textLayerReady()
  await assert(`${pdf}.querySelector('.pdf-selectable-text').style.getPropertyValue('--total-scale-factor') === '2'`, 'text layer includes PDF user unit')
  await screenshot('pdf-user-unit.png')
  await button('Continuous reading')
  await poll(() => evaluate(`${pdf}.querySelector('[data-read-page="1"] [data-render-state="ready"]') !== null`), 'continuous page ready')
  await assert(`${pdf}.querySelector('[data-read-page="1"] .pdf-selectable-text').style.getPropertyValue('--total-scale-factor') === '2'`, 'continuous text includes user unit')
  await screenshot('pdf-continuous-user-unit.png')
  // Exercise a longer, mixed-size PDF to verify navigation and bounded rasters.
  const longDoc = await PDFDocument.create()
  const longFont = await longDoc.embedFont(StandardFonts.Helvetica)
  for (let index = 1; index <= 24; index++) {
    const next = longDoc.addPage(index % 2 ? [612, 792] : [792, 612])
    next.drawText(`Reading page ${index}`, { x: 40, y: 500, size: 20, font: longFont })
  }
  const longBytes = [...await longDoc.save()]
  const longHandle = await cdp.send('Runtime.evaluate', { expression: 'document' })
  try {
    const uploaded = await cdp.send('Runtime.callFunctionOn', {
      objectId: longHandle.result.objectId,
      functionDeclaration: `function(bytes) {
        const transfer = new DataTransfer();
        transfer.items.add(new File([new Uint8Array(bytes)], 'continuous-reading.pdf', { type: 'application/pdf' }));
        const input = this.querySelector('[data-demo-surface="pdf"] [aria-label="Open a PDF file"]');
        input.files = transfer.files;
        input.dispatchEvent(new Event('change', { bubbles: true }));
      }`, arguments: [{ value: longBytes }], returnByValue: true,
    })
    if (uploaded.exceptionDetails) throw new Error(uploaded.exceptionDetails.text)
  } finally { await cdp.send('Runtime.releaseObject', { objectId: longHandle.result.objectId }) }
  await poll(() => evaluate(`${pdf}.querySelector('.native-status').textContent.includes('continuous-reading.pdf')`), 'long PDF opened')
  await ready()
  await evaluate(`${pdf}.querySelector('[aria-label="Go to page 12"]').click()`)
  await poll(() => evaluate(`${pdf}.querySelector('[data-read-page="12"] [data-render-state="ready"]') !== null`), 'thumbnail jumps to loaded page')
  await assert(`${pdf}.querySelector('.pdf-thumbnail-rail [aria-current="page"]').getAttribute('aria-label') === 'Go to page 12'`, 'thumbnail identifies current page')
  await assert(`${pdf}.querySelectorAll('.pdf-continuous-scroll canvas').length <= 3 && ${pdf}.querySelectorAll('.pdf-thumbnail-rail canvas').length <= 5`, 'bounded full-size and thumbnail canvases')
  await assert(`(() => { const canvas = ${pdf}.querySelector('[data-read-page="12"] canvas'); return canvas.width * canvas.height <= 4000000 && canvas.getBoundingClientRect().width > canvas.getBoundingClientRect().height; })()`, 'mixed page geometry and backing-store budget')
  await evaluate(`(() => { const root = ${pdf}.querySelector('.pdf-continuous-scroll'); const next = root.querySelector('[data-read-page="13"]'); root.scrollTop += next.getBoundingClientRect().top - root.getBoundingClientRect().top - 16; })()`)
  await poll(() => evaluate(`${pdf}.querySelector('[aria-label="Page number"]').value === '13'`), 'scroll updates page navigation')
  await poll(() => evaluate(`${pdf}.querySelector('[data-read-page="13"] [data-render-state="ready"]') !== null`), 'scrolled page ready')
  await assert(`(() => { const rail = ${pdf}.querySelector('.pdf-thumbnail-rail'); const item = rail.querySelector('[aria-current="page"]'); const r = rail.getBoundingClientRect(); const b = item.getBoundingClientRect(); return b.top >= r.top - 2 && b.bottom <= r.bottom + 2; })()`, 'selected thumbnail remains in rail viewport')
  await assert(`(() => { const host = ${pdf}.querySelector('.pdf-reader-host').getBoundingClientRect(); const r = ${pdf}.querySelector('.pdf-reader').getBoundingClientRect(); return r.top >= host.top && r.bottom <= host.bottom && r.height > 200; })()`, 'reader fits its clipping parent without nested outer scrolling')
  await thumbnailVisible()
  await screenshot('pdf-continuous-reading.png')
  await evaluate(`(() => { const select = ${pdf}.querySelector('[aria-label="Zoom"]'); select.value = '2'; select.dispatchEvent(new Event('change', { bubbles: true })); })()`)
  await evaluate(`${pdf}.querySelector('[aria-label="Go to page 20"]').click(); ${pdf}.querySelector('[aria-label="Go to page 4"]').click()`)
  await poll(() => evaluate(`${pdf}.querySelector('[data-read-page="4"] [data-render-state="ready"]') !== null`), 'rapid navigation and zoom settle')
  await assert(`${pdf}.querySelectorAll('.pdf-continuous-scroll canvas').length <= 3`, 'rapid navigation retains canvas bound')
  await cdp.send('Emulation.setDeviceMetricsOverride', { width: 390, height: 844, deviceScaleFactor: 1, mobile: true })
  await delay(150)
  await assert(`(() => { const r = ${pdf}.querySelector('.pdf-reader').getBoundingClientRect(); return r.width > 100 && r.width <= innerWidth && ${pdf}.querySelector('.pdf-continuous-scroll').clientWidth > 100; })()`, 'continuous reader remains usable on mobile')
  await assert(`(() => { const rail = ${pdf}.querySelector('.pdf-thumbnail-rail'); const r = rail.getBoundingClientRect(); const b = rail.querySelector('[aria-current="page"]').getBoundingClientRect(); return b.top >= r.top - 2 && b.bottom <= r.bottom + 2; })()`, 'selected thumbnail stays visible after mobile resize')
  await thumbnailVisible()
  await screenshot('pdf-continuous-mobile.png')
  await cdp.send('Emulation.setDeviceMetricsOverride', { width: 1440, height: 1100, deviceScaleFactor: 1, mobile: false })
  await tab('mark')
  await ready()
  await assert(`${pdf}.querySelector('.pdf-reader') === null && ${pdf}.querySelector('.pdf-page-stage canvas') !== null`, 'editing restores single-page interaction')
  if (errors.length) throw new Error(`Browser exceptions: ${errors.join('\n')}`)
  console.log(JSON.stringify({ result: 'PASS', checks: ['native text selection', 'text geometry at rotation and zoom', 'PDF user unit geometry', 'tab keyboard navigation', 'passage markup', 'undo/redo', 'pointer drawing', 'keyboard drawing', 'search highlighting', 'page deletion recovery', 'invalid upload preserves edits', 'continuous reading and thumbnails', 'bounded page canvas allocation', 'mixed page navigation', 'return to single-page editing'], screenshots: artifacts }, null, 2))
} catch (error) {
  console.error(`Screenshots: ${artifacts}`)
  if (cdp) {
    await screenshot('failure.png').catch(() => undefined)
    console.error(await evaluate(`document.querySelector('[data-demo-surface="pdf"]')?.innerText ?? document.body.innerText.slice(0, 3000)`).catch(() => 'No browser diagnostic'))
  }
  throw error
} finally {
  cdp?.close()
  try { if (chrome) await terminateProcess(chrome.child) }
  finally {
    try { await server?.close() }
    finally { for (const profile of profiles) rmSync(profile, { recursive: true, force: true, maxRetries: 10, retryDelay: 100 }) }
  }
}

async function evaluate(expression) {
  const response = await cdp.send('Runtime.evaluate', { expression, awaitPromise: true, returnByValue: true })
  if (response.exceptionDetails) throw new Error(response.exceptionDetails.exception?.description ?? response.exceptionDetails.text)
  return response.result.value
}
async function assert(expression, label) { if (!await evaluate(expression)) throw new Error(`Failed: ${label}`) }
async function poll(check, label, timeout = 30000) {
  const end = Date.now() + timeout
  while (Date.now() < end) { if (await check()) return; await delay(50) }
  throw new Error(`Timed out: ${label}`)
}
async function ready() { await delay(100); await poll(() => evaluate(`${pdf}?.querySelector('.native-status')?.dataset.state === 'ready'`), 'PDF ready') }
async function textLayerReady() {
  await poll(() => evaluate(`(() => {
    const layer = ${pdf}.querySelector('.pdf-selectable-text');
    return layer && !layer.dataset.disabled && layer.querySelector('span')?.getBoundingClientRect().width > 0;
  })()`), 'selectable PDF text ready')
  await assert(`(() => {
    const layer = ${pdf}.querySelector('.pdf-selectable-text');
    const canvas = ${pdf}.querySelector('canvas').getBoundingClientRect();
    const bounds = layer.getBoundingClientRect();
    return Math.abs(bounds.left - canvas.left) < 2 && Math.abs(bounds.top - canvas.top) < 2 &&
      Math.abs(bounds.width - canvas.width) < 2 && Math.abs(bounds.height - canvas.height) < 2 &&
      [...layer.querySelectorAll('span')].filter(span => span.textContent.trim()).every(span => {
        const r = span.getBoundingClientRect();
        return r.width > 0 && r.height > 0 && r.left >= canvas.left - 2 && r.right <= canvas.right + 2 && r.top >= canvas.top - 2 && r.bottom <= canvas.bottom + 2;
      });
  })()`, 'PDF text geometry stays aligned with canvas at rotation and zoom')
}
async function tab(id) { await evaluate(`document.querySelector('#pdf-tab-${id}').click()`); await delay(50) }
async function button(text) {
  await poll(() => evaluate(`(() => { const b = [...${pdf}.querySelectorAll('button')].find(b => b.textContent.trim() === ${JSON.stringify(text)}); if (!b || b.disabled) return false; b.click(); return true; })()`), `button ${text}`)
  await delay(50)
}
async function key(keyName) {
  await cdp.send('Input.dispatchKeyEvent', { type: 'keyDown', key: keyName, code: keyName, windowsVirtualKeyCode: { Enter: 13, Home: 36, End: 35, ArrowLeft: 37, ArrowUp: 38, ArrowRight: 39, ArrowDown: 40 }[keyName] })
  await cdp.send('Input.dispatchKeyEvent', { type: 'keyUp', key: keyName, code: keyName })
  await delay(50)
}
async function clickSelector(selector) {
  const box = await evaluate(`(() => { const e = document.querySelector(${JSON.stringify(selector)}); e.scrollIntoView({ block: 'center' }); const r = e.getBoundingClientRect(); return { x: r.x + r.width / 2, y: r.y + r.height / 2 }; })()`)
  await cdp.send('Input.dispatchMouseEvent', { type: 'mousePressed', ...box, button: 'left', clickCount: 1 })
  await cdp.send('Input.dispatchMouseEvent', { type: 'mouseReleased', ...box, button: 'left', clickCount: 1 })
  await delay(50)
}
async function screenshot(name) {
  await evaluate(`${pdf}?.scrollIntoView({ block: 'start' })`)
  await evaluate(`window.scrollBy(0, -70)`)
  const image = await cdp.send('Page.captureScreenshot', { format: 'png' })
  writeFileSync(resolve(artifacts, name), Buffer.from(image.data, 'base64'))
}
async function thumbnailVisible() {
  await evaluate(`${pdf}.querySelector('.pdf-reader').scrollIntoView({ block: 'center' })`)
  await delay(100)
  const visible = await evaluate(`(() => {
    const selected = ${pdf}.querySelector('.pdf-thumbnail-rail [aria-current="page"]');
    const item = selected.getBoundingClientRect();
    let top = 64, bottom = innerHeight;
    for (let parent = selected.parentElement; parent; parent = parent.parentElement) {
      if (/(auto|scroll|hidden|clip)/.test(getComputedStyle(parent).overflowY)) {
        const bounds = parent.getBoundingClientRect(); top = Math.max(top, bounds.top); bottom = Math.min(bottom, bounds.bottom);
      }
    }
    return { ok: item.top >= top - 2 && item.bottom <= bottom + 2, top, bottom, itemTop: item.top, itemBottom: item.bottom, reader: ${pdf}.querySelector('.pdf-reader').getBoundingClientRect().toJSON(), page: selected.textContent };
  })()`)
  if (!visible.ok) throw new Error(`Selected thumbnail is clipped: ${JSON.stringify(visible)}`)
}
function findChrome() {
  for (const candidate of [process.env.CHROME_PATH, process.env.CHROME_BIN, '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome', '/Applications/Chromium.app/Contents/MacOS/Chromium', 'google-chrome', 'google-chrome-stable', 'chromium', 'chromium-browser'].filter(Boolean)) {
    if (candidate.includes(sep) && !existsSync(candidate)) continue
    if (spawnSync(candidate, ['--version'], { stdio: 'ignore' }).status === 0) return candidate
  }
  throw new Error('Chrome/Chromium is required; set CHROME_PATH')
}
async function connectCDP(url) {
  const socket = new WebSocket(url)
  await new Promise((resolvePromise, reject) => { socket.addEventListener('open', resolvePromise, { once: true }); socket.addEventListener('error', reject, { once: true }) })
  let nextId = 0
  const pending = new Map(), listeners = new Map()
  socket.addEventListener('message', event => {
    const message = JSON.parse(event.data)
    if (message.id !== undefined) {
      const request = pending.get(message.id)
      if (!request) return
      pending.delete(message.id)
      clearTimeout(request.timer)
      if (message.error) request.reject(new Error(message.error.message)); else request.resolve(message.result)
    } else for (const listener of listeners.get(message.method) ?? []) listener(message.params)
  })
  return {
    send(method, params = {}) { return new Promise((resolvePromise, reject) => { const id = ++nextId; const timer = setTimeout(() => { pending.delete(id); reject(new Error(`CDP timeout: ${method}`)) }, 30000); pending.set(id, { resolve: resolvePromise, reject, timer }); socket.send(JSON.stringify({ id, method, params })) }) },
    on(method, listener) { listeners.set(method, [...(listeners.get(method) ?? []), listener]) },
    close() { socket.close() },
  }
}
