// Run with Node 22+. DEMO_UX_URL can target an already-running integrated demo.
import { spawnSync } from 'node:child_process'
import { existsSync, mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { resolve, sep } from 'node:path'
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
  await cdp.send('Page.navigate', { url: `${url.replace(/#.*$/, '')}#/pdf` })
  await ready()
  await evaluate(`${pdf}.scrollIntoView({ block: 'start' })`)

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
  if (errors.length) throw new Error(`Browser exceptions: ${errors.join('\n')}`)
  console.log(JSON.stringify({ result: 'PASS', checks: ['tab keyboard navigation', 'passage markup', 'undo/redo', 'pointer drawing', 'keyboard drawing', 'search highlighting', 'page deletion recovery', 'invalid upload preserves edits'], screenshots: artifacts }, null, 2))
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
