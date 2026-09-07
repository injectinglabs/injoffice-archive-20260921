import assert from 'node:assert/strict'
import { existsSync, mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { resolve } from 'node:path'
import { launchChromeForCDP, terminateProcess } from './chrome-cdp-startup.mjs'

const origin = process.env.SHOWCASE_URL ?? 'http://127.0.0.1:3100/'
const output = mkdtempSync(resolve(tmpdir(), 'injoffice-showcase-review-'))
const chrome = await launchChromeForCDP({
  executable: process.env.CHROME_BIN ?? ['/Applications/Google Chrome.app/Contents/MacOS/Google Chrome', '/usr/bin/google-chrome', '/usr/bin/chromium', '/usr/bin/chromium-browser'].find(existsSync) ?? 'google-chrome',
  createProfile: () => mkdtempSync(resolve(tmpdir(), 'injoffice-showcase-chrome-')),
})
const socket = new WebSocket(chrome.target.webSocketDebuggerUrl)
await new Promise((resolve, reject) => { socket.addEventListener('open', resolve, { once: true }); socket.addEventListener('error', reject, { once: true }) })
let sequence = 0
const pending = new Map()
const errors = []
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
  } else if (message.method === 'Runtime.consoleAPICalled' && message.params.type === 'error') {
    errors.push(message.params.args.map((arg) => arg.value ?? arg.description).join(' '))
  }
})
function send(method, params = {}) {
  return new Promise((resolve, reject) => {
    const id = ++sequence
    const timer = setTimeout(() => { pending.delete(id); reject(new Error(`CDP timed out: ${method}`)) }, 30_000)
    pending.set(id, { resolve, reject, timer })
    socket.send(JSON.stringify({ id, method, params }))
  })
}
async function evaluate(expression) {
  const result = await send('Runtime.evaluate', { expression, returnByValue: true, awaitPromise: true })
  if (result.exceptionDetails) throw new Error(result.exceptionDetails.exception?.description ?? result.exceptionDetails.text)
  return result.result.value
}
async function until(expression, label, timeout = 30_000) {
  const start = Date.now()
  do {
    if (await evaluate(expression)) return
    await new Promise((resolve) => setTimeout(resolve, 100))
  } while (Date.now() - start < timeout)
  throw new Error(`Timed out: ${label}`)
}
const click = (selector) => evaluate(`document.querySelector(${JSON.stringify(selector)}).click()`)
async function screenshot(name) {
  await new Promise((resolve) => setTimeout(resolve, 250))
  const { data } = await send('Page.captureScreenshot', { format: 'png' })
  writeFileSync(resolve(output, `${name}.png`), Buffer.from(data, 'base64'))
}
async function route(surface) {
  await evaluate(`location.hash = '#/${surface}'`)
  await until(`document.querySelector('.app-shell')?.dataset.surface === ${JSON.stringify(surface)} && !document.querySelector('.demo-loading')`, surface)
}
try {
  await send('Runtime.enable')
  await send('Page.enable')
  await send('Emulation.setDeviceMetricsOverride', { width: 1440, height: 1000, deviceScaleFactor: 1, mobile: false })
  await send('Page.navigate', { url: `${origin}#/overview` })
  await until(`document.querySelectorAll('.showcase-item').length === 15`, 'catalogue ready')
  await evaluate(`document.querySelector('.scheme-toggle button:first-child').click()`)
  await until(`document.documentElement.dataset.theme === 'light'`, 'light theme')
  await screenshot('catalogue-desktop')
  await evaluate(`(() => { const input = document.querySelector('#showcase-query'); Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(input, 'redact'); input.dispatchEvent(new Event('input', { bubbles: true })); })()`)
  await until(`document.querySelectorAll('.showcase-item').length === 1 && document.querySelector('.showcase-item').getAttribute('href') === '#/pdf'`, 'search narrows to PDF')
  await evaluate(`(() => { const input = document.querySelector('#showcase-query'); Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(input, 'zzznomatch'); input.dispatchEvent(new Event('input', { bubbles: true })); })()`)
  await until(`!!document.querySelector('.showcase-empty')`, 'empty search')
  await click('.showcase-empty button')
  await until(`document.querySelectorAll('.showcase-item').length === 15`, 'clear filters')
  await route('charts')
  assert.equal(await evaluate(`document.querySelector('.app-shell').dataset.navigation`), 'collapsed')
  await click('.navigation-controls button')
  await until(`document.querySelector('.app-shell').dataset.navigation === 'expanded'`, 'expanded rail')
  await click('.navigation-controls button')
  await click('.source-proof-trigger')
  await until(`document.activeElement?.classList.contains('source-proof-close')`, 'modal initial focus')
  assert.equal(await evaluate(`document.querySelector('.demo-stage').inert`), true)
  await click('.guided-recipe__complete')
  await until(`document.querySelector('[role=progressbar]').getAttribute('aria-valuenow') === '1'`, 'checklist step')
  await click('.demo-source summary')
  await until(`document.querySelector('.demo-source pre')?.textContent.includes('export default')`, 'actual source loaded')
  await screenshot('source-and-guide')
  await send('Input.dispatchKeyEvent', { type: 'keyDown', key: 'Escape', code: 'Escape' })
  await until(`document.querySelector('.source-proof-layer').hidden && document.activeElement?.classList.contains('source-proof-trigger')`, 'Escape restores focus')
  await click('.source-proof-trigger')
  assert.equal(await evaluate(`document.querySelector('[role=progressbar]').getAttribute('aria-valuenow')`), '1')
  await click('.source-proof-close')
  await screenshot('focused-chart')
  for (const surface of ['pivots', 'shapes', 'connectors', 'formulas', 'docs', 'slides', 'pdf', 'history', 'font-metrics', 'pptx-authored', 'pptx-native', 'pptx-render', 'sheets']) await route(surface)
  await until(`document.querySelector('[data-demo-surface="sheets"] canvas') && window.__injoffice?.charts?.list().length > 0`, 'seeded sheet chart')
  await click('.demo-reset-trigger')
  await until(`document.querySelector('[data-demo-surface="sheets"] canvas') && window.__injoffice?.charts?.list().length > 0`, 'reset restores workbook')
  for (let i = 0; i < 2; i++) { await route('charts'); await route('sheets'); await until(`document.querySelector('[data-demo-surface="sheets"] canvas')`, 'remounted sheet') }
  await evaluate(`location.hash = '#/sheets?view=native'`)
  await until(`!!document.querySelector('.native-toolbar')`, 'same-surface deep link switches mode')
  await route('collab')
  await until(`document.querySelectorAll('canvas').length > 0`, 'collaboration rendered')
  await route('overview')
  await evaluate(`document.querySelector('.app-main').scrollTop = 0`)
  for (const width of [1200, 1024, 768]) {
    await send('Emulation.setDeviceMetricsOverride', { width, height: 1000, deviceScaleFactor: 1, mobile: false })
    await until(`document.querySelector('.app-main').scrollWidth <= document.querySelector('.app-main').clientWidth`, `catalogue fits width ${width}`, 1000)
  }
  await send('Emulation.setDeviceMetricsOverride', { width: 390, height: 844, deviceScaleFactor: 1, mobile: true })
  await until(`document.documentElement.scrollWidth <= 390`, 'mobile catalogue has no horizontal overflow')
  await screenshot('catalogue-mobile')
  await route('charts')
  await screenshot('focused-chart-mobile')
  await click('.source-proof-trigger')
  assert.equal(await evaluate(`(() => { const panel = document.querySelector('.guided-recipe'); const footer = panel.querySelector('footer'); return footer.getBoundingClientRect().bottom <= panel.getBoundingClientRect().bottom + 1 })()`), true, 'mobile guide footer is not clipped')
  await screenshot('guide-mobile')
  await click('.source-proof-close')
  await evaluate(`document.querySelector('.scheme-toggle button:last-child').click()`)
  await until(`document.documentElement.dataset.theme === 'dark'`, 'dark theme')
  await screenshot('focused-chart-dark')
  assert.deepEqual(errors, [], 'uncaught or console errors')
  console.log(JSON.stringify({ status: 'passed', screenshots: output, checks: ['search', 'empty recovery', 'navigation', 'modal focus/inert', 'checklist persistence', 'source loading', 'all 15 routes', 'sheet reset', 'same-surface deep link', 'mobile', 'dark theme'], errors }, null, 2))
} catch (error) {
  await screenshot('failure')
  console.error(JSON.stringify({ screenshots: output, errors, state: await evaluate(`({ hash: location.hash, canvasCount: document.querySelectorAll('canvas').length, chartCount: window.__injoffice?.charts?.list().length, text: document.querySelector('.app-main')?.innerText.slice(0, 3500) })`) }, null, 2))
  throw error
} finally {
  for (const task of pending.values()) clearTimeout(task.timer)
  socket.close()
  await terminateProcess(chrome.child)
}
