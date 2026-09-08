import assert from 'node:assert/strict'
import { existsSync, mkdirSync, mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { resolve } from 'node:path'
import { launchChromeForCDP, terminateProcess } from './chrome-cdp-startup.mjs'
import { startShowcaseServer } from './showcase-smoke-server.mjs'

const output = process.env.SHOWCASE_OUTPUT ? resolve(process.env.SHOWCASE_OUTPUT) : mkdtempSync(resolve(tmpdir(), 'injoffice-showcase-review-'))
mkdirSync(output, { recursive: true })
let chrome, socket, staticServer
let sequence = 0
const pending = new Map()
const errors = []
const heldRequests = []
function onMessage({ data }) {
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
  } else if (message.method === 'Fetch.requestPaused') {
    heldRequests.push(message.params.requestId)
  }
}
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
    if (await evaluate(`Boolean(${expression})`)) return
    await new Promise((resolve) => setTimeout(resolve, 100))
  } while (Date.now() - start < timeout)
  throw new Error(`Timed out: ${label}`)
}
const click = (selector) => evaluate(`document.querySelector(${JSON.stringify(selector)}).click()`)
// Public chart/workbook APIs: an installed float is not proof that it has data.
const sheetChartState = `(() => {
  const host = window.__injoffice
  const spec = host?.charts?.list()[0]
  const sheet = spec && host.univerAPI.getActiveWorkbook()?.getSheetBySheetId(spec.range.sheetId)
  if (!sheet) return null
  const ref = spec.range
  const values = sheet.getRange(ref.startRow, ref.startColumn, ref.endRow - ref.startRow + 1, ref.endColumn - ref.startColumn + 1).getRawValues()
  return { series: host.charts.getSeriesNames(spec.id), numericCells: values.flat().filter(value => typeof value === 'number' && Number.isFinite(value)), firstValue: sheet.getRange(5, 1).getRawValues()[0][0] }
})()`
async function screenshot(name) {
  await new Promise((resolve) => setTimeout(resolve, 250))
  const { data } = await send('Page.captureScreenshot', { format: 'png' })
  writeFileSync(resolve(output, `${name}.png`), Buffer.from(data, 'base64'))
}
async function route(path) {
  await evaluate(`location.hash = '#/${path}'`)
  const surface = path.split('?')[0]
  await until(`document.querySelector('.app-shell')?.dataset.surface === ${JSON.stringify(surface)} && !document.querySelector('.demo-loading') && document.querySelector('.app-main')?.getAttribute('aria-busy') !== 'true'`, path)
}
try {
  if (process.argv.includes('--built')) staticServer = await startShowcaseServer(resolve(import.meta.dirname, '../apps/playground/dist'))
  const origin = new URL(staticServer?.url ?? process.env.SHOWCASE_URL ?? 'http://127.0.0.1:3100/')
  origin.hash = '#/overview'
  chrome = await launchChromeForCDP({
    executable: process.env.CHROME_BIN ?? ['/Applications/Google Chrome.app/Contents/MacOS/Google Chrome', '/usr/bin/google-chrome', '/usr/bin/chromium', '/usr/bin/chromium-browser'].find(existsSync) ?? 'google-chrome',
    createProfile: () => mkdtempSync(resolve(tmpdir(), 'injoffice-showcase-chrome-')),
  })
  socket = new WebSocket(chrome.target.webSocketDebuggerUrl)
  await new Promise((resolve, reject) => { socket.addEventListener('open', resolve, { once: true }); socket.addEventListener('error', reject, { once: true }) })
  socket.addEventListener('message', onMessage)
  await send('Runtime.enable')
  await send('Page.enable')
  await send('Emulation.setDeviceMetricsOverride', { width: 1440, height: 1000, deviceScaleFactor: 1, mobile: false })
  await send('Page.navigate', { url: origin.href })
  await until(`document.querySelectorAll('.showcase-item').length === 19`, 'catalogue ready')
  await evaluate(`document.querySelector('.scheme-toggle button:first-child').click()`)
  await until(`document.documentElement.dataset.theme === 'light'`, 'light theme')
  await screenshot('catalogue-desktop')
  await evaluate(`(() => { const input = document.querySelector('#showcase-query'); Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(input, 'redact'); input.dispatchEvent(new Event('input', { bubbles: true })); })()`)
  await until(`document.querySelectorAll('.showcase-item').length === 1 && document.querySelector('.showcase-item').getAttribute('href') === '#/pdf'`, 'search narrows to PDF')
  await evaluate(`Array.from(document.querySelectorAll('.showcase-filter-group button')).find(button => button.textContent === 'Edit files').click()`)
  await evaluate(`Array.from(document.querySelectorAll('.showcase-filter-group--formats button')).find(button => button.textContent === 'PDF').click()`)
  await click('.showcase-item')
  await until(`document.querySelector('.app-shell')?.dataset.surface === 'pdf' && !document.querySelector('.demo-loading')`, 'filtered result opens')
  await click('.demo-back')
  await until(`document.querySelector('#showcase-query')?.value === 'redact' && document.querySelectorAll('.showcase-item').length === 1`, 'Back preserves search')
  assert.deepEqual(await evaluate(`Array.from(document.querySelectorAll('.showcase-filter-group button[aria-pressed=true]')).map(button => button.textContent)`), ['Edit files', 'PDF'], 'Back preserves task and file-type filters')
  await evaluate(`(() => { const input = document.querySelector('#showcase-query'); Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(input, 'zzznomatch'); input.dispatchEvent(new Event('input', { bubbles: true })); })()`)
  await until(`!!document.querySelector('.showcase-empty')`, 'empty search')
  await click('.showcase-empty button')
  await until(`document.querySelectorAll('.showcase-item').length === 19`, 'clear filters')
  assert.equal(await evaluate(`new Set([...document.querySelectorAll('.showcase-item')].map(item => item.getAttribute('href'))).size`), 19, 'catalogue examples have distinct routes')
  await route('charts')
  assert.equal(await evaluate(`document.querySelector('.app-shell').dataset.navigation`), 'text')
  assert.equal(await evaluate(`document.querySelectorAll('.app-sidebar a[href^="#/agent?format="]').length`), 4, 'all AI formats appear in navigation')

  // Hold the first cold route's scripts until continuity has been observed. This
  // catches the loading blink deterministically rather than racing local disk.
  await send('Fetch.enable', { patterns: [{ urlPattern: '*', resourceType: 'Script', requestStage: 'Request' }] })
  await evaluate(`window.__showcasePreviousStage = document.querySelector('.demo-stage'); location.hash = '#/pivots'`)
  const coldStart = Date.now()
  while (heldRequests.length === 0 && Date.now() - coldStart < 10_000) await new Promise((resolve) => setTimeout(resolve, 50))
  assert.ok(heldRequests.length > 0, 'cold navigation requests a lazy script')
  await until(`document.querySelector('.app-main')?.getAttribute('aria-busy') === 'true'`, 'cold navigation announced busy')
  assert.equal(await evaluate(`window.__showcasePreviousStage.isConnected && !document.querySelector('.demo-loading') && document.querySelector('.app-shell').dataset.surface === 'charts'`), true, 'previous editor remains visible during cold navigation')
  await send('Fetch.disable')
  await until(`document.querySelector('.app-shell')?.dataset.surface === 'pivots' && !document.querySelector('.demo-loading')`, 'cold route completes')
  await route('charts')
  await click('.source-proof-trigger')
  await until(`document.activeElement?.classList.contains('source-proof-close')`, 'modal initial focus')
  assert.equal(await evaluate(`!!document.querySelector('.demo-stage').closest('[inert]')`), true, 'preview background is inert')
  await send('Input.dispatchKeyEvent', { type: 'keyDown', key: 'Tab', code: 'Tab', modifiers: 8 })
  assert.equal(await evaluate(`document.querySelector('[role="dialog"]').contains(document.activeElement)`), true, 'Shift+Tab stays inside the guide')
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
  await until(`(${sheetChartState})?.series.length > 0 && (${sheetChartState})?.numericCells.length === 5`, 'seeded chart has a numeric series')
  const originalChart = await evaluate(sheetChartState)
  await evaluate(`(() => { const host = window.__injoffice; const spec = host.charts.list()[0]; host.univerAPI.getActiveWorkbook().getSheetBySheetId(spec.range.sheetId).getRange(5, 1).setValue(999); })()`)
  await until(`(${sheetChartState})?.firstValue === 999`, 'overview chart source is editable')
  await click('.demo-reset-trigger')
  await until(`document.querySelector('[data-demo-surface="sheets"] canvas') && window.__injoffice?.charts?.list().length > 0`, 'reset restores workbook')
  await until(`(${sheetChartState})?.firstValue === ${JSON.stringify(originalChart.firstValue)} && (${sheetChartState})?.series.length > 0`, 'reset restores seeded chart data')
  assert.deepEqual(await evaluate(sheetChartState), originalChart, 'reset restores the complete numeric chart source')
  for (let i = 0; i < 2; i++) { await route('charts'); await route('sheets'); await until(`document.querySelector('[data-demo-surface="sheets"] canvas')`, 'remounted sheet') }
  await evaluate(`location.hash = '#/sheets?view=native'`)
  await until(`!!document.querySelector('.native-toolbar')`, 'same-surface deep link switches mode')
  await route('collab')
  await until(`document.querySelectorAll('canvas').length > 0`, 'collaboration rendered')
  for (const tool of ['sheets', 'docs', 'slides', 'pdf']) {
    await route(`agent?format=${tool}`)
    await until(`document.querySelector('[data-agent-tool=${tool}]') && document.querySelector('.agent-demo__status')?.dataset.state === 'ready' && Array.from(document.querySelectorAll('button')).some(button => button.textContent.trim() === 'Run agent' && !button.disabled)`, `${tool} AI deep link`, 90_000)
    const boundary = await evaluate(`document.querySelector('[data-agent-boundary]')?.textContent`)
    if (tool === 'sheets') assert.match(boundary, /Real XLSX file.*native browser write and exact-byte reopen/, 'native workflow describes real byte proof')
    else assert.match(boundary, /Lifecycle simulation.*not Office file bytes/, 'simulated formats disclose their file boundary')
    await evaluate(`Array.from(document.querySelectorAll('button')).find(button => button.textContent.trim() === 'Run agent').click()`)
    await until(`document.querySelector('.agent-demo__status')?.dataset.state === 'awaiting-approval'`, `${tool} preview and validation`, 90_000)
    assert.equal(await evaluate(`Array.from(document.querySelectorAll('button')).find(button => button.textContent.trim() === 'Commit approved change').disabled`), true, 'commit requires explicit approval')
    await click('.agent-approval input[type=checkbox]')
    await evaluate(`Array.from(document.querySelectorAll('button')).find(button => button.textContent.trim() === 'Commit approved change').click()`)
    await until(`document.querySelector('.agent-demo__status')?.dataset.state === 'verified'`, `${tool} approved commit verifies`, 90_000)
    if (tool === 'sheets') {
      const download = await evaluate(`(async () => {
        const link = document.querySelector('[data-agent-download]')
        if (!link) return null
        const bytes = new Uint8Array(await (await fetch(link.href)).arrayBuffer())
        return { name: link.download, length: bytes.length, signature: Array.from(bytes.slice(0, 4)) }
      })()`)
      assert.ok(download?.name.endsWith('.xlsx'), 'verified AI workflow offers an XLSX download')
      assert.ok(download.length > 1000, 'download contains a real workbook archive')
      assert.deepEqual(download.signature, [80, 75, 3, 4], 'download is ZIP bytes, not JSON-shaped simulation')
      await screenshot('agent-real-xlsx-verified')
    }
    await evaluate(`Array.from(document.querySelectorAll('button')).find(button => button.textContent.trim() === 'Refusal proof').click()`)
    await until(`document.querySelector('.agent-demo__status')?.dataset.state === 'ready' && Array.from(document.querySelectorAll('button')).some(button => button.textContent.trim() === 'Run agent' && !button.disabled)`, `${tool} refusal ready`, 90_000)
    await evaluate(`Array.from(document.querySelectorAll('button')).find(button => button.textContent.trim() === 'Run agent').click()`)
    await until(`document.querySelector('.agent-demo__status')?.dataset.state === 'refused'`, `${tool} unsupported operation refused`, 90_000)
    assert.equal(await evaluate(`document.querySelectorAll('.agent-tool-log li[data-state=done]').length === 6 && !document.querySelector('.agent-approval input')`), true, 'refusal stops before commit and verification')
    assert.equal(await evaluate(`document.querySelector('[data-agent-download]') === null`), true, 'refusal does not expose a previous output')
    await click('.demo-reset-trigger')
    await until(`document.querySelector('[data-agent-tool=${tool}]') && document.querySelector('.agent-demo__status')?.dataset.state === 'ready' && Array.from(document.querySelectorAll('button')).some(button => button.textContent.trim() === 'Run agent' && !button.disabled)`, `${tool} reset reloads the source`, 90_000)
    assert.equal(await evaluate(`!document.querySelector('[data-agent-download]') && !document.querySelector('.agent-diff') && document.querySelectorAll('.agent-tool-log li[data-state=done]').length === 0 && document.querySelector('.agent-approval input').checked === false`), true, 'reset clears outputs, proof, and approval')
  }
  await click('.source-proof-trigger')
  await click('.guided-recipe__complete')
  await until(`document.querySelector('[role=progressbar]').getAttribute('aria-valuenow') === '1'`, 'AI guide progress recorded')
  await route('agent?format=docs')
  await until(`document.querySelector('.source-proof-layer').hidden && document.querySelector('[data-agent-tool=docs]')`, 'same-surface format switch closes stale guide')
  await click('.source-proof-trigger')
  assert.equal(await evaluate(`document.querySelector('[role=progressbar]').getAttribute('aria-valuenow')`), '0', 'new AI format starts its own guide progress')
  await click('.source-proof-close')
  await route('overview')
  await evaluate(`document.querySelector('.app-main').scrollTop = 0`)
  for (const width of [1200, 1024, 768]) {
    await send('Emulation.setDeviceMetricsOverride', { width, height: 1000, deviceScaleFactor: 1, mobile: false })
    await until(`document.querySelector('.app-main').scrollWidth <= document.querySelector('.app-main').clientWidth`, `catalogue fits width ${width}`, 1000)
  }
  await send('Emulation.setDeviceMetricsOverride', { width: 390, height: 844, deviceScaleFactor: 1, mobile: true })
  await until(`document.documentElement.scrollWidth <= 390`, 'mobile catalogue has no horizontal overflow')
  assert.ok(await evaluate(`document.querySelector('.app-sidebar').getBoundingClientRect().height < 160`), 'mobile navigation does not retain a desktop-height blank area')
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
  console.log(JSON.stringify({ status: 'passed', screenshots: output, checks: ['19 distinct examples', 'search', 'filter return persistence', 'empty recovery', 'text navigation', 'cold-route continuity', 'modal focus/inert', 'checklist persistence', 'source loading', 'all 16 surfaces', 'four AI approvals and refusals', 'AI proof boundaries and resets', 'AI format-specific guides', 'numeric chart source and reset', 'same-surface deep link', 'mobile layout', 'dark theme'], errors }, null, 2))
} catch (error) {
  if (socket?.readyState === WebSocket.OPEN) {
    await screenshot('failure')
    console.error(JSON.stringify({ screenshots: output, errors, state: await evaluate(`({ hash: location.hash, canvasCount: document.querySelectorAll('canvas').length, chartCount: window.__injoffice?.charts?.list().length, text: document.querySelector('.app-main')?.innerText.slice(0, 3500) })`) }, null, 2))
  }
  throw error
} finally {
  for (const task of pending.values()) clearTimeout(task.timer)
  socket?.close()
  if (chrome) await terminateProcess(chrome.child)
  await staticServer?.close()
}
