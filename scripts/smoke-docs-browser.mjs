import assert from 'node:assert/strict'
import { existsSync, mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { resolve } from 'node:path'
import { launchChromeForCDP, terminateProcess } from './chrome-cdp-startup.mjs'
import { startShowcaseServer } from './showcase-smoke-server.mjs'

const output = mkdtempSync(resolve(tmpdir(), 'injoffice-docs-smoke-'))
const server = process.env.DOCS_URL ? undefined : await startShowcaseServer(resolve(import.meta.dirname, '../apps/docs/.vitepress/dist'), process.env.DOCS_BASE || '/')
const url = process.env.DOCS_URL || server.url
let chrome, socket, sequence = 0
const pending = new Map()
const errors = []
const requests = []
const delay = ms => new Promise(resolve => setTimeout(resolve, ms))
function send(method, params = {}) {
  return new Promise((resolve, reject) => {
    const id = ++sequence
    const timer = setTimeout(() => { pending.delete(id); reject(new Error(`Timeout: ${method}`)) }, 30000)
    pending.set(id, { resolve, reject, timer })
    socket.send(JSON.stringify({ id, method, params }))
  })
}
async function evaluate(expression) {
  const result = await send('Runtime.evaluate', { expression, returnByValue: true, awaitPromise: true })
  if (result.exceptionDetails) throw new Error(result.exceptionDetails.text)
  return result.result.value
}
async function until(expression, label) {
  const deadline = Date.now() + 30000
  while (Date.now() < deadline) { if (await evaluate(expression)) return; await delay(100) }
  throw new Error(`Timed out: ${label}`)
}
async function screenshot(name) {
  await delay(300)
  const { data } = await send('Page.captureScreenshot', { format: 'png' })
  writeFileSync(resolve(output, `${name}.png`), Buffer.from(data, 'base64'))
}
async function open(path, heading) {
  await send('Page.navigate', { url: new URL(path, url).href })
  await until(`document.querySelector('h1')?.textContent.includes(${JSON.stringify(heading)})`, heading)
  await delay(300)
}
try {
  chrome = await launchChromeForCDP({
    executable: process.env.CHROME_BIN ?? ['/Applications/Google Chrome.app/Contents/MacOS/Google Chrome', '/usr/bin/google-chrome', '/usr/bin/chromium'].find(existsSync) ?? 'google-chrome',
    createProfile: () => mkdtempSync(resolve(tmpdir(), 'injoffice-docs-chrome-')),
  })
  socket = new WebSocket(chrome.target.webSocketDebuggerUrl)
  await new Promise((resolve, reject) => { socket.addEventListener('open', resolve, { once: true }); socket.addEventListener('error', reject, { once: true }) })
  socket.addEventListener('message', ({ data }) => {
    const message = JSON.parse(data)
    if (message.id) {
      const task = pending.get(message.id)
      if (!task) return
      pending.delete(message.id); clearTimeout(task.timer)
      if (message.error) task.reject(new Error(message.error.message))
      else task.resolve(message.result)
    } else if (message.method === 'Runtime.exceptionThrown') {
      errors.push(message.params.exceptionDetails.exception?.description ?? message.params.exceptionDetails.text)
    } else if (message.method === 'Network.requestWillBeSent') requests.push(message.params.request.url)
  })
  await send('Runtime.enable')
  await send('Page.enable')
  await send('Network.enable')
  await send('Emulation.setDeviceMetricsOverride', { width: 1440, height: 1000, deviceScaleFactor: 1, mobile: false })
  await send('Emulation.setEmulatedMedia', { features: [{ name: 'prefers-color-scheme', value: 'light' }] })
  // Clipboard is tested without reading or modifying the user's clipboard.
  await send('Page.addScriptToEvaluateOnNewDocument', { source: `Object.defineProperty(navigator, 'clipboard', { value: { writeText: async text => { window.__copiedCode = text } }, configurable: true })` })
  await open('./', 'Build document workflows')
  assert.match(await evaluate('document.title'), /InjOffice/)
  await screenshot('desktop-introduction')
  await evaluate(`document.querySelector('.VPSidebar a[href$="/guides/pdf.html"]').click()`)
  await until(`document.querySelector('h1')?.textContent.includes('PDFs')`, 'sidebar article navigation')
  await evaluate(`document.querySelector('.vp-doc button.copy').click()`)
  await until(`window.__copiedCode?.includes('rotateLastPage')`, 'copy actual sample')
  await screenshot('desktop-pdf')
  await evaluate(`document.querySelector('.DocSearch-Button').click()`)
  await until(`document.querySelector('#localsearch-input') === document.activeElement`, 'search keyboard focus')
  await evaluate(`{ const input = document.querySelector('#localsearch-input'); input.value = 'idempotency'; input.dispatchEvent(new Event('input', { bubbles: true })); }`)
  await until(`document.querySelectorAll('#localsearch-list a').length > 0`, 'local full-text results')
  await evaluate(`document.querySelector('#localsearch-list a').click()`)
  await until(`!document.querySelector('#localsearch-input')`, 'search selection navigates and closes')
  await open('agents/quickstart.html', 'Your first agent workflow')
  assert.ok(await evaluate(`document.querySelector('.vp-doc').textContent.includes('commitFromReview')`))
  await evaluate(`document.querySelector('.VPSwitchAppearance').click()`)
  await until(`document.documentElement.classList.contains('dark')`, 'dark theme')
  await screenshot('desktop-agent-dark')
  await send('Emulation.setDeviceMetricsOverride', { width: 390, height: 844, deviceScaleFactor: 1, mobile: true })
  await evaluate(`window.scrollTo(0, 0)`)
  await delay(200)
  assert.ok(await evaluate('document.documentElement.scrollWidth <= innerWidth + 1'), 'mobile page has no horizontal overflow')
  await evaluate(`document.querySelector('.VPLocalNav button.menu').click()`)
  await until(`document.querySelector('.VPSidebar')?.classList.contains('open')`, 'mobile sidebar opens')
  await screenshot('mobile-navigation')
  await evaluate(`document.querySelector('.VPSidebar a[href$="/getting-started/quickstart.html"]').click()`)
  await until(`document.querySelector('h1')?.textContent.includes('Quickstart') && !document.querySelector('.VPSidebar')?.classList.contains('open')`, 'mobile navigation closes after selection')
  await screenshot('mobile-quickstart')
  assert.ok(!requests.some(request => /\.(wasm|xlsx|docx|pptx|pdf)(\?|$)|\/api\/agent\//.test(request)), 'no demo/engine/model requests')
  assert.deepEqual(errors, [])
  console.log(JSON.stringify({ status: 'passed', screenshots: output, checks: ['standalone HTML', 'sidebar navigation', 'copied source example', 'local search and focus', 'direct article deep link', 'dark theme', 'mobile menu and overflow', 'no document engines or model calls'], errors }, null, 2))
} catch (error) {
  if (socket?.readyState === WebSocket.OPEN) {
    await screenshot('failure')
    console.error(JSON.stringify({ screenshots: output, errors, state: await evaluate(`({ url: location.href, heading: document.querySelector('h1')?.textContent })`) }))
  }
  throw error
} finally {
  for (const task of pending.values()) clearTimeout(task.timer)
  socket?.close()
  try { if (chrome) await terminateProcess(chrome.child) }
  finally { await server?.close() }
}
