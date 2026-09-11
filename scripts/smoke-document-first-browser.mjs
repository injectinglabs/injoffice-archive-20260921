import assert from 'node:assert/strict'
import { existsSync, mkdtempSync, mkdirSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { resolve } from 'node:path'
import { launchChromeForCDP, terminateProcess } from './chrome-cdp-startup.mjs'
import { startShowcaseServer } from './showcase-smoke-server.mjs'

const output = process.env.SHOWCASE_OUTPUT || mkdtempSync(resolve(tmpdir(), 'injoffice-document-first-'))
mkdirSync(output, { recursive: true })
let server, chrome, socket, sequence = 0
const pending = new Map(), errors = []
const delay = ms => new Promise(resolve => setTimeout(resolve, ms))
function send(method, params = {}) {
  return new Promise((resolve, reject) => {
    const id = ++sequence
    const timer = setTimeout(() => { pending.delete(id); reject(new Error(`Timeout: ${method}`)) }, 30000)
    pending.set(id, { resolve, reject, timer }); socket.send(JSON.stringify({ id, method, params }))
  })
}
async function evaluate(expression) {
  const result = await send('Runtime.evaluate', { expression, returnByValue: true, awaitPromise: true })
  if (result.exceptionDetails) throw new Error(result.exceptionDetails.exception?.description || result.exceptionDetails.text)
  return result.result.value
}
async function until(expression) {
  const deadline = Date.now() + 45000
  while (Date.now() < deadline) { if (await evaluate(`Boolean(${expression})`)) return; await delay(100) }
  throw new Error(`Timed out: ${expression}`)
}
async function screenshot(name) {
  const { data } = await send('Page.captureScreenshot', { format: 'png' })
  writeFileSync(resolve(output, `${name}.png`), Buffer.from(data, 'base64'))
}
try {
  server = await startShowcaseServer(resolve(import.meta.dirname, '../apps/playground/dist'))
  chrome = await launchChromeForCDP({
    executable: process.env.CHROME_BIN ?? ['/Applications/Google Chrome.app/Contents/MacOS/Google Chrome', '/usr/bin/google-chrome', '/usr/bin/chromium'].find(existsSync) ?? 'google-chrome',
    createProfile: () => mkdtempSync(resolve(tmpdir(), 'injoffice-document-first-chrome-')),
  })
  socket = new WebSocket(chrome.target.webSocketDebuggerUrl)
  await new Promise((resolve, reject) => { socket.addEventListener('open', resolve, { once: true }); socket.addEventListener('error', reject, { once: true }) })
  socket.addEventListener('message', ({ data }) => {
    const message = JSON.parse(data)
    if (message.id) {
      const task = pending.get(message.id); if (!task) return
      pending.delete(message.id); clearTimeout(task.timer)
      if (message.error) task.reject(new Error(message.error.message)); else task.resolve(message.result)
    } else if (message.method === 'Runtime.exceptionThrown') errors.push(message.params.exceptionDetails.text)
  })
  await send('Runtime.enable'); await send('Page.enable')
  await send('Emulation.setEmulatedMedia', { features: [{ name: 'prefers-color-scheme', value: 'light' }] })
  await send('Page.navigate', { url: server.url })
  for (const width of [1440, 768, 390, 320]) {
    await send('Emulation.setDeviceMetricsOverride', { width, height: 900, deviceScaleFactor: 1, mobile: width < 760 })
    for (const tool of ['sheets', 'docs', 'slides', 'pdf']) {
      await evaluate(`location.hash = '#/${tool}'`)
      const root = `document.querySelector('[data-scroll-section="${tool}"] [data-workspace-panel="agent"]')`
      await until(`${root} && !${root}.hidden && ${root}.querySelector('[data-agent-prepare]:not(:disabled)')`)
      await delay(200)
      assert.equal(await evaluate(`(() => {
        const header = document.querySelector('.app-header-inner');
        const actions = header.querySelector('.app-header-actions');
        const edge = header.getBoundingClientRect().right - parseFloat(getComputedStyle(header).paddingRight);
        return Math.abs(actions.getBoundingClientRect().right - edge) <= 1
          && !!actions.querySelector('.scheme-toggle') && !!actions.querySelector('.github-link')
          && !header.textContent.includes('About this demo');
      })()`), true, `Header controls are right-aligned without the demo disclosure at ${tool}/${width}`)
      const geometry = await evaluate(`(() => {
        const panel = ${root}, artifact = panel.querySelector('.agent-artifact'), footer = panel.querySelector('.agent-demo__document > footer');
        const documentBox = panel.querySelector('.agent-demo__document').getBoundingClientRect(), assistant = panel.querySelector('.agent-demo__inspector').getBoundingClientRect();
        const contentBottom = Math.max(...Array.from(artifact.querySelectorAll('p, h3, table')).map(el => el.getBoundingClientRect().bottom));
        return { width: document.documentElement.scrollWidth, contentBottom, footerTop: footer.getBoundingClientRect().top, adjacent: assistant.left >= documentBox.right - 1,
          boxes: Array.from(document.querySelectorAll('body, .app-shell, .app-header-inner, .app-header-actions, .app-main, .agent-demo__workspace, .tool-workspace__navigation')).map(el => ({ class: el.className, width: el.clientWidth, scroll: el.scrollWidth, right: el.getBoundingClientRect().right, x: scrollX })),
          overflow: Array.from(document.querySelectorAll('body *')).filter(el => !el.closest('[hidden], .agent-artifact__sheet-scroll') && el.getBoundingClientRect().right > innerWidth + 1 && getComputedStyle(el).visibility !== 'hidden').slice(0, 16).map(el => ({ tag: el.tagName, class: String(el.className), right: el.getBoundingClientRect().right })) };
      })()`)
      await screenshot(`${tool}-${width}`)
      assert.ok(geometry.width <= width, `No page overflow at ${tool}/${width}: ${JSON.stringify(geometry)}`)
      assert.ok(geometry.contentBottom <= geometry.footerTop + 1, `Document text stays above its note at ${tool}/${width}: ${JSON.stringify(geometry)}`)
      if (width === 1440) assert.ok(geometry.adjacent, 'assistant sits beside its document')
      assert.equal(await evaluate(`${root}.querySelector('[data-agent-technical]').open`), false, 'technical details start collapsed')
    }
  }
  await evaluate(`document.querySelector('.scheme-toggle button:last-child').click()`)
  await delay(300)
  await screenshot('mobile-dark')
  assert.deepEqual(errors, [])
  console.log(JSON.stringify({ status: 'passed', output, checks: ['four default document tasks', 'adjacent assistant', '320–1440px layouts', 'no text overlap', 'collapsed technical evidence'] }))
} catch (error) {
  if (socket?.readyState === WebSocket.OPEN) await screenshot('failure').catch(() => {})
  console.error(JSON.stringify({ output, errors })); throw error
} finally {
  socket?.close()
  if (chrome) await terminateProcess(chrome.child)
  await server?.close()
}
