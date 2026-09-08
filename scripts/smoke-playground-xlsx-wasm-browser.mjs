import { spawn, spawnSync } from 'node:child_process'
import { existsSync, mkdtempSync, readFileSync, rmSync, statSync } from 'node:fs'
import { createServer } from 'node:http'
import { tmpdir } from 'node:os'
import { extname, resolve, sep } from 'node:path'

const root = resolve(import.meta.dirname, '..')
const dist = resolve(root, 'apps/playground/dist')
const base = '/injoffice-smoke/'
const profile = mkdtempSync(resolve(tmpdir(), 'injoffice-office-browser-'))
const requests = []
const browserRequests = []
const pageErrors = []
const loadingFailures = []
let chrome
let cdp
let staticServer
let sectionKey = 'sheets'

const mime = new Map([
  ['.css', 'text/css; charset=utf-8'],
  ['.html', 'text/html; charset=utf-8'],
  ['.js', 'text/javascript; charset=utf-8'],
  ['.json', 'application/json; charset=utf-8'],
  ['.mjs', 'text/javascript; charset=utf-8'],
  ['.svg', 'image/svg+xml'],
  ['.wasm', 'application/wasm'],
  ['.docx', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'],
  ['.pptx', 'application/vnd.openxmlformats-officedocument.presentationml.presentation'],
  ['.xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'],
])

try {
  if (!existsSync(resolve(dist, 'index.html'))) throw new Error('playground dist is missing; build it with the smoke base first')
  staticServer = createServer((request, response) => {
    const url = new URL(request.url ?? '/', 'http://127.0.0.1')
    requests.push(url.pathname)
    if (!url.pathname.startsWith(base)) {
      response.writeHead(404).end('not found')
      return
    }
    const relative = url.pathname === base ? 'index.html' : decodeURIComponent(url.pathname.slice(base.length))
    const file = resolve(dist, relative)
    if (file !== dist && !file.startsWith(`${dist}${sep}`)) {
      response.writeHead(400).end('invalid path')
      return
    }
    try {
      if (!statSync(file).isFile()) throw new Error('not a file')
      response.writeHead(200, { 'Content-Type': mime.get(extname(file)) ?? 'application/octet-stream' })
      response.end(readFileSync(file))
    } catch {
      response.writeHead(404).end('not found')
    }
  })
  await listen(staticServer)
  const address = staticServer.address()
  if (address == null || typeof address === 'string') throw new Error('static server did not expose a TCP port')
  const siteUrl = `http://127.0.0.1:${address.port}${base}`

  const executable = findChrome()
  const debugPort = await reservePort()
  chrome = spawn(executable, [
    '--headless=new',
    '--no-sandbox',
    '--disable-dev-shm-usage',
    '--disable-gpu',
    '--disable-background-networking',
    '--disable-component-update',
    '--remote-allow-origins=*',
    `--user-data-dir=${profile}`,
    `--remote-debugging-port=${debugPort}`,
    'about:blank',
  ], { stdio: ['ignore', 'ignore', 'pipe'] })
  let chromeError = ''
  chrome.stderr.setEncoding('utf8')
  chrome.stderr.on('data', (chunk) => { chromeError = `${chromeError}${chunk}`.slice(-8000) })

  const target = await waitForTarget(debugPort, chrome)
  cdp = await connectCDP(target.webSocketDebuggerUrl)
  cdp.on('Network.requestWillBeSent', ({ request }) => browserRequests.push(request.url))
  cdp.on('Runtime.exceptionThrown', ({ exceptionDetails }) => pageErrors.push(
    exceptionDetails.exception?.description ?? exceptionDetails.text,
  ))
  cdp.on('Network.loadingFailed', (event) => loadingFailures.push(`${event.errorText}: ${event.blockedReason ?? ''}`))
  await Promise.all([
    cdp.send('Network.enable'),
    cdp.send('Page.enable'),
    cdp.send('Runtime.enable'),
  ])
  await cdp.send('Page.navigate', { url: `${siteUrl}#/sheets` })

  await pollExpression(cdp, `(() => {
    const button = [...document.querySelectorAll('button')].find((item) => item.textContent?.trim() === 'Test XLSX round trip')
    if (!button) return false
    button.click()
    return true
  })()`, 'spreadsheet native view')
  await pollExpression(cdp, `(() => {
    const button = [...document.querySelectorAll('button')].find((item) => item.textContent?.includes('Server fallback'))
    return Boolean(button?.disabled && button.textContent?.includes('not configured'))
  })()`, 'disabled unconfigured server fallback')
  await pollExpression(cdp, `(() => {
    const button = [...document.querySelectorAll('button')].find((item) => item.textContent?.trim() === 'Use bundled .xlsx')
    if (!button) return false
    button.click()
    return true
  })()`, 'bundled XLSX action')
  await pollExpression(cdp, `document.querySelector('.native-status')?.textContent?.includes('Extracted ') === true`, 'browser extraction', 90_000)

  const expectedXlsx = `browser-smoke-${Date.now().toString(36)}`
  await evaluate(cdp, `(() => {
    const input = document.querySelector('.native-field input')
    if (!(input instanceof HTMLInputElement)) throw new Error('mutation input missing')
    const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set
    if (!setter) throw new Error('input value setter missing')
    setter.call(input, ${JSON.stringify(expectedXlsx)})
    input.dispatchEvent(new Event('input', { bubbles: true }))
  })()`)
  await pollExpression(cdp, `(() => {
    const button = [...document.querySelectorAll('button')].find((item) => item.textContent?.trim() === 'Save to XLSX')
    if (!button || button.disabled) return false
    button.click()
    return true
  })()`, 'enabled XLSX save action')
  await pollExpression(cdp, `(() => {
    const status = document.querySelector('.native-status')?.textContent ?? ''
    return status.includes('Saved, reopened, and verified') || Boolean(document.querySelector('.native-error'))
  })()`, 'browser apply and readback', 90_000)

  const xlsxOutcome = await evaluate(cdp, `({
    status: document.querySelector('.native-status')?.textContent ?? '',
    error: document.querySelector('.native-error')?.textContent ?? '',
  })`)
  if (!xlsxOutcome.status.includes('Saved, reopened, and verified')) {
    throw new Error(`XLSX browser readback failed: ${xlsxOutcome.error || xlsxOutcome.status || 'unknown error'}`)
  }

  const xlsxProof = await proofRows(cdp)
  if (xlsxProof.After !== expectedXlsx) throw new Error(`XLSX readback mismatch: expected ${JSON.stringify(expectedXlsx)}, received ${JSON.stringify(xlsxProof.After)}`)
  assertBrowserLocalProof(xlsxProof, 'XLSX')

  sectionKey = 'docs'
  await cdp.send('Page.navigate', { url: `${siteUrl}#/docs` })
  await pollExpression(cdp, `(() => {
    const button = [...document.querySelectorAll('button')].find((item) => item.textContent?.includes('Server fallback'))
    return Boolean(button?.disabled && button.title?.includes('VITE_INJOFFICE_API_BASE'))
  })()`, 'disabled unconfigured DOCX server fallback')
  await pollExpression(cdp, `(() => {
    const button = [...document.querySelectorAll('[data-demo-surface="docs"] button')].find((item) => item.textContent?.trim() === 'Open sample')
    if (!button) return false
    button.click()
    return true
  })()`, 'bundled DOCX action')
  await pollExpression(cdp, `document.querySelector('.native-status')?.textContent?.includes('editable passage') === true`, 'browser DOCX extraction', 90_000)
  const selectedDocxText = await evaluate(cdp, `(() => {
    const runs = [...document.querySelectorAll('.docx-editable-run')]
    const run = runs[1] ?? runs[0]
    if (!(run instanceof HTMLButtonElement)) throw new Error('Selectable DOCX text missing')
    run.click()
    return run.textContent
  })()`)
  await pollExpression(cdp, `document.querySelector('.docx-editable-run[aria-pressed="true"]')?.textContent === ${JSON.stringify(selectedDocxText)} && document.querySelector('#docx-replacement-text')?.value === ${JSON.stringify(selectedDocxText)}`, 'DOCX preview selection linked to text editor')

  const expectedDocx = `docx-browser-smoke-${Date.now().toString(36)}`
  await evaluate(cdp, `(() => {
    const input = document.querySelector('.native-field textarea')
    if (!(input instanceof HTMLTextAreaElement)) throw new Error('DOCX mutation input missing')
    const setter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')?.set
    if (!setter) throw new Error('textarea value setter missing')
    setter.call(input, ${JSON.stringify(expectedDocx)})
    input.dispatchEvent(new Event('input', { bubbles: true }))
  })()`)
  await pollExpression(cdp, `(() => {
    const button = [...document.querySelectorAll('button')].find((item) => item.textContent?.trim() === 'Apply text change')
    if (!button || button.disabled) return false
    button.click()
    return true
  })()`, 'enabled DOCX save action')
  await pollExpression(cdp, `(() => {
    const status = document.querySelector('.native-status')?.textContent ?? ''
    return status.includes('Saved, reopened, and verified') || Boolean(document.querySelector('.native-error'))
  })()`, 'browser DOCX apply and readback', 90_000)

  const docxOutcome = await evaluate(cdp, `({
    status: document.querySelector('.native-status')?.textContent ?? '',
    error: document.querySelector('.native-error')?.textContent ?? '',
  })`)
  if (!docxOutcome.status.includes('Saved, reopened, and verified')) {
    throw new Error(`DOCX browser readback failed: ${docxOutcome.error || docxOutcome.status || 'unknown error'}`)
  }
  const docxProof = await proofRows(cdp)
  if (docxProof.After !== expectedDocx) throw new Error(`DOCX readback mismatch: expected ${JSON.stringify(expectedDocx)}, received ${JSON.stringify(docxProof.After)}`)
  assertBrowserLocalProof(docxProof, 'DOCX')
  await evaluate(cdp, `(() => {
    const button = [...document.querySelectorAll('[data-demo-surface="docs"] button')].find((item) => item.textContent?.trim() === 'Undo change')
    if (!button || button.disabled) throw new Error('DOCX undo unavailable after edit')
    button.click()
  })()`)
  await pollExpression(cdp, `document.querySelector('.native-status')?.textContent?.includes('Last change undone') === true`, 'DOCX exact-byte undo', 90_000)
  await pollExpression(cdp, `document.querySelector('#docx-replacement-text')?.value === ${JSON.stringify(selectedDocxText)} && Boolean(document.querySelector('.native-download[href^="blob:"]'))`, 'DOCX undo restores editable text and download')

  sectionKey = 'pptx-native'
  await cdp.send('Page.navigate', { url: `${siteUrl}#/pptx-native` })
  await pollExpression(cdp, `document.querySelector('[data-demo-surface="pptx-native"]') !== null`, 'native PPTX view')
  await pollExpression(cdp, `(() => {
    const button = [...document.querySelectorAll('button')].find((item) => item.textContent?.includes('Server fallback'))
    return Boolean(button?.disabled && button.textContent?.includes('not configured'))
  })()`, 'disabled unconfigured PPTX server fallback')
  await pollExpression(cdp, `(() => {
    const button = [...document.querySelectorAll('button')].find((item) => item.textContent?.trim() === 'Use bundled .pptx')
    if (!button) return false
    button.click()
    return true
  })()`, 'bundled PPTX action')
  await pollExpression(cdp, `document.querySelector('.native-status')?.textContent?.includes('Extracted ') === true`, 'PPTX browser extraction', 90_000)

  const expectedPptx = `pptx-browser-smoke-${Date.now().toString(36)}`
  await evaluate(cdp, `(() => {
    const input = document.querySelector('.native-field input')
    if (!(input instanceof HTMLInputElement)) throw new Error('PPTX mutation input missing')
    const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set
    if (!setter) throw new Error('input value setter missing')
    setter.call(input, ${JSON.stringify(expectedPptx)})
    input.dispatchEvent(new Event('input', { bubbles: true }))
  })()`)
  await pollExpression(cdp, `(() => {
    const button = [...document.querySelectorAll('button')].find((item) => item.textContent?.trim() === 'Save to PPTX')
    if (!button || button.disabled) return false
    button.click()
    return true
  })()`, 'enabled PPTX save action')
  await pollExpression(cdp, `(() => {
    const status = document.querySelector('.native-status')?.textContent ?? ''
    return status.includes('Saved, reopened, and verified') || Boolean(document.querySelector('.native-error'))
  })()`, 'PPTX browser apply and readback', 90_000)

  const pptxOutcome = await evaluate(cdp, `({
    status: document.querySelector('.native-status')?.textContent ?? '',
    error: document.querySelector('.native-error')?.textContent ?? '',
  })`)
  if (!pptxOutcome.status.includes('Saved, reopened, and verified')) {
    throw new Error(`PPTX browser readback failed: ${pptxOutcome.error || pptxOutcome.status || 'unknown error'}`)
  }
  const pptxProof = await proofRows(cdp)
  if (pptxProof.After !== expectedPptx) throw new Error(`PPTX readback mismatch: expected ${JSON.stringify(expectedPptx)}, received ${JSON.stringify(pptxProof.After)}`)
  assertBrowserLocalProof(pptxProof, 'PPTX')

  const apiRequests = browserRequests.filter((value) => {
    try { return new URL(value).pathname.startsWith('/v1/') } catch { return false }
  })
  if (apiRequests.length > 0) throw new Error(`browser mode made API requests: ${apiRequests.join(', ')}`)
  if (requests.some((value) => value.startsWith('/v1/'))) throw new Error('static host received an unexpected /v1/ request')
  console.log(`Browser XLSX, DOCX, and PPTX WASM smoke passed at ${base}: extract, apply, re-extract, readback, zero /v1/ requests.`)
} catch (error) {
  const detail = [
    error instanceof Error ? error.message : String(error),
    `Static requests: ${requests.join(', ')}`,
    `Browser requests: ${browserRequests.join(', ')}`,
    `Page errors: ${pageErrors.join(', ')}`,
    `Loading failures: ${loadingFailures.join(', ')}`,
  ].join('\n')
  if (chrome && chrome.exitCode !== null) {
    throw new Error(`${detail}\nChrome exited ${chrome.exitCode}.`, { cause: error })
  }
  throw new Error(detail, { cause: error })
} finally {
  cdp?.close()
  if (chrome && chrome.exitCode === null) {
    chrome.kill('SIGTERM')
    await Promise.race([
      new Promise((resolvePromise) => chrome.once('exit', resolvePromise)),
      delay(3_000),
    ])
    if (chrome.exitCode === null) {
      const forcedExit = new Promise((resolvePromise) => chrome.once('exit', resolvePromise))
      chrome.kill('SIGKILL')
      await forcedExit
    }
  }
  if (staticServer) await closeServer(staticServer)
  try {
    rmSync(profile, { recursive: true, force: true, maxRetries: 10, retryDelay: 100 })
  } catch (error) {
    console.warn(`Could not remove temporary Chrome profile ${profile}:`, error)
  }
}

function findChrome() {
  const candidates = [
    process.env.CHROME_PATH,
    process.env.CHROME_BIN,
    '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
    '/Applications/Chromium.app/Contents/MacOS/Chromium',
    'google-chrome',
    'google-chrome-stable',
    'chromium',
    'chromium-browser',
  ].filter(Boolean)
  for (const candidate of candidates) {
    if (candidate.includes(sep) && !existsSync(candidate)) continue
    if (spawnSync(candidate, ['--version'], { stdio: 'ignore' }).status === 0) return candidate
  }
  throw new Error('Chrome/Chromium is required for the native Office browser smoke; set CHROME_PATH explicitly')
}

async function proofRows(client) {
  return evaluate(client, `Object.fromEntries([...document.querySelectorAll('.native-proof-row')].map((row) => [row.querySelector('dt')?.textContent, row.querySelector('dd')?.textContent]))`)
}

function assertBrowserLocalProof(proof, format) {
  if (proof.Runtime !== 'browser-local') throw new Error(`${format} has unexpected runtime proof ${JSON.stringify(proof.Runtime)}`)
  if (proof.Artifact !== 'browser memory') throw new Error(`${format} has unexpected artifact proof ${JSON.stringify(proof.Artifact)}`)
}

function listen(server) {
  return new Promise((resolvePromise, reject) => {
    server.once('error', reject)
    server.listen(0, '127.0.0.1', resolvePromise)
  })
}

function closeServer(server) {
  return new Promise((resolvePromise) => server.close(resolvePromise))
}

async function reservePort() {
  const server = createServer()
  await listen(server)
  const address = server.address()
  if (address == null || typeof address === 'string') throw new Error('could not reserve a Chrome debugging port')
  const { port } = address
  await closeServer(server)
  return port
}

async function waitForTarget(port, process) {
  const deadline = Date.now() + 15_000
  while (Date.now() < deadline) {
    if (process.exitCode !== null) throw new Error('Chrome exited before exposing its debugging target')
    try {
      const response = await fetch(`http://127.0.0.1:${port}/json/list`)
      const targets = await response.json()
      const target = targets.find((item) => item.type === 'page')
      if (target?.webSocketDebuggerUrl) return target
    } catch {
      // Chrome is still starting.
    }
    await delay(50)
  }
  throw new Error('timed out waiting for the Chrome debugging target')
}

async function connectCDP(url) {
  if (typeof WebSocket !== 'function') {
    throw new Error('Node.js 22 or newer is required for the browser smoke (global WebSocket is unavailable)')
  }
  const socket = new WebSocket(url)
  await new Promise((resolvePromise, reject) => {
    socket.addEventListener('open', resolvePromise, { once: true })
    socket.addEventListener('error', reject, { once: true })
  })
  let nextId = 0
  const pending = new Map()
  const listeners = new Map()
  socket.addEventListener('message', (event) => {
    const message = JSON.parse(event.data)
    if (message.id !== undefined) {
      const request = pending.get(message.id)
      if (!request) return
      pending.delete(message.id)
      if (message.error) request.reject(new Error(message.error.message))
      else request.resolve(message.result)
      return
    }
    for (const listener of listeners.get(message.method) ?? []) listener(message.params)
  })
  return {
    send(method, params = {}) {
      const id = ++nextId
      return new Promise((resolvePromise, reject) => {
        pending.set(id, { resolve: resolvePromise, reject })
        socket.send(JSON.stringify({ id, method, params }))
      })
    },
    on(method, listener) {
      const current = listeners.get(method) ?? []
      current.push(listener)
      listeners.set(method, current)
    },
    close() { socket.close() },
  }
}

async function evaluate(client, expression) {
  // Continuous showcase sections stay mounted. Read and mutate only the target
  // format's inputs, status, and proof rows; a previous XLSX result is not DOCX
  // evidence. Keep compatibility with standalone pages without scroll sections.
  const scopedExpression = expression.replaceAll('document.querySelectorAll(', 'nativeQueryAll(').replaceAll('document.querySelector(', 'nativeQuery(')
  const response = await client.send('Runtime.evaluate', { expression: `{
    const nativeSection = document.querySelector(${JSON.stringify(`[data-scroll-section="${sectionKey}"]`)}) ?? (document.querySelector('[data-scroll-section]') ? null : document);
    const nativeQuery = selector => nativeSection?.querySelector(selector) ?? null;
    const nativeQueryAll = selector => nativeSection?.querySelectorAll(selector) ?? [];
    ${scopedExpression}
  }`, awaitPromise: true, returnByValue: true })
  if (response.exceptionDetails) throw new Error(response.exceptionDetails.exception?.description ?? response.exceptionDetails.text)
  return response.result.value
}

async function pollExpression(client, expression, label, timeout = 20_000) {
  const deadline = Date.now() + timeout
  while (Date.now() < deadline) {
    if (await evaluate(client, expression)) return
    await delay(50)
  }
  const diagnostic = await evaluate(client, `({
    href: location.href,
    body: document.body?.innerText?.slice(0, 2000) ?? '',
    error: document.querySelector('.native-error')?.textContent ?? '',
    status: document.querySelector('.native-status')?.textContent ?? '',
    readyState: document.readyState,
  })`)
  throw new Error(`timed out waiting for ${label}: ${JSON.stringify(diagnostic)}`)
}

function delay(milliseconds) {
  return new Promise((resolvePromise) => setTimeout(resolvePromise, milliseconds))
}
