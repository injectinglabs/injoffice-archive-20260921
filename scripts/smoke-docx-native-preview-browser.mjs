// Node 22+, Go, Chrome, and built workspace packages are required. The test
// starts its own local helper and Vite server; it never uses a running user app.
import { spawn, spawnSync } from 'node:child_process'
import { createHash } from 'node:crypto'
import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { createServer } from 'node:net'
import { tmpdir } from 'node:os'
import { resolve, sep } from 'node:path'
import { launchChromeForCDP, terminateProcess } from './chrome-cdp-startup.mjs'
import { startShowcaseDevServer } from './showcase-smoke-dev-server.mjs'

const root = resolve(process.env.DOCX_NATIVE_SMOKE_ROOT || resolve(import.meta.dirname, '..'))
const scratch = mkdtempSync(resolve(tmpdir(), 'injoffice-docx-native-smoke-'))
const artifacts = process.env.SHOWCASE_OUTPUT ? resolve(process.env.SHOWCASE_OUTPUT) : mkdtempSync(resolve(tmpdir(), 'injoffice-docx-native-screenshots-'))
mkdirSync(artifacts, { recursive: true })
const profiles = []
const errors = []
const docs = `document.querySelector('[data-demo-surface="docs"]')`
const native = `document.querySelector('[aria-label="Native document pages"]')`
let helper, server, chrome, cdp
let helperLog = ''
const previousApi = process.env.VITE_INJOFFICE_API_BASE
const previousServer = process.env.INJOFFICE_SERVER

try {
  const worker = resolve(root, 'apps/docx-page-paint-worker/dist/worker.js')
  if (!existsSync(worker)) throw new Error('Build the workspace packages and DOCX page-paint worker before running this smoke.')
  const fixture = resolve(scratch, 'native-two-pages.docx')
  await command('go', ['run', './cmd/nativepreviewfixture', '-font', resolve(root, 'node_modules/dejavu-fonts-ttf/ttf/DejaVuSans.ttf'), '-out', fixture], resolve(root, 'go/docxpatch'))
  const originalHash = hash(readFileSync(fixture))
  const binary = resolve(scratch, process.platform === 'win32' ? 'injoffice-server.exe' : 'injoffice-server')
  await command('go', ['build', '-o', binary, './cmd/injoffice-server'], resolve(root, 'go/injoffice-server'))
  const port = await unusedPort()
  helper = spawn(binary, ['-addr', `127.0.0.1:${port}`, '-artifacts', resolve(scratch, 'artifacts'), '-docx-preview-worker', worker], { cwd: root, stdio: ['ignore', 'ignore', 'pipe'] })
  helper.stderr.setEncoding('utf8')
  helper.stderr.on('data', chunk => { helperLog = `${helperLog}${chunk}`.slice(-8000) })
  helper.on('error', error => { helperLog = error.message })
  await poll(async () => { try { return (await fetch(`http://127.0.0.1:${port}/healthz`, { signal: AbortSignal.timeout(1000) })).ok } catch { return false } }, 'isolated helper startup')
  // A nonempty relative base explicitly enables the helper via Vite's local
  // proxy, without introducing cross-origin permissions or an upload fallback.
  process.env.VITE_INJOFFICE_API_BASE = '.'
  process.env.INJOFFICE_SERVER = `http://127.0.0.1:${port}`
  server = await startShowcaseDevServer(resolve(root, 'apps/playground'))
  chrome = await launchChromeForCDP({ executable: findChrome(), createProfile: () => {
    const profile = mkdtempSync(resolve(tmpdir(), 'injoffice-docx-native-chrome-')); profiles.push(profile); return profile
  } })
  cdp = await connectCDP(chrome.target.webSocketDebuggerUrl)
  cdp.on('Runtime.exceptionThrown', ({ exceptionDetails }) => errors.push(exceptionDetails.exception?.description ?? exceptionDetails.text))
  await cdp.send('Runtime.enable')
  await cdp.send('Page.enable')
  await cdp.send('Page.addScriptToEvaluateOnNewDocument', { source: `{
    const originalFetch = window.fetch.bind(window); window.__nativeDocxPosts = [];
    window.fetch = async (input, init) => {
      if (init?.method === 'POST' && String(input).includes('/v1/')) {
        const bytes = init.body instanceof Blob ? await init.body.arrayBuffer() : new TextEncoder().encode(String(init.body));
        const digest = await crypto.subtle.digest('SHA-256', bytes);
        window.__nativeDocxPosts.push({ url: String(input), hash: [...new Uint8Array(digest)].map(n => n.toString(16).padStart(2, '0')).join('') });
      }
      return originalFetch(input, init);
    };
  }` })
  await cdp.send('Emulation.setDeviceMetricsOverride', { width: 1440, height: 1100, deviceScaleFactor: 1, mobile: false })
  await cdp.send('Page.navigate', { url: `${server.url}#/docs?feature=editor` })
  await poll(() => evaluate(`Boolean(document.querySelector('[aria-label="Open a DOCX file"]'))`), 'DOCX workbench')
  await upload(fixture)
  await poll(() => evaluate(`${native}?.textContent.includes('Nothing is uploaded')`), 'native preview consent')
  await assert(`window.__nativeDocxPosts.length === 0`, 'opening a real DOCX uploads nothing')
  await click('Upload to helper and render native pages')
  await poll(() => evaluate(`${native}?.querySelector('svg[aria-label="Native document page 1"] path') !== null && ${native}?.textContent.includes('2 native pages')`), 'real font-shaped native pages', 45000)
  await assert(`window.__nativeDocxPosts.length === 1 && window.__nativeDocxPosts[0].hash === ${JSON.stringify(originalHash)}`, 'preview uploads exact original bytes only after consent')
  await assert(`${native}.querySelectorAll('svg').length === 1 && ${native}.querySelector('svg').getBoundingClientRect().height > 100`, 'one bounded native page is mounted')
  await screenshot('docx-native-page-one.png')
  await click('Next native page')
  await poll(() => evaluate(`${native}?.querySelector('svg[aria-label="Native document page 2"] path') !== null`), 'next native page')
  await assert(`${native}.querySelectorAll('svg').length === 1`, 'navigation keeps a single mounted SVG')
  await screenshot('docx-native-page-two.png')
  await click('Previous native page')
  await poll(() => evaluate(`${native}?.querySelector('svg[aria-label="Native document page 1"] path') !== null`), 'previous native page')
  await assert(`${docs}?.dataset.demoDirty !== 'true' && window.__nativeDocxPosts.every(request => request.url.endsWith('/v1/docx/page-preview'))`, 'preview never mutates the document')
  if (hash(readFileSync(fixture)) !== originalHash) throw new Error('The source fixture changed')
  // This existing real DOCX has no embedded qualified font assets. It must
  // retain its approximate content view rather than invent native glyphs.
  const unsupported = resolve(scratch, 'unsupported-font.docx')
  writeFileSync(unsupported, Buffer.from(readFileSync(resolve(root, 'apps/playground/public/native-docx/northstar-launch-brief.docx.b64'), 'utf8').trim(), 'base64'))
  await upload(unsupported)
  await poll(() => evaluate(`${native}?.textContent.includes('Nothing is uploaded') && ${native}?.querySelector('svg') === null`), 'source replacement clears stale pages')
  await assert(`window.__nativeDocxPosts.length === 1`, 'replacement document also requires explicit consent')
  await click('Upload to helper and render native pages')
  await poll(() => evaluate(`${native}?.textContent.includes('original file is unchanged')`), 'unsupported document explicitly refused', 45000)
  await assert(`${native}.querySelector('svg') === null && document.querySelectorAll('.docx-editable-run').length > 0 && ${docs}?.dataset.demoDirty !== 'true'`, 'refusal retains approximate editable content and original source')
  await screenshot('docx-native-refusal.png')
  if (errors.length) throw new Error(`Browser exceptions: ${errors.join('\n')}`)
  console.log(JSON.stringify({ result: 'PASS', checks: ['explicit upload consent', 'real embedded-font shaping and pagination', 'native SVG glyphs', 'bounded page navigation', 'original source unchanged', 'source replacement clears stale output', 'unsupported rendering refusal'], screenshots: artifacts }, null, 2))
} catch (error) {
  console.error(`Native DOCX screenshots: ${artifacts}\nHelper diagnostics: ${helperLog}`)
  if (cdp) {
    await screenshot('failure.png').catch(() => undefined)
    console.error(await evaluate(`document.body.innerText.slice(-6000)`).catch(() => 'No browser diagnostic'))
  }
  throw error
} finally {
  cdp?.close()
  try { if (chrome) await terminateProcess(chrome.child) }
  finally {
    try { await server?.close() }
    finally {
      if (helper) await terminateProcess(helper)
      for (const profile of profiles) rmSync(profile, { recursive: true, force: true, maxRetries: 10, retryDelay: 100 })
      rmSync(scratch, { recursive: true, force: true, maxRetries: 10, retryDelay: 100 })
      if (previousApi === undefined) delete process.env.VITE_INJOFFICE_API_BASE; else process.env.VITE_INJOFFICE_API_BASE = previousApi
      if (previousServer === undefined) delete process.env.INJOFFICE_SERVER; else process.env.INJOFFICE_SERVER = previousServer
    }
  }
}

function hash(bytes) { return createHash('sha256').update(bytes).digest('hex') }
async function command(executable, args, cwd) {
  await new Promise((resolvePromise, reject) => {
    const child = spawn(executable, args, { cwd, stdio: ['ignore', 'ignore', 'pipe'] }); let stderr = ''
    const timer = setTimeout(() => { child.kill('SIGKILL'); reject(new Error(`${executable} timed out`)) }, 120000)
    child.stderr.on('data', chunk => { stderr = `${stderr}${chunk}`.slice(-8000) })
    child.on('error', error => { clearTimeout(timer); reject(error) })
    child.on('exit', code => { clearTimeout(timer); code === 0 ? resolvePromise() : reject(new Error(`${executable} exited ${code}: ${stderr}`)) })
  })
}
async function unusedPort() {
  const socket = createServer()
  await new Promise((done, reject) => { socket.once('error', reject); socket.listen(0, '127.0.0.1', done) })
  const port = socket.address().port
  await new Promise(done => socket.close(done)); return port
}
async function upload(path) {
  const handle = await cdp.send('Runtime.evaluate', { expression: `document.querySelector('[aria-label="Open a DOCX file"]')` })
  const node = await cdp.send('DOM.describeNode', { objectId: handle.result.objectId })
  await cdp.send('DOM.setFileInputFiles', { backendNodeId: node.node.backendNodeId, files: [path] })
  await cdp.send('Runtime.releaseObject', { objectId: handle.result.objectId })
}
async function evaluate(expression) {
  const value = await cdp.send('Runtime.evaluate', { expression, awaitPromise: true, returnByValue: true })
  if (value.exceptionDetails) throw new Error(value.exceptionDetails.exception?.description ?? value.exceptionDetails.text)
  return value.result.value
}
async function assert(expression, label) { if (!await evaluate(expression)) throw new Error(`Failed: ${label}`) }
async function poll(check, label, timeout = 30000) {
  const end = Date.now() + timeout
  while (Date.now() < end) { if (await check()) return; await new Promise(done => setTimeout(done, 75)) }
  throw new Error(`Timed out: ${label}`)
}
async function click(label) {
  const handle = await cdp.send('Runtime.evaluate', { expression: 'globalThis' })
  try {
    await poll(async () => {
      const value = await cdp.send('Runtime.callFunctionOn', {
        functionDeclaration: `function (label) {
          const section = document.querySelector('[aria-label="Native document pages"]');
          const button = [...(section?.querySelectorAll('button') ?? [])].find(button => button.textContent.trim() === label);
          if (!button || button.disabled) return false;
          button.click();
          return true;
        }`,
        objectId: handle.result.objectId,
        arguments: [{ value: label }],
        returnByValue: true,
      })
      if (value.exceptionDetails) throw new Error(value.exceptionDetails.exception?.description ?? value.exceptionDetails.text)
      return value.result.value
    }, label)
  } finally { await cdp.send('Runtime.releaseObject', { objectId: handle.result.objectId }) }
}
async function screenshot(name) {
  await evaluate(`${native}?.scrollIntoView({ block: 'start' }); window.scrollBy(0, -80)`)
  const result = await cdp.send('Page.captureScreenshot', { format: 'png' })
  writeFileSync(resolve(artifacts, name), Buffer.from(result.data, 'base64'))
}
function findChrome() {
  for (const candidate of [process.env.CHROME_PATH, process.env.CHROME_BIN, '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome', '/Applications/Chromium.app/Contents/MacOS/Chromium', 'google-chrome', 'google-chrome-stable', 'chromium', 'chromium-browser'].filter(Boolean)) {
    if (candidate.includes(sep) && !existsSync(candidate)) continue
    if (spawnSync(candidate, ['--version'], { stdio: 'ignore', timeout: 5000 }).status === 0) return candidate
  }
  throw new Error('Chrome/Chromium is required; set CHROME_PATH')
}
async function connectCDP(url) {
  const socket = new WebSocket(url)
  await new Promise((done, reject) => { socket.addEventListener('open', done, { once: true }); socket.addEventListener('error', reject, { once: true }) })
  let next = 0
  const pending = new Map(), listeners = new Map()
  socket.addEventListener('message', event => {
    const message = JSON.parse(event.data)
    if (message.id !== undefined) {
      const request = pending.get(message.id); if (!request) return
      pending.delete(message.id); clearTimeout(request.timer)
      if (message.error) request.reject(new Error(message.error.message)); else request.resolve(message.result)
    } else for (const callback of listeners.get(message.method) ?? []) callback(message.params)
  })
  return {
    send(method, params = {}) { return new Promise((resolvePromise, reject) => { const id = ++next; const timer = setTimeout(() => { pending.delete(id); reject(new Error(`CDP timeout: ${method}`)) }, 30000); pending.set(id, { resolve: resolvePromise, reject, timer }); socket.send(JSON.stringify({ id, method, params })) }) },
    on(method, callback) { listeners.set(method, [...listeners.get(method) ?? [], callback]) },
    close() { for (const request of pending.values()) { clearTimeout(request.timer); request.reject(new Error('CDP closed')) } pending.clear(); socket.close() },
  }
}
