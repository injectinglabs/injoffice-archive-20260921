import assert from 'node:assert/strict'
import { existsSync, mkdirSync, mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { resolve } from 'node:path'
import { launchChromeForCDP, terminateProcess } from './chrome-cdp-startup.mjs'
import { startShowcaseServer } from './showcase-smoke-server.mjs'
import { startShowcaseDevServer } from './showcase-smoke-dev-server.mjs'
import { startShowcaseProposalMock } from './showcase-smoke-proposal-mock.mjs'

const output = process.env.SHOWCASE_OUTPUT ? resolve(process.env.SHOWCASE_OUTPUT) : mkdtempSync(resolve(tmpdir(), 'injoffice-showcase-review-'))
mkdirSync(output, { recursive: true })
let chrome, socket, staticServer
let proposalMock, restoreProposalEnv
let sequence = 0
let scopeKey = 'overview'
const pending = new Map()
const errors = []
const heldRequests = []
const proposalRequests = []
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
  } else if (message.method === 'Network.requestWillBeSent') {
    const { request } = message.params
    if (/\/api\/agent\/(?:mock-propose|propose)(?:\?|$)/.test(request.url)) proposalRequests.push(request)
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
  // Multiple editors now remain mounted. Every editor assertion and action is
  // scoped to the section selected by this test, never the first agent in DOM.
  // Shell, catalogue, and the one shared source drawer remain document-wide.
  const globalSelectors = ['.app-', '.scheme-toggle', '.showcase-', '#showcase-', '.source-proof-layer', '.source-proof-close', '.guided-recipe', '.demo-source', '[role="dialog"]', '[role=progressbar]', '[data-scroll-section]']
  const scopedExpression = expression.replaceAll('document.querySelectorAll(', 'testQueryAll(').replaceAll('document.querySelector(', 'testQuery(')
  const result = await send('Runtime.evaluate', { expression: `{
    const smokeSection = document.querySelector(${JSON.stringify(`[data-scroll-section="${scopeKey}"]`)});
    const smokeRoot = selector => ${JSON.stringify(globalSelectors)}.some(prefix => selector.startsWith(prefix)) ? document : smokeSection;
    const testQuery = selector => smokeRoot(selector)?.querySelector(selector) ?? null;
    const testQueryAll = selector => smokeRoot(selector)?.querySelectorAll(selector) ?? [];
    ${scopedExpression}
  }`, returnByValue: true, awaitPromise: true })
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
const clickButton = (label) => evaluate(`Array.from(document.querySelectorAll('button')).find(button => button.textContent.trim() === ${JSON.stringify(label)}).click()`)
const agentReady = `document.querySelector('.agent-demo__status')?.dataset.state === 'ready' && Array.from(document.querySelectorAll('button')).some(button => button.textContent.trim() === 'Run agent' && !button.disabled)`
const agentWrites = () => evaluate(`Number(document.querySelector('[data-agent-native-writes]')?.textContent)`)
const setAgentRequest = (request) => evaluate(`(() => { const input = document.querySelector('[data-agent-request]'); Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set.call(input, ${JSON.stringify(request)}); input.dispatchEvent(new Event('input', { bubbles: true })); })()`)
async function prepareAgentRequest(request, target) {
  await clickButton('Reload sample')
  await until(agentReady, 'real XLSX sample reload', 90_000)
  await setAgentRequest(request)
  await clickButton('Run agent')
  await until(`document.querySelector('.agent-demo__status')?.dataset.state === 'awaiting-approval'`, 'editable request produces a reviewable plan', 90_000)
  assert.ok((await evaluate(`document.querySelector('.agent-diff')?.textContent`)).includes(target), `request resolves real target ${target}`)
  assert.match(await evaluate(`document.querySelector('[data-agent-trace]')?.textContent`), /office\.read/, 'trace includes a real bounded read')
  assert.equal(await agentWrites(), 0, 'planning does not write native bytes')
  assert.equal(await evaluate(`document.querySelector('.agent-approval input').checked`), false, 'each mock proposal requires fresh human approval')
  assert.equal(await evaluate(`Array.from(document.querySelectorAll('button')).find(button => button.textContent.trim() === 'Commit approved change').disabled`), true, 'mock endpoint cannot authorize its own proposed edit')
}
async function approveAgentCommit() {
  await click('.agent-approval input[type=checkbox]')
  await clickButton('Commit approved change')
}
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
  const [surface, query = ''] = path.split('?')
  scopeKey = surface === 'agent' ? `agent-${new URLSearchParams(query).get('format') ?? 'sheets'}` : surface
  await evaluate(`location.hash = '#/${path}'`)
  await until(`document.querySelector('.app-shell')?.dataset.surface === ${JSON.stringify(surface)} && !document.querySelector('.demo-loading') && (smokeSection?.dataset.scrollState === 'ready' || ${JSON.stringify(surface)} === 'overview')`, path, 90_000)
}
try {
  assert.ok(!(process.argv.includes('--built') && process.argv.includes('--dev')), 'choose either --built or --dev')
  if (process.argv.includes('--built')) staticServer = await startShowcaseServer(resolve(import.meta.dirname, '../apps/playground/dist'))
  if (process.argv.includes('--dev')) {
    proposalMock = await startShowcaseProposalMock()
    const saved = ['INJOFFICE_AGENT_PROPOSAL_URL', 'INJOFFICE_AGENT_PROPOSAL_TOKEN'].map((key) => [key, process.env[key]])
    restoreProposalEnv = () => { for (const [key, value] of saved) { if (value === undefined) delete process.env[key]; else process.env[key] = value } }
    // Override even a pre-existing provider configuration: this test must never
    // contact a real model endpoint or forward the developer's credentials.
    process.env.INJOFFICE_AGENT_PROPOSAL_URL = proposalMock.url
    delete process.env.INJOFFICE_AGENT_PROPOSAL_TOKEN
    staticServer = await startShowcaseDevServer(resolve(import.meta.dirname, '../apps/playground'))
  }
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
  await send('Network.enable')
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
  scopeKey = 'pdf'
  await until(`document.querySelector('.app-shell')?.dataset.surface === 'pdf' && !document.querySelector('.demo-loading')`, 'filtered result opens')
  await click('.demo-back')
  scopeKey = 'overview'
  await until(`document.querySelector('#showcase-query')?.value === 'redact' && document.querySelectorAll('.showcase-item').length === 1`, 'Back preserves search')
  assert.deepEqual(await evaluate(`Array.from(document.querySelectorAll('.showcase-filter-group button[aria-pressed=true]')).map(button => button.textContent)`), ['Edit files', 'PDF'], 'Back preserves task and file-type filters')
  await evaluate(`(() => { const input = document.querySelector('#showcase-query'); Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(input, 'zzznomatch'); input.dispatchEvent(new Event('input', { bubbles: true })); })()`)
  await until(`!!document.querySelector('.showcase-empty')`, 'empty search')
  await click('.showcase-empty button')
  await until(`document.querySelectorAll('.showcase-item').length === 19`, 'clear filters')
  assert.equal(await evaluate(`new Set([...document.querySelectorAll('.showcase-item')].map(item => item.getAttribute('href'))).size`), 19, 'catalogue examples have distinct routes')
  await route('charts')
  assert.equal(await evaluate(`document.querySelector('.app-shell').dataset.navigation`), 'scroll')
  assert.equal(await evaluate(`document.querySelectorAll('.app-sidebar a[href^="#/agent?format="]').length`), 4, 'all AI formats appear in navigation')

  // A cold section loads independently: already visited editors stay mounted,
  // and every section heading remains usable while its script is held.
  await send('Fetch.enable', { patterns: [{ urlPattern: '*', resourceType: 'Script', requestStage: 'Request' }] })
  await evaluate(`window.__showcasePreviousStage = document.querySelector('.demo-stage'); location.hash = '#/pptx-render'`)
  scopeKey = 'pptx-render'
  const coldStart = Date.now()
  while (heldRequests.length === 0 && Date.now() - coldStart < 10_000) await new Promise((resolve) => setTimeout(resolve, 50))
  assert.ok(heldRequests.length > 0, 'cold navigation requests a lazy script')
  await until(`smokeSection?.dataset.scrollState === 'loading'`, 'cold section reports loading')
  assert.equal(await evaluate(`window.__showcasePreviousStage.isConnected && document.querySelectorAll('[data-scroll-section]').length === 20 && Array.from(document.querySelectorAll('[data-scroll-section]')).every(section => section.querySelector('h1,h2'))`), true, 'loaded editor and all headings remain during cold section navigation')
  await send('Fetch.disable')
  await until(`smokeSection?.dataset.scrollState === 'ready' && !document.querySelector('.demo-loading')`, 'cold section completes', 90_000)
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
  await until(`(!document.querySelector('.source-proof-layer') || document.querySelector('.source-proof-layer').hidden) && document.activeElement?.classList.contains('source-proof-trigger')`, 'Escape restores focus')
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
  for (let i = 0; i < 2; i++) { await route('charts'); await route('sheets'); await until(`document.querySelector('[data-demo-surface="sheets"] canvas')`, 'preserved sheet') }
  await evaluate(`location.hash = '#/sheets?view=native'`)
  await until(`!!document.querySelector('.native-toolbar')`, 'same-surface deep link switches mode')
  await route('collab')
  await until(`document.querySelectorAll('canvas').length > 0`, 'collaboration rendered')
  for (const tool of ['sheets', 'docs', 'slides', 'pdf']) {
    const fileFormat = { sheets: 'xlsx', docs: 'docx', slides: 'pptx', pdf: 'pdf' }[tool]
    await route(`agent?format=${tool}`)
    await until(`document.querySelector('[data-agent-tool=${tool}]') && document.querySelector('.agent-demo__status')?.dataset.state === 'ready' && Array.from(document.querySelectorAll('button')).some(button => button.textContent.trim() === 'Run agent' && !button.disabled)`, `${tool} AI deep link`, 90_000)
    const boundary = await evaluate(`document.querySelector('[data-agent-boundary]')?.textContent`)
    const defaultRequest = await evaluate(`document.querySelector('[data-agent-request]').value`)
    assert.match(boundary, new RegExp(`Simulated AI proposal.*real ${fileFormat}.*file write.*verification.*no language model`, 'i'), `${tool} distinguishes mocked proposal from real file proof`)
    assert.equal(await evaluate(`document.querySelector('[data-agent-proposal-source]').value`), 'mock', `${tool} defaults to the bundled mock`)
    const proposalsBeforeRun = proposalRequests.length
    if (tool === 'sheets') {
      assert.equal(await evaluate(`document.querySelector('[data-agent-proposal-source]').value`), 'mock', 'bundled mock is the zero-configuration default')
      const mockNotice = await evaluate(`document.querySelector('[data-agent-mock]')?.textContent`)
      assert.match(mockNotice, /mock/i, 'mock proposer is explicitly labelled')
      assert.match(mockNotice, /no (?:real )?(?:language model|model|LLM)|not (?:a |an )?(?:language model|LLM)/i, 'mock is not presented as real model reasoning')
      assert.equal(await evaluate(`document.querySelector('[data-agent-live-consent]') === null`), true, 'default mock requires no external-provider consent')
      assert.equal(proposalRequests.length, 0, 'loading the mock demo does not request a proposal')
    }
    await evaluate(`Array.from(document.querySelectorAll('button')).find(button => button.textContent.trim() === 'Run agent').click()`)
    await until(`document.querySelector('.agent-demo__status')?.dataset.state === 'awaiting-approval'`, `${tool} preview and validation`, 90_000)
    assert.equal(await evaluate(`Array.from(document.querySelectorAll('button')).find(button => button.textContent.trim() === 'Commit approved change').disabled`), true, 'commit requires explicit approval')
    assert.equal(await evaluate(`document.querySelector('.agent-approval input').checked`), false, 'proposal does not preapprove itself')
    assert.equal(await agentWrites(), 0, `${tool} preview does not write the source`)
    assert.equal(await evaluate(`document.querySelector('[data-agent-download]') === null`), true, `${tool} proposal cannot release a download`)
    assert.match(await evaluate(`document.querySelector('[data-agent-trace]').textContent`), /office.read/, `${tool} discovers targets through actual public reads`)
    if (process.argv.includes('--dev')) assert.equal(proposalRequests.length, proposalsBeforeRun + 1, `${tool} mock uses one local HTTP proposal`)
    if (process.argv.includes('--built')) assert.equal(proposalRequests.length, proposalsBeforeRun, `${tool} static mock needs no endpoint`)
    if (proposalMock) assert.equal(proposalMock.requests.length, 0, `${tool} mock never contacts a live upstream`)
    if (tool === 'sheets') {
      assert.equal(await agentWrites(), 0, 'mock proposal and native preview leave the source untouched')
      assert.match(await evaluate(`document.querySelector('[data-agent-proposal-trace]')?.textContent`), /request/i, 'mock proposal exposes its request trace separately from actual Office tools')
      assert.match(await evaluate(`document.querySelector('[data-agent-proposal-trace]')?.textContent`), /response/i, 'mock proposal exposes its response for inspection')
      assert.equal(proposalRequests.filter((request) => new URL(request.url).pathname.endsWith('/propose')).length, 0, 'default mock never invokes the live proposal relay')
      if (process.argv.includes('--dev')) {
        assert.equal(proposalRequests.length, 1, 'development mock executes through one same-origin endpoint request')
        assert.equal(proposalRequests[0].method, 'POST')
        assert.equal(new URL(proposalRequests[0].url).origin, origin.origin, 'mock context stays on the demo origin')
      } else if (process.argv.includes('--built')) assert.equal(proposalRequests.length, 0, 'static mock transport works without an API server')
      if (proposalMock) assert.equal(proposalMock.requests.length, 0, 'bundled mock never contacts the configured upstream')
      await screenshot('agent-mock-proposal-review')
    }
    await click('.agent-approval input[type=checkbox]')
    await evaluate(`Array.from(document.querySelectorAll('button')).find(button => button.textContent.trim() === 'Commit approved change').click()`)
    await until(`document.querySelector('.agent-demo__status')?.dataset.state === 'verified'`, `${tool} approved commit verifies`, 90_000)
    await until(`!!document.querySelector('[data-agent-download]')`, `${tool} verified bytes released`)
    const fileProof = await evaluate(`(async () => {
      const link = document.querySelector('[data-agent-download]');
      const bytes = new Uint8Array(await (await fetch(link.href)).arrayBuffer());
      const hash = Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256', bytes))).map(value => value.toString(16).padStart(2, '0')).join('');
      const field = Array.from(document.querySelectorAll('.agent-evidence dl > div')).find(item => item.querySelector('dt').textContent === 'Output fingerprint');
      return { name: link.download, length: bytes.length, signature: Array.from(bytes.slice(0, 4)), hash, receiptFingerprint: field.querySelector('dd').textContent };
    })()`)
    assert.ok(fileProof.name.endsWith(`.${fileFormat}`), `${tool} download has the real format extension`)
    assert.ok(fileProof.length > 1000, `${tool} has a populated output file`)
    assert.deepEqual(fileProof.signature, tool === 'pdf' ? [37, 80, 68, 70] : [80, 75, 3, 4], `${tool} output is real file bytes, not JSON`)
    assert.ok(fileProof.receiptFingerprint.endsWith(fileProof.hash), `${tool} download bytes match the verified receipt fingerprint`)
    assert.equal(await agentWrites(), 1, `${tool} approval writes exactly once`)
    await screenshot(`agent-real-${fileFormat}-verified`)
    await evaluate(`document.querySelector('.agent-artifact').scrollIntoView({ block: 'start' })`)
    await screenshot(`agent-real-${fileFormat}-content`)
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
      const originalWrites = await agentWrites()
      assert.equal(originalWrites, 1, 'one approved plan causes exactly one native write')
      await click('[data-agent-retry]')
      await until(`!document.querySelector('[data-agent-retry]').disabled && document.querySelector('.agent-demo__status')?.dataset.state === 'verified'`, 'verified commit retry completes', 90_000)
      assert.equal(await agentWrites(), originalWrites, 'idempotent retry does not duplicate the native write')

      await prepareAgentRequest('Mark Mobile as On track', 'C3')
      await approveAgentCommit()
      await until(`document.querySelector('.agent-demo__status')?.dataset.state === 'verified'`, 'different request writes and verifies the Mobile cell', 90_000)
      assert.match(await evaluate(`document.querySelector('.agent-diff')?.textContent`), /On track/, 'the requested new value is visible in the diff')
      assert.equal(await agentWrites(), 1, 'different request also applies exactly once')

      await clickButton('Reload sample')
      await until(agentReady, 'sample reload before unsupported mock request', 90_000)
      await setAgentRequest('Write a poem about this workbook')
      await clickButton('Run agent')
      await until(`document.querySelector('.agent-demo__status')?.dataset.state === 'error'`, 'unsupported mock prompt is rejected honestly', 90_000)
      assert.match(await evaluate(`document.querySelector('.tool-error[role=alert]')?.textContent`), /mock|support|try|workstream/i, 'mock failure explains its bounded request support')
      assert.equal(await agentWrites(), 0, 'unsupported mock prompt never writes native bytes')
      assert.equal(await evaluate(`document.querySelector('[data-agent-download]') === null && !document.querySelector('.agent-diff')`), true, 'unsupported mock prompt leaves no previous preview or download')
      assert.equal(await evaluate(`document.querySelector('.agent-approval input')?.checked ?? false`), false, 'unsupported mock prompt does not grant approval')

      await prepareAgentRequest('Mark Security as Ready', 'C5')
      await click('[data-agent-concurrent-edit]')
      await until(`document.querySelector('.agent-safety')?.textContent.includes('The source has changed.')`, 'concurrent source edit completes', 90_000)
      const writesBeforeStaleCommit = await agentWrites()
      await approveAgentCommit()
      await until(`document.querySelector('.agent-demo__status')?.dataset.state === 'error'`, 'stale approval fails closed', 90_000)
      assert.match(await evaluate(`document.querySelector('.agent-demo')?.textContent ?? document.querySelector('.app-main').textContent`), /stale|revision/i, 'stale failure explains the revision conflict')
      assert.equal(await agentWrites(), writesBeforeStaleCommit, 'stale commit causes no additional native write')
      assert.equal(await evaluate(`document.querySelector('[data-agent-download]') === null`), true, 'stale approval offers no verified output')

      await prepareAgentRequest('Mark Security as Ready', 'C5')
      await click('[data-agent-fail-verification]')
      await approveAgentCommit()
      await until(`document.querySelector('.agent-demo__status')?.dataset.state === 'unverified'`, 'post-write verification failure has a distinct state', 90_000)
      assert.match(await evaluate(`document.querySelector('.app-main').textContent`), /Write completed; verification failed/, 'UI distinguishes written bytes from verified output')
      assert.equal(await agentWrites(), 1, 'verification failure happened after a completed native write')
      assert.equal(await evaluate(`document.querySelector('[data-agent-download]') === null`), true, 'unverified bytes are not offered as a verified download')
      await screenshot('agent-written-but-unverified')
    }
    else {
      await click('[data-agent-retry]')
      await until(`!document.querySelector('[data-agent-retry]').disabled && document.querySelector('.agent-safety').textContent.includes('No additional native write')`, `${tool} cached retry completes`, 90_000)
      assert.equal(await agentWrites(), 1, `${tool} retry does not write again`)
      const prepareAgain = async () => {
        await clickButton('Reload sample')
        await until(`document.querySelector('.agent-demo__status')?.dataset.state === 'ready' && Array.from(document.querySelectorAll('button')).some(button => button.textContent.trim() === 'Run agent' && !button.disabled)`, `${tool} sample reloaded`, 90_000)
        assert.equal(await evaluate(`document.querySelector('[data-agent-request]').value`), defaultRequest, `${tool} reload restores its real-file request`)
        await clickButton('Run agent')
        await until(`document.querySelector('.agent-demo__status')?.dataset.state === 'awaiting-approval'`, `${tool} fresh proposal reviewed`, 90_000)
      }
      await prepareAgain()
      await click('[data-agent-concurrent-edit]')
      await until(`document.querySelector('.agent-safety').textContent.includes('The source has changed.')`, `${tool} concurrent edit completes`, 90_000)
      const writesBeforeStale = await agentWrites()
      await approveAgentCommit()
      await until(`document.querySelector('.agent-demo__status')?.dataset.state === 'error'`, `${tool} stale approval rejected`, 90_000)
      assert.equal(await agentWrites(), writesBeforeStale, `${tool} stale plan cannot add a source write`)
      assert.equal(await evaluate(`document.querySelector('[data-agent-download]') === null`), true, `${tool} stale plan cannot release output`)
      await prepareAgain()
      await click('[data-agent-fail-verification]')
      await approveAgentCommit()
      await until(`document.querySelector('.agent-demo__status')?.dataset.state === 'unverified'`, `${tool} failed readback distinguished from write`, 90_000)
      assert.equal(await agentWrites(), 1, `${tool} readback fault follows a completed write`)
      assert.equal(await evaluate(`document.querySelector('[data-agent-download]') === null`), true, `${tool} unverified bytes remain unavailable`)
      await screenshot(`agent-${fileFormat}-unverified`)
    }
    await evaluate(`Array.from(document.querySelectorAll('button')).find(button => button.textContent.trim() === 'Refusal proof').click()`)
    await until(`document.querySelector('.agent-demo__status')?.dataset.state === 'ready' && Array.from(document.querySelectorAll('button')).some(button => button.textContent.trim() === 'Run agent' && !button.disabled)`, `${tool} refusal ready`, 90_000)
    await evaluate(`Array.from(document.querySelectorAll('button')).find(button => button.textContent.trim() === 'Run agent').click()`)
    await until(`document.querySelector('.agent-demo__status')?.dataset.state === 'refused'`, `${tool} unsupported operation refused`, 90_000)
    assert.equal(await evaluate(`!document.querySelector('.agent-approval input')`), true, 'refusal does not allow approval')
    assert.equal(await agentWrites(), 0, `${tool} unsupported proposal never reaches the writer`)
    assert.equal(await evaluate(`document.querySelector('[data-agent-download]') === null`), true, 'refusal does not expose a previous output')
    await click('.demo-reset-trigger')
    await until(`document.querySelector('[data-agent-tool=${tool}]') && document.querySelector('.agent-demo__status')?.dataset.state === 'ready' && Array.from(document.querySelectorAll('button')).some(button => button.textContent.trim() === 'Run agent' && !button.disabled)`, `${tool} reset reloads the source`, 90_000)
    assert.equal(await evaluate(`!document.querySelector('[data-agent-download]') && !document.querySelector('.agent-diff') && !Array.from(document.querySelectorAll('.agent-tool-log li[data-state=done]')).some(item => /office\\.(plan|commit)/.test(item.textContent)) && document.querySelector('.agent-approval input').checked === false`), true, 'reset clears outputs, plans, commits, and approval; initial capability discovery is allowed')
  }
  if (proposalMock) {
    await route('agent?format=sheets')
    await until(agentReady, 'mock sample ready for isolated live proposal test', 90_000)
    assert.equal(proposalMock.requests.length, 0, 'default mock workflows never contact the proposal upstream')
    await evaluate(`(() => { const select = document.querySelector('[data-agent-proposal-source]'); select.value = 'live'; select.dispatchEvent(new Event('change', { bubbles: true })); })()`)
    await until(`document.querySelector('[data-agent-live-context] pre') && document.querySelector('[data-agent-live-consent]')`, 'live mode discloses bounded context before consent', 90_000)
    await setAgentRequest('Mark Mobile as On track')
    const disclosedContext = await evaluate(`JSON.parse(document.querySelector('[data-agent-live-context] pre').textContent)`)
    assert.ok(disclosedContext.constraints.allowedTargets.some((target) => target.ref === 'C3' && target.workstream === 'Mobile'), 'disclosure identifies the allowed Mobile cell')
    assert.equal(proposalMock.requests.length, 0, 'selection, request editing, and context discovery do not send anything upstream')
    assert.equal(await evaluate(`document.querySelector('[data-agent-live-consent]').checked`), false, 'data-sharing consent starts unchecked')
    assert.equal(await evaluate(`Array.from(document.querySelectorAll('button')).find(button => button.textContent.trim() === 'Run agent').disabled`), true, 'no live request can run without consent')
    await click('[data-agent-live-consent]')
    await clickButton('Run agent')
    await until(`document.querySelector('.agent-demo__status')?.dataset.state === 'awaiting-approval'`, 'mock live proposal enters human review', 90_000)
    assert.equal(proposalMock.requests.length, 1, 'one consented Run produces exactly one mock upstream request')
    const sent = proposalMock.requests[0]
    assert.equal(sent.hasAuthorization, false, 'no existing provider token is forwarded to the test mock')
    assert.equal(sent.body.request, 'Mark Mobile as On track')
    const { capabilities: disclosedCapabilities, ...disclosedWorkbook } = disclosedContext
    assert.deepEqual(sent.body.context, disclosedWorkbook, 'only the exact disclosed bounded workbook context is sent')
    assert.deepEqual(sent.body.capabilities, disclosedCapabilities, 'only the disclosed capabilities are sent')
    assert.deepEqual(Object.keys(sent.body).sort(), ['capabilities', 'context', 'request'], 'upstream receives proposal context, not file bytes or approval authority')
    assert.match(await evaluate(`document.querySelector('.agent-diff').textContent`), /C3/, 'live proposal resolves the disclosed Mobile target')
    assert.equal(await agentWrites(), 0, 'a live proposal cannot write native bytes')
    assert.equal(await evaluate(`document.querySelector('.agent-approval input').checked`), false, 'upstream confirmation approved cannot grant host approval')
    assert.equal(await evaluate(`Array.from(document.querySelectorAll('button')).find(button => button.textContent.trim() === 'Commit approved change').disabled`), true, 'commit remains disabled until separate human approval')
    assert.equal(await evaluate(`document.querySelector('[data-agent-download]') === null`), true, 'the unapproved proposal cannot expose output bytes')
    await screenshot('agent-live-proposal-review')
    await approveAgentCommit()
    await until(`document.querySelector('.agent-demo__status')?.dataset.state === 'verified' && document.querySelector('[data-agent-download]')`, 'human-approved live proposal produces verified native output', 90_000)
    assert.equal(await agentWrites(), 1, 'the approved live proposal writes exactly once')
    assert.equal(proposalMock.requests.length, 1, 'native commit and verification do not contact the proposal provider')
    const signature = await evaluate(`(async () => Array.from(new Uint8Array(await (await fetch(document.querySelector('[data-agent-download]').href)).arrayBuffer()).slice(0, 4)))()`)
    assert.deepEqual(signature, [80, 75, 3, 4], 'live-proposed verified download is a real XLSX archive')
    await clickButton('Reload sample')
    await until(`document.querySelector('.agent-demo__status')?.dataset.state === 'ready' && !document.querySelector('[data-agent-proposal-source]').disabled`, 'live sample reloaded without automatic proposal', 90_000)
    await evaluate(`(() => { const select = document.querySelector('[data-agent-proposal-source]'); select.value = 'local'; select.dispatchEvent(new Event('change', { bubbles: true })); })()`)
    await until(agentReady, 'local proposal mode restored', 90_000)
    assert.equal(proposalMock.requests.length, 1, 'reset and mode changes do not make hidden upstream requests')
  }
  await click('.source-proof-trigger')
  await click('.guided-recipe__complete')
  await until(`document.querySelector('[role=progressbar]').getAttribute('aria-valuenow') === '1'`, 'AI guide progress recorded')
  await route('agent?format=docs')
  await until(`(!document.querySelector('.source-proof-layer') || document.querySelector('.source-proof-layer').hidden) && document.querySelector('[data-agent-tool=docs]')`, 'same-surface format switch closes stale guide')
  await click('.source-proof-trigger')
  assert.equal(await evaluate(`document.querySelector('[role=progressbar]').getAttribute('aria-valuenow')`), '0', 'new AI format starts its own guide progress')
  await click('.source-proof-close')
  await route('overview')
  await evaluate(`window.scrollTo({ top: 0, behavior: 'instant' })`)
  for (const width of [1200, 1024, 768]) {
    await send('Emulation.setDeviceMetricsOverride', { width, height: 1000, deviceScaleFactor: 1, mobile: false })
    await until(`document.querySelector('.app-main').scrollWidth <= document.querySelector('.app-main').clientWidth`, `catalogue fits width ${width}`, 1000)
  }
  await send('Emulation.setDeviceMetricsOverride', { width: 390, height: 844, deviceScaleFactor: 1, mobile: true })
  await until(`document.documentElement.scrollWidth <= 390`, 'mobile catalogue has no horizontal overflow')
  assert.ok(await evaluate(`document.querySelector('.app-sidebar').getBoundingClientRect().height < 160`), 'mobile navigation does not retain a desktop-height blank area')
  await screenshot('catalogue-mobile')
  await route('agent?format=sheets')
  await until(agentReady, 'mobile AI sample ready', 90_000)
  await clickButton('Run agent')
  await until(`document.querySelector('.agent-demo__status')?.dataset.state === 'awaiting-approval'`, 'mobile AI preview prepared', 90_000)
  await evaluate(`document.querySelector('[data-agent-trace]').open = true`)
  assert.equal(await evaluate(`(() => {
    const workspace = document.querySelector('.agent-demo__workspace').getBoundingClientRect();
    const inspector = document.querySelector('.agent-demo__inspector').getBoundingClientRect();
    const trace = document.querySelector('[data-agent-trace]').getBoundingClientRect();
    const section = document.querySelector('[data-agent-tool]');
    return workspace.bottom <= trace.top + 1 && inspector.bottom <= workspace.bottom + 1 && section.scrollWidth <= section.clientWidth && document.documentElement.scrollWidth <= 390;
  })()`), true, 'mobile AI workbook and inspector fit their section and never overlap the expanded trace')
  assert.equal(await evaluate(`(() => {
    const frame = document.querySelector('.agent-artifact__sheet-scroll');
    const table = frame.querySelector('table');
    frame.scrollLeft = 80;
    return frame.scrollWidth > frame.clientWidth && frame.scrollLeft > 0 && table.getBoundingClientRect().width >= 720;
  })()`), true, 'mobile workbook retains readable columns with keyboard-focusable internal horizontal scrolling')
  assert.equal(await evaluate(`Array.from(document.querySelectorAll('.agent-flight-recorder li')).every(item => item.getBoundingClientRect().width >= 96)`), true, 'mobile workflow labels keep readable internal scroll columns')
  await evaluate(`document.querySelector('.agent-artifact__sheet-scroll').scrollLeft = 0; document.querySelector('.agent-artifact__sheet-scroll').scrollIntoView({ block: 'start' })`)
  await screenshot('agent-mobile-workbook')
  await evaluate(`document.querySelector('[data-agent-trace]').scrollIntoView({ block: 'start' })`)
  await screenshot('agent-mobile-expanded-trace')
  for (const tool of ['docs', 'slides', 'pdf']) {
    await route(`agent?format=${tool}`)
    await until(`document.querySelector('.agent-demo__status')?.dataset.state === 'ready' && Array.from(document.querySelectorAll('button')).some(button => button.textContent.trim() === 'Run agent' && !button.disabled)`, `${tool} mobile sample ready`, 90_000)
    await clickButton('Run agent')
    await until(`document.querySelector('.agent-demo__status')?.dataset.state === 'awaiting-approval'`, `${tool} mobile preview ready`, 90_000)
    await evaluate(`document.querySelector('[data-agent-trace]').open = true`)
    assert.equal(await evaluate(`(() => {
      const workspace = document.querySelector('.agent-demo__workspace').getBoundingClientRect();
      const inspector = document.querySelector('.agent-demo__inspector').getBoundingClientRect();
      const trace = document.querySelector('[data-agent-trace]').getBoundingClientRect();
      const section = document.querySelector('[data-agent-tool]');
      return workspace.bottom <= trace.top + 1 && inspector.bottom <= workspace.bottom + 1 && section.scrollWidth <= section.clientWidth && document.documentElement.scrollWidth <= 390;
    })()`), true, `${tool} mobile content has no horizontal overflow or trace overlap`)
    await evaluate(`document.querySelector('.agent-artifact').scrollIntoView({ block: 'start' })`)
    await screenshot(`agent-mobile-${tool}-content`)
    await evaluate(`document.querySelector('[data-agent-trace]').scrollIntoView({ block: 'start' })`)
    await screenshot(`agent-mobile-${tool}-trace`)
  }
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
  console.log(JSON.stringify({ status: 'passed', mode: process.argv.includes('--dev') ? 'development' : process.argv.includes('--built') ? 'built' : 'existing-server', screenshots: output, checks: ['19 distinct examples', 'search', 'filter return persistence', 'empty recovery', 'continuous section navigation', 'cold-section isolation', 'modal focus/inert', 'checklist persistence', 'source loading', 'all 16 surfaces', 'four AI approvals and refusals', 'default zero-configuration mock with honest labels and no upstream calls', 'mock transport and unsupported-prompt recovery', 'editable agent requests and public tool trace', 'idempotent native commit retry', 'stale approval refusal', 'post-write verification failure', 'AI proof boundaries and resets', 'AI format-specific guides', 'numeric chart source and reset', 'same-surface deep link', 'mobile layout', 'dark theme'], errors }, null, 2))
} catch (error) {
  if (socket?.readyState === WebSocket.OPEN) {
    await screenshot('failure')
    console.error(JSON.stringify({ screenshots: output, errors, state: await evaluate(`({ hash: location.hash, canvasCount: document.querySelectorAll('canvas').length, chartCount: window.__injoffice?.charts?.list().length, text: document.querySelector('.app-main')?.innerText.slice(0, 3500) })`) }, null, 2))
  }
  throw error
} finally {
  for (const task of pending.values()) clearTimeout(task.timer)
  socket?.close()
  try { if (chrome) await terminateProcess(chrome.child) }
  finally {
    try { await staticServer?.close() }
    finally { try { await proposalMock?.close() } finally { restoreProposalEnv?.() } }
  }
}
