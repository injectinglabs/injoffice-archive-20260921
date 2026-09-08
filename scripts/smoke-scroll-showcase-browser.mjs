import assert from 'node:assert/strict'
import { existsSync, mkdirSync, mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { resolve } from 'node:path'
import { launchChromeForCDP, terminateProcess } from './chrome-cdp-startup.mjs'
import { startShowcaseServer } from './showcase-smoke-server.mjs'
import { startShowcaseDevServer } from './showcase-smoke-dev-server.mjs'

// This suite exercises the document navigator, not the individual editors. The
// existing showcase smoke remains responsible for real file and tool proofs.
const keys = ['overview', 'agent-sheets', 'agent-docs', 'agent-slides', 'agent-pdf', 'sheets', 'docs', 'slides', 'pdf', 'charts', 'pivots', 'shapes', 'connectors', 'formulas', 'collab', 'history', 'font-metrics', 'pptx-authored', 'pptx-native', 'pptx-render']
const href = (key) => key.startsWith('agent-') ? `#/agent?format=${key.slice(6)}` : `#/${key}`
const section = (key) => `[data-scroll-section="${key}"]`
const output = process.env.SHOWCASE_OUTPUT ? resolve(process.env.SHOWCASE_OUTPUT) : mkdtempSync(resolve(tmpdir(), 'injoffice-scroll-showcase-'))
mkdirSync(output, { recursive: true })
const pending = new Map()
const errors = []
const liveProposals = []
let sequence = 0
let chrome, socket, server
let chunkFailures = 0
const reloadDialogs = []
let allowReloadDialog = false
let holdChunks = false
const heldChunks = []

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
  } else if (message.method === 'Network.requestWillBeSent' && /\/api\/agent\/propose(?:\?|$)/.test(message.params.request.url)) {
    liveProposals.push(message.params.request.url)
  } else if (message.method === 'Fetch.requestPaused') {
    if (holdChunks) {
      heldChunks.push(message.params.requestId)
    } else if (chunkFailures === 0) {
      chunkFailures++
      void send('Fetch.failRequest', { requestId: message.params.requestId, errorReason: 'ConnectionReset' })
    } else {
      void send('Fetch.continueRequest', { requestId: message.params.requestId })
    }
  } else if (message.method === 'Page.javascriptDialogOpening') {
    reloadDialogs.push(message.params)
    void send('Page.handleJavaScriptDialog', { accept: allowReloadDialog && message.params.type === 'confirm' && /unsaved.*lost/i.test(message.params.message) })
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

async function screenshot(name) {
  const { data } = await send('Page.captureScreenshot', { format: 'png' })
  writeFileSync(resolve(output, `${name}.png`), Buffer.from(data, 'base64'))
}

const active = (key, hash = href(key)) => `location.hash === ${JSON.stringify(hash)} && document.querySelector('.app-sidebar a[aria-current="location"]')?.getAttribute('href') === ${JSON.stringify(href(key))}`
const ready = (key) => `document.querySelector(${JSON.stringify(section(key))})?.dataset.scrollState === 'ready'`
const agentState = (key, state) => `document.querySelector(${JSON.stringify(`${section(key)} .agent-demo__status`)})?.dataset.state === ${JSON.stringify(state)}${state === 'ready' ? ` && Array.from(document.querySelector(${JSON.stringify(section(key))}).querySelectorAll('button')).some(button => button.textContent.trim() === 'Run agent' && !button.disabled)` : ''}`

async function anchor(key) {
  await evaluate(`document.querySelector(${JSON.stringify(`.app-sidebar a[href="${href(key)}"]`)}).click()`)
  await until(active(key), `${key} anchor updates the current location`)
  if (key !== 'overview') await until(ready(key), `${key} lazy section ready`, 90_000)
}

async function button(key, text) {
  await evaluate(`Array.from(document.querySelector(${JSON.stringify(section(key))}).querySelectorAll('button')).find(button => button.textContent.trim() === ${JSON.stringify(text)}).click()`)
}

async function wheelTo(key, hash = href(key)) {
  // Native wheel input does not run an anchor's navigation handler. The target
  // offset comes from the real rendered document, so lazy section heights can vary.
  const distance = await evaluate(`(() => {
    const target = document.querySelector(${JSON.stringify(section(key))});
    const margin = parseFloat(getComputedStyle(target).scrollMarginTop) || 100;
    return target.getBoundingClientRect().top - margin;
  })()`)
  // Exercise wheel input over the visible demo content, not a special gutter.
  // Agent workbenches should grow naturally within the scrolling document.
  await send('Input.dispatchMouseEvent', { type: 'mouseWheel', x: 1100, y: 450, deltaX: 0, deltaY: distance + 8 })
  await until(active(key, hash), `${key} passive wheel scroll updates location`, 90_000)
  if (key !== 'overview') await until(ready(key), `${key} loads after entering the viewport`, 90_000)
}

try {
  assert.ok(!(process.argv.includes('--built') && process.argv.includes('--dev')), 'choose either --built or --dev')
  if (process.argv.includes('--built')) server = await startShowcaseServer(resolve(import.meta.dirname, '../apps/playground/dist'))
  if (process.argv.includes('--dev')) server = await startShowcaseDevServer(resolve(import.meta.dirname, '../apps/playground'))
  const origin = new URL(server?.url ?? process.env.SHOWCASE_URL ?? 'http://127.0.0.1:3100/')
  origin.hash = '#/overview'
  chrome = await launchChromeForCDP({
    executable: process.env.CHROME_BIN ?? ['/Applications/Google Chrome.app/Contents/MacOS/Google Chrome', '/usr/bin/google-chrome', '/usr/bin/chromium', '/usr/bin/chromium-browser'].find(existsSync) ?? 'google-chrome',
    createProfile: () => mkdtempSync(resolve(tmpdir(), 'injoffice-scroll-chrome-')),
  })
  socket = new WebSocket(chrome.target.webSocketDebuggerUrl)
  await new Promise((resolve, reject) => { socket.addEventListener('open', resolve, { once: true }); socket.addEventListener('error', reject, { once: true }) })
  socket.addEventListener('message', onMessage)
  await send('Runtime.enable')
  await send('Page.enable')
  await send('Network.enable')
  await send('Emulation.setDeviceMetricsOverride', { width: 1440, height: 1000, deviceScaleFactor: 1, mobile: false })
  await send('Page.addScriptToEvaluateOnNewDocument', { source: `
    window.__scrollHistory = { pushes: 0, replacements: 0 };
    for (const [method, counter] of [['pushState', 'pushes'], ['replaceState', 'replacements']]) {
      const original = history[method];
      history[method] = function (...args) { window.__scrollHistory[counter]++; return original.apply(this, args); };
    }
  ` })
  await send('Page.navigate', { url: origin.href })
  await until(`document.querySelector('.app-shell')?.dataset.navigation === 'scroll' && document.querySelectorAll('[data-scroll-section]').length === ${keys.length}`, 'continuous showcase ready')
  assert.deepEqual(await evaluate(`Array.from(document.querySelectorAll('[data-scroll-section]')).map(element => element.dataset.scrollSection)`), keys, 'all sections share one document in navigation order')
  assert.equal(await evaluate(`Array.from(document.querySelectorAll('[data-scroll-section]')).every(element => element.querySelector('h1,h2'))`), true, 'every section has a discoverable heading before its editor loads')
  assert.equal(await evaluate(`document.querySelectorAll('[data-scroll-state="ready"]').length < ${keys.length - 1}`), true, 'initial overview does not eagerly initialize every editor')
  await until(active('overview'), 'overview is the current sidebar location')
  await screenshot('scroll-overview-desktop')

  if (process.argv.includes('--fault')) {
    await send('Fetch.enable', { patterns: [{ urlPattern: '*PptxRenderPage*', resourceType: 'Script', requestStage: 'Request' }] })
    await evaluate(`document.querySelector('.app-sidebar a[href="#/pptx-render"]').click()`)
    await until(`document.querySelector('${section('pptx-render')}')?.dataset.scrollState === 'error'`, 'failed chunk is isolated to its section')
    assert.equal(chunkFailures, 1, 'the test deliberately failed exactly one lazy chunk')
    assert.equal(await evaluate(`document.querySelectorAll('[data-scroll-section]').length === 20 && !!document.querySelector('${section('pptx-render')} [role=alert]')`), true, 'one failed editor leaves the document and navigation intact')
    await button('pptx-render', 'Retry')
    await until(`['ready', 'error'].includes(document.querySelector('${section('pptx-render')}')?.dataset.scrollState)`, 'retry settles without breaking other sections', 30_000)
    if (!await evaluate(ready('pptx-render'))) {
      // Chromium can cache a rejected module URL for the document lifetime.
      // The explicit fallback must warn before losing other edited sections.
      allowReloadDialog = true
      await evaluate(`void setTimeout(() => Array.from(document.querySelector('${section('pptx-render')}').querySelectorAll('button')).find(button => button.textContent.trim() === 'Reload page').click(), 0)`)
      await until(`${ready('pptx-render')} && ${active('pptx-render')}`, 'confirmed reload recovers the cached chunk failure at the same deep link', 90_000)
      allowReloadDialog = false
      assert.equal(reloadDialogs.length, 1, 'fallback requires one explicit confirmation')
      assert.match(reloadDialogs[0].message, /unsaved.*lost/i, 'reload warns about losing uncommitted demo edits')
    }
    await send('Fetch.disable')
    await anchor('overview')
  }

  // Cold navigation preserves the overview DOM and uses a real section anchor.
  await evaluate(`void (window.__overviewNode = document.querySelector('[data-scroll-section="overview"]'))`)
  await anchor('agent-docs')
  await until(agentState('agent-docs', 'ready'), 'DOCX agent initialized', 90_000)
  assert.equal(await evaluate(`window.__overviewNode.isConnected && window.__overviewNode === document.querySelector('[data-scroll-section="overview"]')`), true, 'anchor navigation preserves existing sections')
  const prompt = 'Replace "Northstar Launch Brief" with "Northstar Scrolling Review"'
  await evaluate(`(() => {
    const input = document.querySelector('${section('agent-docs')} [data-agent-request]');
    Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set.call(input, ${JSON.stringify(prompt)});
    input.dispatchEvent(new Event('input', { bubbles: true }));
    input.focus();
    window.__persistedPrompt = input;
    window.__beforePassiveHistory = { ...window.__scrollHistory, length: history.length };
  })()`)
  await wheelTo('agent-slides')
  assert.equal(await evaluate(`document.activeElement === window.__persistedPrompt`), true, 'passive scrolling does not steal focus from an editor')
  assert.equal(await evaluate(`history.length === window.__beforePassiveHistory.length && window.__scrollHistory.pushes === window.__beforePassiveHistory.pushes && window.__scrollHistory.replacements > window.__beforePassiveHistory.replacements`), true, 'scroll spy replaces the URL without adding browser history entries')
  assert.equal(await evaluate(`document.querySelectorAll('.app-sidebar a[aria-current="location"]').length`), 1, 'exactly one sidebar link is current')
  await anchor('agent-docs')
  assert.equal(await evaluate(`window.__persistedPrompt.isConnected && document.querySelector('${section('agent-docs')} [data-agent-request]').value === ${JSON.stringify(prompt)}`), true, 'scrolling away and back preserves typed editor state')
  await button('agent-docs', 'Run agent')
  await until(agentState('agent-docs', 'awaiting-approval'), 'real DOCX preview ready', 90_000)
  const plan = await evaluate(`document.querySelector('${section('agent-docs')} .agent-diff').textContent`)
  await anchor('charts')
  await until(`document.querySelector('${section('charts')} input[aria-label="Jan revenue"]')`, 'editable chart mounted')
  await evaluate(`(() => {
    const input = document.querySelector('${section('charts')} input[aria-label="Jan revenue"]');
    Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(input, '999');
    input.dispatchEvent(new Event('input', { bubbles: true }));
    window.__persistedChartInput = input;
  })()`)
  await until(`document.querySelector('${section('charts')} .tool-metrics').textContent.includes('999')`, 'chart analysis reflects the real source edit')
  await anchor('agent-docs')
  assert.equal(await evaluate(agentState('agent-docs', 'awaiting-approval')), true, 'returning to a section preserves its pending approval')
  assert.equal(await evaluate(`document.querySelector('${section('agent-docs')} .agent-diff').textContent`), plan, 'the exact reviewed change survives navigation')
  assert.equal(await evaluate(`Number(document.querySelector('${section('agent-docs')} [data-agent-native-writes]').textContent)`), 0, 'navigation never approves or writes a proposed edit')
  assert.equal(await evaluate(`document.querySelector('${section('agent-docs')} .agent-approval input').checked`), false, 'pending approval remains unchecked')

  // Re-selecting the current hash still returns to the section heading.
  await evaluate(`window.scrollBy({ top: 320, behavior: 'instant' })`)
  await anchor('agent-docs')
  await until(`(() => { const element = document.querySelector('${section('agent-docs')}'); const top = element.getBoundingClientRect().top; const margin = parseFloat(getComputedStyle(element).scrollMarginTop) || 100; return Math.abs(top - margin) < 40; })()`, 'same anchor returns to its section heading')
  await screenshot('scroll-pending-approval-desktop')

  // Explicit anchors add navigable entries; scroll updates alone do not.
  await anchor('charts')
  assert.equal(await evaluate(`window.__persistedChartInput.isConnected && document.querySelector('${section('charts')} input[aria-label="Jan revenue"]').value === '999'`), true, 'chart source edits survive navigating to another live editor')
  await anchor('shapes')
  await evaluate('history.back()')
  await until(active('charts'), 'browser Back restores the previous section')
  await evaluate('history.forward()')
  await until(active('shapes'), 'browser Forward restores the next section')
  await wheelTo('connectors')
  await wheelTo('formulas')
  await screenshot('scroll-active-sidebar-desktop')

  // A new explicit anchor must win over scroll frames queued by the previous
  // navigation. These warm neighbors reproduce rapid section changes.
  for (let index = 0; index < 3; index++) {
    await anchor('charts')
    await anchor('shapes')
    await anchor('pivots')
    await until(active('pivots'), 'rapid anchors settle on the last explicit destination')
  }

  // Preserve query intent even when a cold section first mounts after the user
  // has already moved elsewhere. Re-enter via passive scrolling, not a hashchange.
  origin.hash = '#/overview'
  origin.searchParams.set('scroll-smoke-document', 'cold-sheets')
  await send('Page.navigate', { url: origin.href })
  await until(active('overview'), 'fresh document ready for cold Sheets deep link')
  holdChunks = true
  await send('Fetch.enable', { patterns: [{ urlPattern: '*SheetsPage*', resourceType: 'Script', requestStage: 'Request' }] })
  await evaluate(`location.hash = '#/sheets?view=native'`)
  const coldStart = Date.now()
  while (heldChunks.length === 0 && Date.now() - coldStart < 10_000) await new Promise(resolve => setTimeout(resolve, 50))
  assert.ok(heldChunks.length > 0, 'cold native Sheets deep link waits for its lazy chunk')
  await anchor('history')
  holdChunks = false
  await send('Fetch.disable')
  await until(`document.querySelector('${section('sheets')} .native-toolbar')`, 'cold Sheets mounts its remembered native view while another section owns the URL', 90_000)
  await wheelTo('sheets', '#/sheets?view=native')
  assert.equal(await evaluate(`document.querySelector('${section('sheets')} .native-toolbar') !== null`), true, 'passive re-entry preserves both the native view and its original deep link')

  // Reloading a shared deep link chooses the exact AI format, not the first
  // mounted agent. Neighboring agents retain their own independent format.
  origin.hash = '#/agent?format=pdf'
  origin.searchParams.set('scroll-smoke-document', 'pdf-deep-link')
  await send('Page.navigate', { url: origin.href })
  await until(active('agent-pdf'), 'direct AI PDF deep link selects its section', 90_000)
  await until(agentState('agent-pdf', 'ready'), 'deep-linked PDF agent is ready', 90_000)
  assert.equal(await evaluate(`document.querySelector('${section('agent-pdf')} [data-agent-tool]')?.dataset.agentTool`), 'pdf', 'deep link initializes the requested agent format')
  await anchor('agent-docs')
  await until(agentState('agent-docs', 'ready'), 'neighbor DOCX agent is ready', 90_000)
  assert.equal(await evaluate(`document.querySelector('${section('agent-pdf')} [data-agent-tool]')?.dataset.agentTool`), 'pdf', 'changing the current section does not retarget an already mounted agent')
  assert.equal(await evaluate(`(() => {
    const prompts = Array.from(document.querySelectorAll('[data-agent-request]'));
    return prompts.length >= 2 && new Set(prompts.map(input => input.id)).size === prompts.length && prompts.every(input => Array.from(input.labels ?? []).some(label => label.closest('[data-scroll-section]') === input.closest('[data-scroll-section]')));
  })()`), true, 'mounted agents have unique prompt ids with labels belonging to their own section')

  await send('Emulation.setDeviceMetricsOverride', { width: 390, height: 844, deviceScaleFactor: 1, mobile: true })
  await anchor('agent-pdf')
  assert.equal(await evaluate(`document.documentElement.scrollWidth <= 390`), true, 'mobile document has no horizontal overflow')
  assert.equal(await evaluate(`(() => { const navigator = document.querySelector('.app-sidebar'); const bounds = navigator.getBoundingClientRect(); return bounds.top >= -1 && bounds.top < 200 && bounds.bottom > 0 && bounds.bottom < innerHeight; })()`), true, 'mobile section navigator stays in the viewport')
  assert.equal(await evaluate(`(() => { const link = document.querySelector('.app-sidebar a[aria-current="location"]'); const bounds = link.getBoundingClientRect(); return bounds.left >= -1 && bounds.right <= innerWidth + 1 && bounds.top >= -1 && bounds.bottom <= innerHeight; })()`), true, 'the current mobile section link is visible')
  await screenshot('scroll-navigation-mobile')
  await anchor('charts')
  assert.equal(await evaluate(`document.documentElement.scrollWidth <= 390`), true, 'later mounted editor does not overflow the mobile document')
  await screenshot('scroll-chart-mobile')
  assert.deepEqual(liveProposals, [], 'scroll demo never calls a real model endpoint')
  assert.deepEqual(errors, [], 'no uncaught errors or console errors')
  console.log(JSON.stringify({ status: 'passed', mode: process.argv.includes('--dev') ? 'development' : process.argv.includes('--built') ? 'built' : 'existing-server', screenshots: output, checks: ['20 continuous sections', 'lazy editor initialization', 'sidebar scroll spy', 'passive scroll replaces history', 'passive scroll preserves focus', 'persistent prompt and pending approval', 'same anchor returns to heading', 'Back and Forward', 'shared format-specific deep links', 'independent mounted agent formats', 'sticky mobile navigator', 'mobile overflow', 'no real model calls'], errors }, null, 2))
} catch (error) {
  if (socket?.readyState === WebSocket.OPEN) {
    try {
      await screenshot('failure')
      console.error(JSON.stringify({ screenshots: output, errors, state: await evaluate(`({ hash: location.hash, y: scrollY, active: document.querySelector('.app-sidebar a[aria-current]')?.outerHTML, sections: Array.from(document.querySelectorAll('[data-scroll-section]')).map(element => ({ key: element.dataset.scrollSection, state: element.dataset.scrollState, top: Math.round(element.getBoundingClientRect().top), agentState: element.querySelector('.agent-demo__status')?.dataset.state, prompt: element.querySelector('[data-agent-request]')?.value, error: element.querySelector('.tool-error')?.textContent })) })`) }, null, 2))
    } catch (diagnosticError) { console.error(`Could not collect failure diagnostics: ${diagnosticError.message}`) }
  }
  throw error
} finally {
  for (const task of pending.values()) clearTimeout(task.timer)
  socket?.close()
  try { if (chrome) await terminateProcess(chrome.child) }
  finally { await server?.close() }
}
