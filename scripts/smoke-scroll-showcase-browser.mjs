import assert from 'node:assert/strict'
import { existsSync, mkdirSync, mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { resolve } from 'node:path'
import { launchChromeForCDP, terminateProcess } from './chrome-cdp-startup.mjs'
import { startShowcaseServer } from './showcase-smoke-server.mjs'
import { startShowcaseDevServer } from './showcase-smoke-dev-server.mjs'

// This suite exercises the document navigator, not the individual editors. The
// existing showcase smoke remains responsible for real file and tool proofs.
const keys = ['sheets', 'docs', 'slides', 'pdf']
const rememberedFeatures = { sheets: 'agent', docs: 'agent', slides: 'agent', pdf: 'agent' }
const target = key => key.startsWith('agent-') ? { tool: key.slice(6), feature: 'agent' }
  : keys.includes(key) ? { tool: key, feature: rememberedFeatures[key] }
    : { tool: key.startsWith('pptx-') ? 'slides' : key === 'font-metrics' ? 'docs' : 'sheets', feature: key }
const href = key => { const { tool, feature } = target(key); return `#/${tool}?feature=${feature}` }
const toolSection = key => `[data-scroll-section="${target(key).tool}"]`
const section = key => keys.includes(key) ? toolSection(key) : `${toolSection(key)} [data-workspace-panel="${target(key).feature}"]`
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
const expectedChunkErrors = []
function recordError(message) {
  // The intentional module failure is surfaced by React's feature boundary.
  // Only that exact failed feature is expected; unrelated errors still fail.
  if (process.argv.includes('--fault') && chunkFailures === 1 && /PptxRenderPage/.test(message) && /Failed to fetch dynamically imported module|Importing a module script failed|error while loading dynamically imported module/i.test(message)) expectedChunkErrors.push(message)
  else errors.push(message)
}

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
    recordError(message.params.exceptionDetails.exception?.description ?? message.params.exceptionDetails.text)
  } else if (message.method === 'Runtime.consoleAPICalled' && message.params.type === 'error') {
    recordError(message.params.args.map((arg) => arg.value ?? arg.description).join(' '))
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
    void send('Page.handleJavaScriptDialog', { accept: message.params.type === 'confirm' && allowReloadDialog && /unsaved.*lost/i.test(message.params.message) })
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

async function assertNavigationSelection(viewport) {
  const originalTheme = await evaluate(`document.documentElement.getAttribute('data-theme')`)
  try {
    for (const theme of ['light', 'dark']) {
      await evaluate(`document.documentElement.setAttribute('data-theme', ${JSON.stringify(theme)})`)
      const style = await evaluate(`(() => {
        const selected = document.querySelector('.app-sidebar a[aria-current="location"]');
        const unselected = document.querySelector('.app-sidebar a:not([aria-current])');
        const current = getComputedStyle(selected);
        return {
          boxShadow: current.boxShadow,
          borders: ['Top', 'Right', 'Bottom', 'Left'].map(side => current['border' + side + 'Width']),
          background: current.backgroundColor,
          unselectedBackground: getComputedStyle(unselected).backgroundColor,
          fontWeight: Number(current.fontWeight),
        };
      })()`)
      const label = `${viewport} ${theme} selected navigation`
      assert.equal(style.boxShadow, 'none', `${label} has no inset edge highlight`)
      assert.deepEqual(style.borders, ['0px', '0px', '0px', '0px'], `${label} has no border highlight`)
      assert.notEqual(style.background, style.unselectedBackground, `${label} uses a distinct background`)
      assert.notEqual(style.background, 'rgba(0, 0, 0, 0)', `${label} has a visible fill`)
      assert.ok(style.fontWeight >= 600, `${label} emphasizes the selected label`)

      // Exercise real keyboard navigation back to the selected link so that
      // removing its selection stripe cannot silently remove its focus ring.
      await new Promise(resolve => setTimeout(resolve, 750))
      await evaluate(`window.__keyboardNavigationLink = document.querySelector('.app-sidebar a[aria-current="location"]'); window.__keyboardNavigationLink.focus({ preventScroll: true })`)
      for (const modifiers of [8, 0]) {
        await send('Input.dispatchKeyEvent', { type: 'keyDown', key: 'Tab', code: 'Tab', windowsVirtualKeyCode: 9, modifiers })
        await send('Input.dispatchKeyEvent', { type: 'keyUp', key: 'Tab', code: 'Tab', windowsVirtualKeyCode: 9, modifiers })
      }
      const focusProof = await evaluate(`(() => {
        const selected = window.__keyboardNavigationLink;
        const style = getComputedStyle(selected);
        return { selected: selected.outerHTML, active: document.activeElement?.outerHTML.slice(0, 500),
          focused: document.activeElement === selected, visible: selected.matches(':focus-visible'),
          width: parseFloat(style.outlineWidth), style: style.outlineStyle, color: style.outlineColor };
      })()`)
      assert.ok(focusProof.focused && focusProof.visible && focusProof.width >= 2 && focusProof.style !== 'none' && focusProof.color !== 'rgba(0, 0, 0, 0)', `${label} retains a visible keyboard focus outline: ${JSON.stringify(focusProof)}`)
      await screenshot(`scroll-selection-${viewport}-${theme}-keyboard`)
      await evaluate(`document.activeElement.blur()`)
    }
  } finally {
    await evaluate(originalTheme === null
      ? `document.documentElement.removeAttribute('data-theme')`
      : `document.documentElement.setAttribute('data-theme', ${JSON.stringify(originalTheme)})`)
  }
}

const active = (key, hash = href(key)) => `(location.hash === ${JSON.stringify(hash)} || location.hash === ${JSON.stringify(href(key))} || (${keys.includes(key)} && location.hash === ${JSON.stringify(`#/${target(key).tool}`)})) && document.querySelector('.app-sidebar a[aria-current="location"]')?.getAttribute('href') === ${JSON.stringify(`#/${target(key).tool}`)}`
const ready = key => `document.querySelector('${toolSection(key)}')?.dataset.scrollState === 'ready' && !!document.querySelector('${toolSection(key)} [data-workspace-panel="${target(key).feature}"]:not([hidden])') && !document.querySelector('${toolSection(key)} [data-workspace-panel="${target(key).feature}"] [data-workspace-loading]') && !document.querySelector('${toolSection(key)} [data-workspace-panel="${target(key).feature}"] [data-workspace-error]')`
const agentState = (key, state) => `document.querySelector(${JSON.stringify(`${section(key)} .agent-demo__status`)})?.dataset.state === ${JSON.stringify(state)}${state === 'ready' ? ` && Array.from(document.querySelector(${JSON.stringify(section(key))}).querySelectorAll('button')).some(button => button.textContent.trim() === 'Preview change' && !button.disabled)` : ''}`

async function anchor(key) {
  const { tool, feature } = target(key)
  if (keys.includes(key)) await evaluate(`document.querySelector('.app-sidebar a[href="#/${tool}"]').click()`)
  else {
    rememberedFeatures[tool] = feature
    await evaluate(`location.hash = ${JSON.stringify(href(key))}`)
  }
  await until(active(key), `${key} anchor updates the current location`)
  await until(ready(key), `${key} lazy section ready`, 90_000)
  // A warmed assistant can be ready before the hash handler's focus frame.
  // Finish explicit navigation before testing subsequent passive interaction.
  await evaluate(`new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)))`)
}

async function button(key, text) {
  await evaluate(`Array.from(document.querySelector(${JSON.stringify(section(key))}).querySelectorAll('button')).find(button => button.textContent.trim() === ${JSON.stringify(text)}).click()`)
}

async function assertInternalGroupNavigation(group, feature, key) {
  const navigation = '.app-sidebar [aria-label="Docs examples"]'
  const selector = `${navigation} [data-workspace-feature="${feature}"]`
  await evaluate(`(() => {
    const button = document.querySelector('${selector}');
    button.scrollIntoView({ block: 'center' });
    button.focus({ preventScroll: true });
  })()`)
  await evaluate(`new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)))`)
  const point = await evaluate(`(() => { const r = document.querySelector('${selector}').getBoundingClientRect(); return { x: r.left + r.width / 2, y: r.top + r.height / 2 } })()`)
  if (key) {
    await send('Input.dispatchKeyEvent', { type: 'keyDown', key: 'Enter', code: 'Enter', windowsVirtualKeyCode: 13, text: '\r' })
    await send('Input.dispatchKeyEvent', { type: 'keyUp', key: 'Enter', code: 'Enter', windowsVirtualKeyCode: 13 })
  } else {
    await send('Input.dispatchMouseEvent', { type: 'mousePressed', button: 'left', clickCount: 1, ...point })
    await send('Input.dispatchMouseEvent', { type: 'mouseReleased', button: 'left', clickCount: 1, ...point })
  }
  rememberedFeatures.docs = feature
  await until(`${ready('docs')} && document.querySelector('${selector}').getAttribute('aria-current') === 'true'`, `${key ? 'keyboard' : 'pointer'} selects ${feature}`)
  await until(`document.activeElement === document.querySelector('#example-docs-${feature} h3')`, `sidebar selection focuses the example heading (was ${await evaluate('document.activeElement?.outerHTML.slice(0, 300)')})`)
  await evaluate(`new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)))`)
  assert.equal(await evaluate(`document.querySelector('.app-sidebar a[aria-current="location"]').getAttribute('href')`), '#/docs', 'internal navigation keeps Docs selected')
}
async function wheelTo(key, hash = href(key)) {
  // Native wheel input does not run an anchor's navigation handler. The target
  // offset comes from the real rendered document, so lazy section heights can vary.
  const distance = await evaluate(`(() => {
    const target = document.querySelector(${JSON.stringify(`${toolSection(key)} [data-workspace-panel="${target(key).feature}"]`)}) ?? document.querySelector(${JSON.stringify(toolSection(key))});
    const margin = parseFloat(getComputedStyle(target).scrollMarginTop) || 100;
    return target.getBoundingClientRect().top - margin;
  })()`)
  // Exercise wheel input over the visible demo content, not a special gutter.
  // Agent workbenches should grow naturally within the scrolling document.
  await send('Input.dispatchMouseEvent', { type: 'mouseWheel', x: 1100, y: 450, deltaX: 0, deltaY: distance + 8 })
  await until(active(key, hash), `${key} passive wheel scroll updates location`, 90_000)
  await until(ready(key), `${key} loads after entering the viewport`, 90_000)
}

try {
  assert.ok(!(process.argv.includes('--built') && process.argv.includes('--dev')), 'choose either --built or --dev')
  if (process.argv.includes('--built')) server = await startShowcaseServer(resolve(import.meta.dirname, '../apps/playground/dist'))
  if (process.argv.includes('--dev')) server = await startShowcaseDevServer(resolve(import.meta.dirname, '../apps/playground'))
  const origin = new URL(server?.url ?? process.env.SHOWCASE_URL ?? 'http://127.0.0.1:3100/')
  origin.hash = '#/docs?feature=font-metrics'
  rememberedFeatures.docs = 'font-metrics'
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
  assert.equal(await evaluate(`document.querySelectorAll('[data-scroll-state="ready"]').length < ${keys.length}`), true, 'direct Docs start does not eagerly initialize every editor')
  await until(active('font-metrics'), 'direct Docs feature is the current sidebar location')
  await until(ready('font-metrics'), 'direct Docs feature is ready', 90_000)
  assert.equal(await evaluate(`!document.querySelector('.overview-page, [data-workspace-entry]') && !Array.from(document.querySelectorAll('.app-main a')).some(link => link.textContent.trim() === 'All demos')`), true, 'there is no introductory page or obsolete All demos link')
  assert.deepEqual(await evaluate(`Array.from(document.querySelectorAll('.app-sidebar .tool-nav-title')).map(link => link.getAttribute('href'))`), ['#/sheets', '#/docs', '#/slides', '#/pdf'], 'examples are grouped under the four tools')
  assert.equal(await evaluate(`document.querySelectorAll('.app-sidebar [data-workspace-feature]').length`), 28, 'all examples are visible in the sidebar index')
  await screenshot('scroll-docs-direct-desktop')
  await assertNavigationSelection('desktop')

  if (process.argv.includes('--fault')) {
    await send('Fetch.enable', { patterns: [{ urlPattern: '*PptxRenderPage*', resourceType: 'Script', requestStage: 'Request' }] })
    rememberedFeatures.slides = 'pptx-render'
    await evaluate(`location.hash = '#/slides?feature=pptx-render'`)
    await until(`!!document.querySelector('${section('pptx-render')} [data-workspace-error]')`, 'failed chunk is isolated to its feature')
    assert.equal(chunkFailures, 1, 'the test deliberately failed exactly one lazy chunk')
    assert.equal(await evaluate(`document.querySelectorAll('[data-scroll-section]').length === 4 && !!document.querySelector('${section('pptx-render')} [role=alert]')`), true, 'one failed feature leaves all four workspaces and navigation intact')
    await evaluate(`document.querySelector('${section('pptx-render')} [data-workspace-retry]').click()`)
    await until(`!!document.querySelector('${section('pptx-render')} [data-workspace-error]') || (${ready('pptx-render')})`, 'retry settles without breaking other features', 30_000)
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
    await anchor('font-metrics')
  }

  // Feature navigation preserves the existing tool shells in the same document.
  await evaluate(`void (window.__sheetsNode = document.querySelector('[data-scroll-section="sheets"]'))`)
  await anchor('agent-docs')
  await until(agentState('agent-docs', 'ready'), 'DOCX agent initialized', 90_000)
  assert.equal(await evaluate(`window.__sheetsNode.isConnected && window.__sheetsNode === document.querySelector('[data-scroll-section="sheets"]')`), true, 'anchor navigation preserves existing sections')
  const prompt = 'Northstar Scrolling Review'
  await evaluate(`(() => {
    const input = document.querySelector('${section('agent-docs')} [data-agent-task-value]');
    Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(input, ${JSON.stringify(prompt)});
    input.dispatchEvent(new Event('input', { bubbles: true }));
    input.focus();
    window.__persistedPrompt = input;
    window.__beforePassiveHistory = { ...window.__scrollHistory, length: history.length };
  })()`)
  await wheelTo('slides')
  assert.equal(await evaluate(`document.activeElement === window.__persistedPrompt`), true, `passive scrolling does not steal focus: ${await evaluate('document.activeElement?.outerHTML.slice(0, 400)')}`)
  assert.equal(await evaluate(`history.length === window.__beforePassiveHistory.length && window.__scrollHistory.pushes === window.__beforePassiveHistory.pushes && window.__scrollHistory.replacements > window.__beforePassiveHistory.replacements`), true, 'scroll spy replaces the URL without adding browser history entries')
  assert.equal(await evaluate(`document.querySelectorAll('.app-sidebar a[aria-current="location"]').length`), 1, 'exactly one sidebar link is current')
  await anchor('agent-docs')
  assert.equal(await evaluate(`window.__persistedPrompt.isConnected && document.querySelector('${section('agent-docs')} [data-agent-task-value]').value === ${JSON.stringify(prompt)}`), true, 'scrolling away and back preserves typed guided task state')
  await button('agent-docs', 'Preview change')
  await until(agentState('agent-docs', 'awaiting-approval'), 'real DOCX preview ready', 90_000)
  const plan = await evaluate(`document.querySelector('${section('agent-docs')} .agent-diff').textContent`)
  rememberedFeatures.sheets = 'editor'
  await evaluate(`location.hash = '#/sheets?feature=editor'`)
  await until(`!!document.querySelector('${toolSection('sheets')} [data-workspace-panel="editor"] canvas')`, 'main workbook rendered before retention check', 90_000)
  await anchor('charts')
  await until(`document.querySelector('${section('charts')} input[aria-label="Jan revenue"]')`, 'editable chart mounted')
  assert.equal(await evaluate(`(() => {
    const panel = document.querySelector('${toolSection('sheets')} [data-workspace-panel="editor"]');
    const rect = panel.getBoundingClientRect();
    return !panel.hidden && !panel.inert && getComputedStyle(panel).visibility === 'visible' && rect.width > 0 && rect.height > 0;
  })()`), true, 'the workbook remains in document flow with a measurable canvas while visiting another example')
  await evaluate(`(() => {
    const input = document.querySelector('${section('charts')} input[aria-label="Jan revenue"]');
    Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(input, '999');
    input.dispatchEvent(new Event('input', { bubbles: true }));
    input.focus();
    window.__persistedChartInput = input;
  })()`)
  await send('Input.dispatchKeyEvent', { type: 'keyDown', key: 'Tab', code: 'Tab', windowsVirtualKeyCode: 9 })
  await send('Input.dispatchKeyEvent', { type: 'keyUp', key: 'Tab', code: 'Tab', windowsVirtualKeyCode: 9 })
  assert.equal(await evaluate(`!!document.activeElement && !document.querySelector('${toolSection('sheets')} [data-workspace-panel="editor"]').contains(document.activeElement) && !document.activeElement.closest('[hidden], [inert]')`), true, 'keyboard Tab never enters a retained hidden workbook')
  await until(`document.querySelector('${section('charts')} .tool-metrics').textContent.includes('999')`, 'chart analysis reflects the real source edit')
  await anchor('agent-docs')
  assert.equal(await evaluate(agentState('agent-docs', 'awaiting-approval')), true, 'returning to a section preserves its pending approval')
  assert.equal(await evaluate(`document.querySelector('${section('agent-docs')} .agent-diff').textContent`), plan, 'the exact reviewed change survives navigation')
  assert.equal(await evaluate(`Number(document.querySelector('${section('agent-docs')} [data-agent-native-writes]').textContent)`), 0, 'navigation never approves or writes a proposed edit')
  assert.equal(await evaluate(`document.querySelector('${section('agent-docs')} .agent-approval input').checked`), false, 'pending approval remains unchecked')
  rememberedFeatures.docs = 'editor'
  await evaluate(`location.hash = '#/docs?feature=editor'`)
  await until(ready('docs'), 'primary Docs editor opens without destroying the agent')
  assert.equal(await evaluate(`document.querySelector('${section('agent-docs')}').hidden`), false, 'the agent example remains in the page alongside the editor')
  await anchor('agent-docs')
  assert.equal(await evaluate(agentState('agent-docs', 'awaiting-approval')), true, 'feature changes preserve the pending proposal')
  assert.equal(await evaluate(`document.querySelector('${section('agent-docs')} .agent-diff').textContent`), plan, 'returning from the native editor preserves the exact agent diff')
  assert.equal(await evaluate(`document.querySelector('${section('agent-docs')} .agent-approval input').checked`), false, 'feature changes never grant approval')
  await anchor('docs')
  await assertInternalGroupNavigation('Edit', 'editor')
  await assertInternalGroupNavigation('AI', 'agent')
  await assertInternalGroupNavigation('Edit', 'editor', 'ArrowLeft')
  await assertInternalGroupNavigation('AI', 'agent', 'ArrowRight')
  assert.equal(await evaluate(`document.querySelector('${section('agent-docs')} .agent-diff').textContent`), plan, 'pointer and keyboard group changes retain the exact reviewed proposal')
  assert.equal(await evaluate(`Number(document.querySelector('${section('agent-docs')} [data-agent-native-writes]').textContent)`), 0, 'internal group navigation never authorizes a document write')

  // Re-selecting the current hash still returns to the example heading.
  await evaluate(`window.scrollBy({ top: 320, behavior: 'instant' })`)
  await anchor('docs')
  await until(`(() => { const element = document.querySelector('${section('agent-docs')}'); const top = element.getBoundingClientRect().top; const margin = parseFloat(getComputedStyle(element).scrollMarginTop) || 100; return Math.abs(top - margin) < 40; })()`, 'same tool anchor returns to its example heading without changing the selected feature')
  await screenshot('scroll-pending-approval-desktop')

  // Explicit anchors add navigable entries; scroll updates alone do not.
  await anchor('charts')
  assert.equal(await evaluate(`window.__persistedChartInput.isConnected && document.querySelector('${section('charts')} input[aria-label="Jan revenue"]').value === '999'`), true, 'chart source edits survive navigating to another live editor')
  await anchor('shapes')
  await evaluate('history.back()')
  await until(active('charts'), 'browser Back restores the previous section')
  await evaluate('history.forward()')
  await until(active('shapes'), 'browser Forward restores the next section')
  await wheelTo('docs')
  await wheelTo('pdf')
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
  origin.hash = '#/docs?feature=font-metrics'
  origin.searchParams.set('scroll-smoke-document', 'cold-sheets')
  await send('Page.navigate', { url: origin.href })
  Object.assign(rememberedFeatures, { sheets: 'agent', docs: 'font-metrics', slides: 'agent', pdf: 'agent' })
  await until(active('font-metrics'), 'fresh Docs document ready for cold Sheets deep link')
  await until(ready('font-metrics'), 'fresh Docs font metrics ready', 90_000)
  holdChunks = true
  await send('Fetch.enable', { patterns: [{ urlPattern: '*NativeRoundTripPage*', resourceType: 'Script', requestStage: 'Request' }] })
  rememberedFeatures.sheets = 'native'
  await evaluate(`location.hash = '#/sheets?feature=native'`)
  const coldStart = Date.now()
  while (heldChunks.length === 0 && Date.now() - coldStart < 10_000) await new Promise(resolve => setTimeout(resolve, 50))
  assert.ok(heldChunks.length > 0, 'cold native Sheets deep link waits for its lazy chunk')
  await anchor('docs')
  holdChunks = false
  await send('Fetch.disable')
  await until(`document.querySelector('${section('sheets')} .native-toolbar')`, 'cold Sheets mounts its remembered native view while another section owns the URL', 90_000)
  await wheelTo('sheets', '#/sheets?feature=native')
  assert.equal(await evaluate(`document.querySelector('${section('sheets')} .native-toolbar') !== null`), true, 'passive re-entry preserves both the native view and its original deep link')
  await button('sheets', 'Use bundled .xlsx')
  await until(`document.querySelector('${section('sheets')} .native-status')?.textContent.includes('Extracted ')`, 'retained native view loads a real workbook', 90_000)
  await evaluate(`(() => {
    const input = document.querySelector('${section('sheets')} [data-workspace-panel="native"] .native-field input');
    Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(input, 'Retained native edit');
    input.dispatchEvent(new Event('input', { bubbles: true }));
    window.__retainedNativeInput = input;
  })()`)
  await anchor('charts')
  await until(`document.querySelector('${section('charts')} input[aria-label="Jan revenue"]')`, 'chart feature opens beside retained native workbook')
  rememberedFeatures.sheets = 'native'
  await evaluate(`location.hash = '#/sheets?feature=native'`)
  await until(ready('sheets'), 'return to native feature')
  assert.equal(await evaluate(`window.__retainedNativeInput.isConnected && window.__retainedNativeInput.value === 'Retained native edit' && !window.__retainedNativeInput.closest('[hidden]')`), true, 'changing feature preserves the exact native editor and pending text')

  // Reloading a shared deep link chooses the exact AI format, not the first
  // mounted agent. Neighboring agents retain their own independent format.
  origin.hash = '#/agent?format=pdf'
  origin.searchParams.set('scroll-smoke-document', 'pdf-deep-link')
  await send('Page.navigate', { url: origin.href })
  Object.assign(rememberedFeatures, { sheets: 'agent', docs: 'agent', slides: 'agent', pdf: 'agent' })
  await until(active('agent-pdf', '#/agent?format=pdf'), 'legacy AI PDF deep link selects its workspace and feature', 90_000)
  await until(agentState('agent-pdf', 'ready'), 'deep-linked PDF agent is ready', 90_000)
  assert.equal(await evaluate(`document.querySelector('${section('agent-pdf')} [data-agent-tool]')?.dataset.agentTool`), 'pdf', 'deep link initializes the requested agent format')
  await anchor('agent-docs')
  await until(agentState('agent-docs', 'ready'), 'neighbor DOCX agent is ready', 90_000)
  assert.equal(await evaluate(`document.querySelector('${section('agent-pdf')} [data-agent-tool]')?.dataset.agentTool`), 'pdf', 'changing the current section does not retarget an already mounted agent')
  assert.equal(await evaluate(`(() => {
    const prompts = Array.from(document.querySelectorAll('[data-agent-request]'));
    return prompts.length >= 2 && new Set(prompts.map(input => input.id)).size === prompts.length && prompts.every(input => Array.from(input.labels ?? []).some(label => label.closest('[data-scroll-section]') === input.closest('[data-scroll-section]')));
  })()`), true, 'mounted agents have unique prompt ids with labels belonging to their own section')

  // Reproduce a preceding lazy editor growing after an explicit destination has
  // already been reached. Disable native scroll anchoring just for this fixture
  // so Chromium cannot mask the application's anchor-preservation behavior.
  await anchor('agent-sheets')
  await until(agentState('agent-sheets', 'ready'), 'preceding agent initialized for layout regression', 90_000)
  const layoutFixture = await evaluate(`({
    documentStyle: document.documentElement.getAttribute('style'),
    precedingStyle: document.querySelector('${toolSection('agent-sheets')}').getAttribute('style'),
    precedingPadding: parseFloat(getComputedStyle(document.querySelector('${toolSection('agent-sheets')}')).paddingBottom) || 0,
  })`)
  try {
    await evaluate(`document.documentElement.style.overflowAnchor = 'none'`)
    await anchor('agent-docs')
    const docsAligned = `(() => {
      const element = document.querySelector('${section('agent-docs')}');
      return Math.abs(element.getBoundingClientRect().top - parseFloat(getComputedStyle(element).scrollMarginTop)) < 4;
    })()`
    await until(docsAligned, 'warm DOCX anchor reaches its heading before delayed growth')
    await evaluate(`new Promise(resolve => setTimeout(() => {
      document.querySelector('${toolSection('agent-sheets')}').style.paddingBottom = '${layoutFixture.precedingPadding + 400}px';
      resolve();
    }, 50))`)
    await until(`${docsAligned} && ${active('agent-docs')}`, 'delayed preceding growth preserves the explicit DOCX destination')

    await wheelTo('slides')
    await evaluate(`new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)))`)
    const positionAfterWheel = await evaluate('scrollY')
    await evaluate(`document.querySelector('${toolSection('agent-sheets')}').style.paddingBottom = '${layoutFixture.precedingPadding + 800}px'`)
    await evaluate(`new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(() => requestAnimationFrame(resolve))))`)
    assert.ok(Math.abs(await evaluate('scrollY') - positionAfterWheel) < 4, 'native wheel releases the destination pin before another layout change')
    assert.equal(await evaluate(docsAligned), false, 'later growth does not snap the user back to the previous DOCX destination')
  } finally {
    await evaluate(`(() => {
      for (const [element, style] of [
        [document.documentElement, ${JSON.stringify(layoutFixture.documentStyle)}],
        [document.querySelector('${toolSection('agent-sheets')}'), ${JSON.stringify(layoutFixture.precedingStyle)}],
      ]) {
        if (style === null) element.removeAttribute('style');
        else element.setAttribute('style', style);
      }
    })()`)
  }

  await send('Emulation.setDeviceMetricsOverride', { width: 390, height: 844, deviceScaleFactor: 1, mobile: true })
  await anchor('agent-pdf')
  assert.equal(await evaluate(`document.documentElement.scrollWidth <= 390`), true, 'mobile document has no horizontal overflow')
  await evaluate(`document.querySelector('.examples-index-link').click()`)
  assert.equal(await evaluate(`document.activeElement.id === 'demo-navigation' && getComputedStyle(document.querySelector('.tool-nav')).display === 'grid'`), true, 'mobile header returns to the fully listed examples index')
  await assertNavigationSelection('mobile')
  await screenshot('scroll-navigation-mobile')
  await anchor('charts')
  assert.equal(await evaluate(`document.documentElement.scrollWidth <= 390`), true, 'later mounted editor does not overflow the mobile document')
  await screenshot('scroll-chart-mobile')
  // Use a fresh document so the typography workspace really is untouched: a
  // prior edited DOCX agent correctly pins its entire containing workspace.
  origin.hash = '#/docs?feature=font-metrics'
  origin.searchParams.set('scroll-smoke-document', 'retention')
  await send('Page.navigate', { url: origin.href })
  Object.assign(rememberedFeatures, { sheets: 'agent', docs: 'font-metrics', slides: 'agent', pdf: 'agent' })
  await until(ready('font-metrics'), 'untouched typography workspace loads')
  await anchor('charts')
  await until(`document.querySelector('${section('charts')} input[aria-label="Jan revenue"]')`, 'chart input available after retention reload')
  const sampleChartRevenue = await evaluate(`document.querySelector('${section('charts')} input[aria-label="Jan revenue"]').value`)
  assert.notEqual(sampleChartRevenue, '999', 'reset sample differs from the retained edit')
  await evaluate(`(() => {
    const input = document.querySelector('${section('charts')} input[aria-label="Jan revenue"]');
    Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(input, '999');
    input.dispatchEvent(new Event('input', { bubbles: true }));
  })()`)
  // Untouched offscreen engines are released; interacted documents stay intact.
  // Viewport-sized canvases can leave the next section within the 160px preload
  // margin at any viewport height. Move this untouched test target beyond it.
  await evaluate(`(() => {
    const target = document.querySelector('${toolSection('font-metrics')}');
    target.style.marginTop = Math.max(0, innerHeight + 200 - target.getBoundingClientRect().top) + 'px';
  })()`)
  assert.ok(await evaluate(`document.querySelector('${toolSection('font-metrics')}').getBoundingClientRect().top > innerHeight + 160`), 'retention target is outside the prefetch margin')
  await until(`document.querySelector('${toolSection('font-metrics')}').dataset.scrollState === 'idle'`, 'offscreen untouched workspace released', 45_000)
  assert.equal(await evaluate(`document.querySelector('${section('charts')} input[aria-label="Jan revenue"]').value`), '999', 'interacted chart is never automatically discarded')
  assert.equal(await evaluate(`!!document.querySelector('.demo-options, .demo-reset-trigger')`), false, 'demo reset/close menu is absent')
  await anchor('sheets')
  await until(ready('charts'), 'same sidebar destination keeps the edited example available')
  assert.equal(await evaluate(`document.querySelector('${section('charts')} input[aria-label="Jan revenue"]').value`), '999', 'reselecting the tool preserves edits')
  assert.deepEqual(liveProposals, [], 'scroll demo never calls a real model endpoint')
  assert.deepEqual(errors, [], 'no uncaught errors or console errors')
  console.log(JSON.stringify({ status: 'passed', mode: process.argv.includes('--dev') ? 'development' : process.argv.includes('--built') ? 'built' : 'existing-server', screenshots: output, checks: ['exactly four tools in one document without an intro', 'direct Docs feature start', '28 sidebar examples', 'lazy feature initialization', 'sidebar scroll spy', 'passive scroll replaces history', 'passive scroll preserves focus', 'sidebar pointer and keyboard navigation focuses example headings', 'retained workbook stays measurable in document flow', 'feature changes retain native edits and exact pending approval', 'same tool anchor returns to heading', 'Back and Forward across features', 'legacy format-specific deep links', 'independent mounted agent formats', 'untouched workspace release and edited state retention', 'wrapped mobile examples index', 'mobile overflow', 'no real model calls'], errors }, null, 2))
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
