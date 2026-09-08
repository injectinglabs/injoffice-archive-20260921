import assert from 'node:assert/strict'
import { existsSync, mkdirSync, mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { resolve } from 'node:path'
import { launchChromeForCDP, terminateProcess } from './chrome-cdp-startup.mjs'
import { startShowcaseServer } from './showcase-smoke-server.mjs'
import { startShowcaseDevServer } from './showcase-smoke-dev-server.mjs'

// This suite exercises the document navigator, not the individual editors. The
// existing showcase smoke remains responsible for real file and tool proofs.
const keys = ['overview', 'sheets', 'docs', 'slides', 'pdf']
const rememberedFeatures = { sheets: 'editor', docs: 'editor', slides: 'editor', pdf: 'editor' }
const target = key => key.startsWith('agent-') ? { tool: key.slice(6), feature: 'agent' }
  : keys.includes(key) ? { tool: key, feature: rememberedFeatures[key] }
    : { tool: key.startsWith('pptx-') ? 'slides' : key === 'font-metrics' ? 'docs' : 'sheets', feature: key }
const href = key => { const { tool, feature } = target(key); return tool === 'overview' ? '#/overview' : `#/${tool}?feature=${feature}` }
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
let allowResetDialog = false
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
    void send('Page.handleJavaScriptDialog', { accept: message.params.type === 'confirm' && ((allowReloadDialog && /unsaved.*lost/i.test(message.params.message)) || (allowResetDialog && /edits.*lost/i.test(message.params.message))) })
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
      await evaluate(`document.querySelector('.app-sidebar a[aria-current="location"]').focus({ preventScroll: true })`)
      for (const modifiers of [8, 0]) {
        await send('Input.dispatchKeyEvent', { type: 'keyDown', key: 'Tab', code: 'Tab', windowsVirtualKeyCode: 9, modifiers })
        await send('Input.dispatchKeyEvent', { type: 'keyUp', key: 'Tab', code: 'Tab', windowsVirtualKeyCode: 9, modifiers })
      }
      const focusProof = await evaluate(`(() => {
        const selected = document.querySelector('.app-sidebar a[aria-current="location"]');
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

const active = (key, hash = href(key)) => key === 'overview'
  ? `location.hash === '#/overview' && !document.querySelector('.app-sidebar a[aria-current="location"]')`
  : `(location.hash === ${JSON.stringify(hash)} || location.hash === ${JSON.stringify(href(key))} || (${keys.includes(key)} && location.hash === ${JSON.stringify(`#/${target(key).tool}`)})) && document.querySelector('.app-sidebar a[aria-current="location"]')?.getAttribute('href') === ${JSON.stringify(`#/${target(key).tool}`)}`
const ready = key => key === 'overview' ? `document.querySelector('${toolSection(key)}')?.dataset.scrollState === 'ready'`
  : `document.querySelector('${toolSection(key)}')?.dataset.scrollState === 'ready' && !!document.querySelector('${toolSection(key)} [data-workspace-panel="${target(key).feature}"]:not([hidden])') && !document.querySelector('${toolSection(key)} [data-workspace-panel="${target(key).feature}"] [data-workspace-loading]') && !document.querySelector('${toolSection(key)} [data-workspace-panel="${target(key).feature}"] [data-workspace-error]')`
const agentState = (key, state) => `document.querySelector(${JSON.stringify(`${section(key)} .agent-demo__status`)})?.dataset.state === ${JSON.stringify(state)}${state === 'ready' ? ` && Array.from(document.querySelector(${JSON.stringify(section(key))}).querySelectorAll('button')).some(button => button.textContent.trim() === 'Run agent' && !button.disabled)` : ''}`

async function anchor(key) {
  const { tool, feature } = target(key)
  if (key === 'overview') await evaluate(`document.querySelector('.app-brand').click()`)
  else if (keys.includes(key)) await evaluate(`document.querySelector('.app-sidebar a[href="#/${tool}"]').click()`)
  else {
    rememberedFeatures[tool] = feature
    await evaluate(`location.hash = ${JSON.stringify(href(key))}`)
  }
  await until(active(key), `${key} anchor updates the current location`)
  if (key !== 'overview') await until(ready(key), `${key} lazy section ready`, 90_000)
}

async function button(key, text) {
  await evaluate(`Array.from(document.querySelector(${JSON.stringify(section(key))}).querySelectorAll('button')).find(button => button.textContent.trim() === ${JSON.stringify(text)}).click()`)
}

async function assertInternalGroupNavigation(group, feature, key) {
  const navigation = `${toolSection('docs')} .tool-workspace__navigation`
  const groupSelector = `${toolSection('docs')} [data-workspace-group="${group}"]`
  // A same-tool anchor can already have the expected hash while its queued
  // animation frame still positions and focuses the heading. Wait for stable,
  // hittable real geometry before sending pointer input, not just a ready panel.
  const targetPoint = await evaluate(`new Promise((resolve, reject) => {
    let previous = null, stable = 0, attempts = 0;
    const sample = () => {
      const button = document.querySelector('${groupSelector}');
      const rect = button?.getBoundingClientRect();
      const bar = document.querySelector('${navigation}')?.getBoundingClientRect();
      if (rect && bar) {
        const point = { x: rect.left + rect.width / 2, y: rect.top + rect.height / 2, top: bar.top };
        const hit = document.elementFromPoint(point.x, point.y)?.closest('button') === button;
        const unchanged = previous && ['x', 'y', 'top'].every(axis => Math.abs(point[axis] - previous[axis]) < 0.5);
        stable = hit && unchanged ? stable + 1 : 0;
        previous = point;
        if (stable >= 4) { resolve(point); return; }
      }
      if (++attempts > 80) reject(new Error('Capability tab never reached stable, visible hit-test geometry'));
      else setTimeout(sample, 50);
    };
    sample();
  })`)
  const beforeTop = targetPoint.top
  if (key) {
    await send('Input.dispatchKeyEvent', { type: 'keyDown', key, code: key, windowsVirtualKeyCode: key === 'ArrowLeft' ? 37 : 39 })
    await send('Input.dispatchKeyEvent', { type: 'keyUp', key, code: key, windowsVirtualKeyCode: key === 'ArrowLeft' ? 37 : 39 })
  } else {
    await evaluate(`(() => {
      window.__workspacePointerProof = {};
      for (const type of ['pointerdown', 'click']) window.addEventListener(type, event => {
        const group = event.target instanceof Element ? event.target.closest('[data-workspace-group]') : null;
        window.__workspacePointerProof[type] = { group: group?.dataset.workspaceGroup, tool: group?.closest('[data-tool-workspace]')?.dataset.toolWorkspace };
      }, { once: true, capture: true });
    })()`)
    const point = { x: targetPoint.x, y: targetPoint.y }
    await send('Input.dispatchMouseEvent', { type: 'mousePressed', button: 'left', clickCount: 1, ...point })
    await send('Input.dispatchMouseEvent', { type: 'mouseReleased', button: 'left', clickCount: 1, ...point })
    assert.deepEqual(await evaluate('window.__workspacePointerProof'), { pointerdown: { group, tool: 'docs' }, click: { group, tool: 'docs' } }, 'real pointer input hits the intended Docs capability tab')
  }
  rememberedFeatures.docs = feature
  await until(`${ready('docs')} && document.querySelector('${groupSelector}').getAttribute('aria-selected') === 'true'`, `${key ?? 'pointer click'} selects the ${group} capability group`)
  const samples = await evaluate(`new Promise(resolve => {
    const samples = [];
    const sample = () => {
      const bar = document.querySelector('${navigation}');
      const selected = document.querySelector('${groupSelector}');
      samples.push({ top: bar.getBoundingClientRect().top, focused: document.activeElement === selected, active: document.activeElement?.outerHTML.slice(0, 240), current: document.querySelector('.app-sidebar a[aria-current="location"]')?.getAttribute('href') });
      if (samples.length === 5) resolve(samples); else setTimeout(sample, 50);
    };
    sample();
  })`)
  assert.ok(samples.every(sample => sample.focused), `${key ?? 'pointer click'} retains focus on the selected ${group} tab after its hash change: ${JSON.stringify(samples)}`)
  assert.ok(samples.every(sample => sample.current === '#/docs'), 'internal navigation keeps Docs selected in the sidebar')
  assert.ok(samples.every(sample => Math.abs(sample.top - beforeTop) <= 2), `${key ?? 'pointer click'} preserves the capability bar offset over 200ms: before=${beforeTop}, samples=${JSON.stringify(samples)}`)
}

async function wheelTo(key, hash = href(key)) {
  // Native wheel input does not run an anchor's navigation handler. The target
  // offset comes from the real rendered document, so lazy section heights can vary.
  const distance = await evaluate(`(() => {
    const target = document.querySelector(${JSON.stringify(toolSection(key))});
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
  assert.deepEqual(await evaluate(`Array.from(document.querySelectorAll('.app-sidebar a')).map(link => link.getAttribute('href'))`), ['#/sheets', '#/docs', '#/slides', '#/pdf'], 'only four comprehensive tool links appear in navigation')
  await screenshot('scroll-overview-desktop')
  await anchor('font-metrics')
  await assertNavigationSelection('desktop')
  await anchor('overview')

  if (process.argv.includes('--fault')) {
    await send('Fetch.enable', { patterns: [{ urlPattern: '*PptxRenderPage*', resourceType: 'Script', requestStage: 'Request' }] })
    rememberedFeatures.slides = 'pptx-render'
    await evaluate(`location.hash = '#/slides?feature=pptx-render'`)
    await until(`!!document.querySelector('${section('pptx-render')} [data-workspace-error]')`, 'failed chunk is isolated to its feature')
    assert.equal(chunkFailures, 1, 'the test deliberately failed exactly one lazy chunk')
    assert.equal(await evaluate(`document.querySelectorAll('[data-scroll-section]').length === 5 && !!document.querySelector('${section('pptx-render')} [role=alert]')`), true, 'one failed feature leaves all four workspaces and navigation intact')
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
    await anchor('overview')
  }

  // Cold navigation preserves the overview DOM and uses a real section anchor.
  await evaluate(`void (window.__overviewNode = document.querySelector('[data-scroll-section="overview"]'))`)
  await anchor('agent-docs')
  await until(agentState('agent-docs', 'ready'), 'DOCX agent initialized', 90_000)
  assert.equal(await evaluate(`window.__overviewNode.isConnected && window.__overviewNode === document.querySelector('[data-scroll-section="overview"]')`), true, 'anchor navigation preserves existing sections')
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
  assert.equal(await evaluate(`document.activeElement === window.__persistedPrompt`), true, 'passive scrolling does not steal focus from an editor')
  assert.equal(await evaluate(`history.length === window.__beforePassiveHistory.length && window.__scrollHistory.pushes === window.__beforePassiveHistory.pushes && window.__scrollHistory.replacements > window.__beforePassiveHistory.replacements`), true, 'scroll spy replaces the URL without adding browser history entries')
  assert.equal(await evaluate(`document.querySelectorAll('.app-sidebar a[aria-current="location"]').length`), 1, 'exactly one sidebar link is current')
  await anchor('agent-docs')
  assert.equal(await evaluate(`window.__persistedPrompt.isConnected && document.querySelector('${section('agent-docs')} [data-agent-task-value]').value === ${JSON.stringify(prompt)}`), true, 'scrolling away and back preserves typed guided task state')
  await button('agent-docs', 'Run agent')
  await until(agentState('agent-docs', 'awaiting-approval'), 'real DOCX preview ready', 90_000)
  const plan = await evaluate(`document.querySelector('${section('agent-docs')} .agent-diff').textContent`)
  await anchor('sheets')
  await until(`!!document.querySelector('${toolSection('sheets')} [data-workspace-panel="editor"] canvas')`, 'main workbook rendered before retention check', 90_000)
  await anchor('charts')
  await until(`document.querySelector('${section('charts')} input[aria-label="Jan revenue"]')`, 'editable chart mounted')
  assert.equal(await evaluate(`(() => {
    const panel = document.querySelector('${toolSection('sheets')} [data-workspace-panel="editor"]');
    const rect = panel.getBoundingClientRect();
    return panel.hidden && panel.inert && panel.getAttribute('aria-hidden') === 'true' && getComputedStyle(panel).visibility === 'hidden' && rect.width > 0 && rect.height > 0;
  })()`), true, 'retained workbook stays measurable for its canvas but inert and hidden from assistive technology')
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
  assert.equal(await evaluate(`document.querySelector('${section('agent-docs')}').hidden`), true, 'inactive agent feature is retained but hidden')
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

  // Re-selecting the current hash still returns to the section heading.
  await evaluate(`window.scrollBy({ top: 320, behavior: 'instant' })`)
  await anchor('docs')
  await until(`(() => { const element = document.querySelector('${toolSection('docs')}'); const top = element.getBoundingClientRect().top; const margin = parseFloat(getComputedStyle(element).scrollMarginTop) || 100; return Math.abs(top - margin) < 40; })()`, 'same tool anchor returns to its heading without changing the selected feature')
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
  origin.hash = '#/overview'
  origin.searchParams.set('scroll-smoke-document', 'cold-sheets')
  await send('Page.navigate', { url: origin.href })
  Object.assign(rememberedFeatures, { sheets: 'editor', docs: 'editor', slides: 'editor', pdf: 'editor' })
  await until(active('overview'), 'fresh document ready for cold Sheets deep link')
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
  Object.assign(rememberedFeatures, { sheets: 'editor', docs: 'editor', slides: 'editor', pdf: 'agent' })
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
      const element = document.querySelector('${toolSection('agent-docs')}');
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
  assert.equal(await evaluate(`(() => { const navigator = document.querySelector('.app-sidebar'); const bounds = navigator.getBoundingClientRect(); return bounds.top >= -1 && bounds.top < 200 && bounds.bottom > 0 && bounds.bottom < innerHeight; })()`), true, 'mobile section navigator stays in the viewport')
  assert.equal(await evaluate(`(() => { const link = document.querySelector('.app-sidebar a[aria-current="location"]'); const bounds = link.getBoundingClientRect(); return bounds.left >= -1 && bounds.right <= innerWidth + 1 && bounds.top >= -1 && bounds.bottom <= innerHeight; })()`), true, 'the current mobile section link is visible')
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
  Object.assign(rememberedFeatures, { sheets: 'editor', docs: 'font-metrics', slides: 'editor', pdf: 'editor' })
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
  await until(`document.querySelector('${toolSection('font-metrics')}').dataset.scrollState === 'idle'`, 'offscreen untouched workspace released', 45_000)
  assert.equal(await evaluate(`document.querySelector('${section('charts')} input[aria-label="Jan revenue"]').value`), '999', 'interacted chart is never automatically discarded')
  const resetChart = `Array.from(document.querySelectorAll('${toolSection('charts')} .demo-context-actions button')).find(button => button.textContent.trim() === 'Reset demo').click()`
  await evaluate(resetChart)
  assert.equal(await evaluate(`document.querySelector('${section('charts')} input[aria-label="Jan revenue"]').value`), '999', 'cancel reset preserves edits')
  allowResetDialog = true
  await evaluate(resetChart)
  allowResetDialog = false
  await until(`document.querySelector('${section('charts')} input[aria-label="Jan revenue"]')?.value === ${JSON.stringify(sampleChartRevenue)}`, 'confirmed reset remounts the chart with the exact original sample')
  await evaluate(`Array.from(document.querySelectorAll('${toolSection('charts')} .demo-context-actions button')).find(button => button.textContent.trim() === 'Close demo').click()`)
  await until(`document.querySelector('${toolSection('charts')}').dataset.scrollState === 'idle' && !document.querySelector('${section('charts')} input')`, 'close releases the mounted workspace')
  await evaluate(`document.querySelector('${toolSection('charts')} .demo-section-placeholder button').click()`)
  await until(ready('charts'), 'closed demo can be reopened')
  allowResetDialog = true
  await evaluate(`Array.from(document.querySelectorAll('${toolSection('charts')} .demo-context-actions button')).find(button => button.textContent.trim() === 'Close demo').click()`)
  allowResetDialog = false
  await anchor('sheets')
  await until(ready('charts'), 'same sidebar destination reopens a closed demo')
  assert.deepEqual(liveProposals, [], 'scroll demo never calls a real model endpoint')
  assert.deepEqual(errors, [], 'no uncaught errors or console errors')
  console.log(JSON.stringify({ status: 'passed', mode: process.argv.includes('--dev') ? 'development' : process.argv.includes('--built') ? 'built' : 'existing-server', screenshots: output, checks: ['four tools plus landing in one document', 'four sidebar links', 'lazy feature initialization', 'sidebar scroll spy', 'passive scroll replaces history', 'passive scroll preserves focus', 'pointer and arrow-key groups retain focus and toolbar position', 'retained hidden workbook remains inert and keyboard-inaccessible', 'feature changes retain native edits and exact pending approval', 'same tool anchor returns to heading', 'Back and Forward across features', 'legacy format-specific deep links', 'independent mounted agent formats', 'untouched workspace release and guarded resets', 'sticky mobile navigator', 'mobile overflow', 'no real model calls'], errors }, null, 2))
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
