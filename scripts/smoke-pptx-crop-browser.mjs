// Pixel oracle for the actual real-file preview component; no Office fidelity claim.
import { spawnSync } from 'node:child_process'
import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { resolve, sep } from 'node:path'
import { createHash } from 'node:crypto'
import { build } from 'vite'
import { PNG } from 'pngjs'
import { launchChromeForCDP, terminateProcess } from './chrome-cdp-startup.mjs'

const root = resolve(import.meta.dirname, '..')
const artifacts = process.env.SHOWCASE_OUTPUT ? resolve(process.env.SHOWCASE_OUTPUT) : mkdtempSync(resolve(tmpdir(), 'injoffice-crop-pixels-'))
mkdirSync(artifacts, { recursive: true })
const image = new PNG({ width: 80, height: 80 })
const colors = [[255, 0, 0], [0, 255, 0], [0, 0, 255], [255, 255, 0]]
for (let y = 0; y < 80; y++) for (let x = 0; x < 80; x++) {
  const color = colors[(y >= 40 ? 2 : 0) + (x >= 40 ? 1 : 0)]
  image.data.set([...color, 255], (y * 80 + x) * 4)
}
const bytes = PNG.sync.write(image)
const crops = [undefined, { left: 50000, top: 0, right: 0, bottom: 50000 }, { left: 0, top: 50000, right: 50000, bottom: 0 }]
const deck = {
  contractVersion: 'pptx-native/v1', documentId: 'crop-pixel-oracle', origin: 'authored',
  size: { cx: 9600000, cy: 3200000 },
  assets: [{ id: 'quadrants', provenance: 'authored', contentType: 'image/png', sha256: createHash('sha256').update(bytes).digest('hex'), byteLength: bytes.length, dataBase64: bytes.toString('base64'), passthrough: [] }],
  compatibility: { status: 'editable', diagnostics: [] },
  slides: [{ id: 'slide', provenance: 'authored', passthrough: [], compatibility: { status: 'editable', diagnostics: [] }, elements: crops.map((crop, i) => ({ kind: 'picture', id: `picture-${i}`, provenance: 'authored', assetId: 'quadrants', transform: { x: i * 3200000, y: 0, cx: 3200000, cy: 3200000 }, ...(crop ? { crop } : {}), passthrough: [], compatibility: { status: 'editable', diagnostics: [] } })) }],
}
const entry = resolve(root, 'virtual-pptx-crop-smoke.js')
const built = await build({
  configFile: false, root, logLevel: 'warn',
  plugins: [{
    name: 'pptx-crop-smoke-entry',
    resolveId(id) { if (id === entry) return entry },
    load(id) { if (id === entry) return "import {createElement} from 'react'; import {createRoot} from 'react-dom/client'; import Preview from './apps/playground/src/components/PptxFilePreview.tsx'; const root = createRoot(document.body); globalThis.__injofficeRenderFixture = deck => root.render(createElement(Preview, {deck})); globalThis.__injofficeRenderFixture(globalThis.__injofficeCropFixture);" },
  }],
  define: { 'process.env.NODE_ENV': '"production"' },
  build: {
    write: false, minify: false,
    lib: { entry, name: 'InjOfficeCropSmoke', formats: ['iife'] },
  },
})
const bundle = (Array.isArray(built) ? built : [built]).flatMap(result => result.output).find(output => output.type === 'chunk' && output.isEntry)
if (!bundle) throw new Error('Vite did not produce the crop smoke entry bundle')
const profiles = []
let chrome, cdp
try {
  chrome = await launchChromeForCDP({ executable: findChrome(), createProfile: () => {
    const profile = mkdtempSync(resolve(tmpdir(), 'injoffice-crop-chrome-'))
    profiles.push(profile)
    return profile
  } })
  cdp = await connect(chrome.target.webSocketDebuggerUrl)
  await cdp.send('Page.enable')
  await cdp.send('Emulation.setDeviceMetricsOverride', { width: 1000, height: 700, deviceScaleFactor: 1, mobile: false })
  const global = await cdp.send('Runtime.evaluate', { expression: 'globalThis' })
  const assigned = await cdp.send('Runtime.callFunctionOn', { objectId: global.result.objectId, functionDeclaration: 'function (deck) { this.__injofficeCropFixture = deck; }', arguments: [{ value: deck }] })
  if (assigned.exceptionDetails) throw new Error('Could not transfer crop fixture as data')
  await evaluate(bundle.code)
  await evaluate(`new Promise((resolve, reject) => { const end = Date.now() + 10000; function check() { const svg = document.querySelector('svg[role="img"]'); if (svg) { Promise.all([...svg.querySelectorAll('image')].map(node => new Promise((done, fail) => { const image = new Image(); image.onload = done; image.onerror = fail; image.src = node.getAttribute('href'); }))).then(() => requestAnimationFrame(() => requestAnimationFrame(resolve)), reject); } else if (Date.now() > end) reject(new Error('preview not mounted')); else setTimeout(check, 20); } check(); })`)
  const bounds = await evaluate(`(() => { const r = document.querySelector('svg[role="img"]').getBoundingClientRect(); return {x:r.x,y:r.y,width:r.width,height:r.height}; })()`)
  const capture = await cdp.send('Page.captureScreenshot', { format: 'png' })
  const pngBytes = Buffer.from(capture.data, 'base64')
  writeFileSync(resolve(artifacts, 'pptx-positive-crop.png'), pngBytes)
  const pixels = PNG.sync.read(pngBytes)
  // Uncropped control retains four colors. Top-right crop is uniformly green;
  // bottom-left crop uniformly blue, including near every frame corner.
  for (const [panel, expected] of [[0, colors], [1, Array(4).fill(colors[1])], [2, Array(4).fill(colors[2])]]) {
    for (const [i, [u, v]] of [[0.2, 0.2], [0.8, 0.2], [0.2, 0.8], [0.8, 0.8]].entries()) {
      const x = Math.round(bounds.x + bounds.width * (panel + u) / 3)
      const y = Math.round(bounds.y + bounds.height * v)
      const actual = [...pixels.data.subarray((y * pixels.width + x) * 4, (y * pixels.width + x) * 4 + 3)]
      if (actual.some((channel, j) => Math.abs(channel - expected[i][j]) > 2)) throw new Error(`crop panel ${panel} sample ${i}: ${actual}, expected ${expected[i]}`)
    }
  }
  console.log(`PPTX crop Chrome pixel oracle: PASS (${artifacts})`)
  if (process.argv.includes('--styles')) {
    const generated = spawnSync('go', ['test', '-count=1', '-run', '^TestNativeTextStyleBrowserFixture$', '.'], { cwd: resolve(root, 'go/pptxpatch'), env: { ...process.env, INJOFFICE_PPTX_STYLE_FIXTURE_DIR: artifacts }, encoding: 'utf8', timeout: 60000 })
    if (generated.status !== 0) throw new Error(`Could not generate/extract the real style fixture: ${generated.stderr}\n${generated.stdout}`)
    const native = JSON.parse(readFileSync(resolve(artifacts, 'styled-native.json'), 'utf8'))
    const rendered = await cdp.send('Runtime.callFunctionOn', { objectId: global.result.objectId, functionDeclaration: 'function (deck) { this.__injofficeRenderFixture(deck); }', arguments: [{ value: native }] })
    if (rendered.exceptionDetails) throw new Error('Could not render real extracted style fixture')
    const check = await evaluate(`new Promise((resolve, reject) => { const end = Date.now() + 10000; function check() { const p = document.querySelector('foreignObject p'); if (p?.textContent === '▪ Hello world') { const spans = [...p.querySelectorAll(':scope > span')]; const styles = spans.map(span => { const style=getComputedStyle(span); return {text:span.textContent,size:style.fontSize,weight:style.fontWeight,italic:style.fontStyle,color:style.color}; }); const paragraph = getComputedStyle(p); resolve({styles,margin:parseFloat(paragraph.paddingLeft),indent:parseFloat(paragraph.textIndent),label:document.body.textContent.includes('Approximate file preview')}); } else if (Date.now() > end) reject(new Error('inherited style content not rendered')); else setTimeout(check,20); } check(); })`)
    if (check.styles.length !== 3 || check.styles.some((style, i) => style.text !== ['▪ ', 'Hello ', 'world'][i] || style.size !== '24px' || style.weight !== (i === 2 ? '400' : '700') || style.italic !== 'italic' || style.color !== 'rgb(47, 111, 237)') || Math.abs(check.margin - 23.622) > .01 || Math.abs(check.indent + 7.874) > .01 || !check.label) throw new Error(`Real-file marker/text-run inherited styles differ: ${JSON.stringify(check)}`)
    const styled = await cdp.send('Page.captureScreenshot', { format: 'png' })
    writeFileSync(resolve(artifacts, 'pptx-native-styles.png'), Buffer.from(styled.data, 'base64'))
    console.log(`PPTX real-file local style cascade and authored bullet: PASS (${artifacts})`)
    const placeholder = JSON.parse(readFileSync(resolve(artifacts, 'placeholder-native.json'), 'utf8'))
    const inherited = await cdp.send('Runtime.callFunctionOn', { objectId: global.result.objectId, functionDeclaration: 'function (deck) { this.__injofficeRenderFixture(deck); }', arguments: [{ value: placeholder }] })
    if (inherited.exceptionDetails) throw new Error('Could not render extracted placeholder fixture')
    const inheritance = await evaluate(`new Promise((resolve,reject)=>{const end=Date.now()+10000;function check(){const p=document.querySelector('foreignObject p');const margin=p?parseFloat(getComputedStyle(p).paddingLeft):0;if(Math.abs(margin-31.496)<.01){resolve({text:p.textContent,transform:p.closest('g').getAttribute('transform'),diagnostic:document.body.textContent.includes('inherited targets remain preserve-only')})}else if(Date.now()>end)reject(new Error('placeholder projection not rendered'));else setTimeout(check,20)}check()})`)
    if (inheritance.text !== '▪ Hello world' || !inheritance.transform.includes('72') || !inheritance.diagnostic) throw new Error(`Real-file placeholder inheritance differs: ${JSON.stringify(inheritance)}`)
    const inheritedImage = await cdp.send('Page.captureScreenshot', { format: 'png' })
    writeFileSync(resolve(artifacts, 'pptx-native-placeholder.png'), Buffer.from(inheritedImage.data, 'base64'))
    console.log(`PPTX real-file relationship-bound master/layout placeholder: PASS (${artifacts})`)
  }
} finally {
  cdp?.close()
  try { if (chrome) await terminateProcess(chrome.child) }
  finally { for (const profile of profiles) rmSync(profile, { recursive: true, force: true, maxRetries: 10, retryDelay: 100 }) }
}

async function evaluate(expression) {
  const result = await cdp.send('Runtime.evaluate', { expression, awaitPromise: true, returnByValue: true })
  if (result.exceptionDetails) throw new Error(result.exceptionDetails.exception?.description ?? result.exceptionDetails.text)
  return result.result.value
}
function findChrome() {
  for (const candidate of [process.env.CHROME_PATH, process.env.CHROME_BIN, '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome', 'google-chrome', 'google-chrome-stable', 'chromium', 'chromium-browser'].filter(Boolean)) {
    if (candidate.includes(sep) && !existsSync(candidate)) continue
    if (spawnSync(candidate, ['--version'], { stdio: 'ignore' }).status === 0) return candidate
  }
  throw new Error('Chrome/Chromium required; set CHROME_PATH')
}
async function connect(url) {
  const socket = new WebSocket(url)
  await new Promise((resolvePromise, reject) => { socket.addEventListener('open', resolvePromise, { once: true }); socket.addEventListener('error', reject, { once: true }) })
  let next = 0
  const pending = new Map()
  socket.addEventListener('message', event => {
    const message = JSON.parse(event.data), request = pending.get(message.id)
    if (!request) return
    pending.delete(message.id); clearTimeout(request.timer)
    if (message.error) request.reject(new Error(message.error.message)); else request.resolve(message.result)
  })
  return {
    send(method, params = {}) { return new Promise((resolvePromise, reject) => { const id = ++next; const timer = setTimeout(() => { pending.delete(id); reject(new Error(`CDP timeout: ${method}`)) }, 30000); pending.set(id, { resolve: resolvePromise, reject, timer }); socket.send(JSON.stringify({ id, method, params })) }) },
    close() { socket.close() },
  }
}
