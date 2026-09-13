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
const scratch = mkdtempSync(resolve(tmpdir(), 'injoffice-docx-textbox-notes-smoke-'))
const artifacts = process.env.SHOWCASE_OUTPUT ? resolve(process.env.SHOWCASE_OUTPUT) : mkdtempSync(resolve(tmpdir(), 'injoffice-docx-native-screenshots-'))
mkdirSync(artifacts, { recursive: true })
const profiles = []
const errors = []
const docs = `document.querySelector('[data-demo-surface="docs"]')`
const native = `document.querySelector('[aria-label="Native document pages"]')`
let helper, server, chrome, cdp
let helperLog = ''
let checks=0
const previousApi = process.env.VITE_INJOFFICE_API_BASE
const previousServer = process.env.INJOFFICE_SERVER

try {
  const worker = resolve(root, 'apps/docx-page-paint-worker/dist/worker.js')
  if (!existsSync(worker)) throw new Error('Build the workspace packages and DOCX page-paint worker before running this smoke.')
  const textboxDir=resolve(scratch,'textbox'),footnoteDir=resolve(scratch,'footnote')
  for(const [dir,test,env] of [[textboxDir,'TestNativeTextboxPageSource','INJOFFICE_TEXTBOX_PAGE_EVIDENCE_DIR'],[footnoteDir,'TestNativeFootnoteContinuationSource','INJOFFICE_FOOTNOTE_EVIDENCE_DIR']]){
    const result=spawnSync('go',['test','-count=1','-run','^'+test+'$','./cmd/nativepreviewfixture'],{cwd:resolve(root,'go/docxpatch'),env:{...process.env,[env]:dir},encoding:'utf8',timeout:120000})
    if(result.status!==0)throw new Error('Source fixture export failed: '+result.stderr+' '+result.stdout)
  }
  const textboxFixture=resolve(textboxDir,'page-textbox.docx'),footnoteFixture=resolve(footnoteDir,'footnote-continuation.docx')
  const textboxHash=hash(readFileSync(textboxFixture)),footnoteHash=hash(readFileSync(footnoteFixture))
  const noteSource=JSON.parse(readFileSync(resolve(footnoteDir,'source.json'),'utf8'))
  const label=noteSource.document.notes.flatMap(n=>n.blocks).flatMap(b=>b.paragraph?.runs??[]).find(r=>r.reference?.role==='label').id
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
      const response = await originalFetch(input, init);
      if(init?.method==='POST'&&String(input).endsWith('/v1/docx/page-preview-textboxes')) {
        const value=await response.clone().json();window.__nativeDocxTextbox=value;
        if(window.__corruptTextbox&&value.preview){value.preview.textbox.x_millipoints++;return new Response(JSON.stringify(value),{status:response.status,headers:response.headers});}
      }
      if (init?.method === 'POST' && String(input).endsWith('/v1/docx/page-preview')) window.__nativeDocxPaint = (await response.clone().json()).page_paint_output;
      return response;
    };
  }` })
  await cdp.send('Emulation.setDeviceMetricsOverride', { width: 1440, height: 1100, deviceScaleFactor: 1, mobile: false })
  await cdp.send('Page.navigate', { url: `${server.url}#/docs?feature=editor` })
  await poll(() => evaluate(`Boolean(document.querySelector('[aria-label="Open a DOCX file"]'))`), 'DOCX workbench')
  await upload(textboxFixture)
  await poll(()=>evaluate(`${native}?.textContent.includes('Nothing is uploaded')`),'textbox consent')
  await assert(`window.__nativeDocxPosts.length===0`,'opening textbox document uploads nothing')
  await click('Upload to helper and preview page-placed textboxes')
  await poll(()=>evaluate(`${native}?.querySelector('[data-native-textbox]')!==null&&${native}?.textContent.includes('1 approximate, read-only pages')`),'source rectangle page',45000)
  await assert(`window.__nativeDocxPosts.length===1&&window.__nativeDocxPosts[0].hash===${JSON.stringify(textboxHash)}`,'textbox upload preserves exact original bytes')
  await assert(`(() => {const box=${native}.querySelector('[data-native-textbox]'),rect=box.querySelector('rect');return box.getAttribute('transform')==='translate(72000 144000)'&&rect.getAttribute('width')==='216000'&&rect.getAttribute('height')==='72000'&&rect.getAttribute('stroke-width')==='1000'&&rect.getAttribute('fill')==='#FFF2CC'&&rect.getAttribute('stroke')==='#204060'})()`,'page offsets, bounds, colors and stroke match source')
  await assert(`${native}.querySelectorAll('svg').length===1&&${native}.querySelectorAll('svg path').length===window.__nativeDocxTextbox.preview.body_paint.pages[0].commands.filter(c=>c.kind==='fill_glyph_path').length+window.__nativeDocxTextbox.preview.textbox.paint.paths.length`,'all body and textbox glyph outlines share one mounted page')
  await assert(`${native}.textContent.includes(window.__nativeDocxTextbox.preview.source_diagnostics[0].message)&&${docs}?.dataset.demoDirty!=='true'`,'source drawing warning remains visible and source stays clean')
  await assert(`(async()=>{const svg=${native}.querySelector('svg').cloneNode(true);svg.setAttribute('width','612');svg.setAttribute('height','792');const image=new Image();image.src='data:image/svg+xml;base64,'+btoa(new XMLSerializer().serializeToString(svg));await image.decode();const canvas=document.createElement('canvas');canvas.width=612;canvas.height=792;const context=canvas.getContext('2d');context.drawImage(image,0,0,612,792);const pixel=context.getImageData(80,205,1,1).data;return pixel[0]===255&&pixel[1]===242&&pixel[2]===204&&pixel[3]===255})()`,'actual SVG raster fills the authored physical-page rectangle')
  await screenshot('docx-page-textbox.png')
  await evaluate('window.__corruptTextbox=true')
  await click('Upload to helper and preview page-placed textboxes')
  await poll(()=>evaluate(`${native}?.textContent.includes('Textbox placement does not fit')&&${native}?.querySelector('svg')===null`),'forged coordinate refusal',45000)
  await assert(`window.__nativeDocxPosts.length===2`,'failed preview requires a new explicit upload')
  await evaluate('window.__corruptTextbox=false')
  await upload(footnoteFixture)
  await poll(()=>evaluate(`${native}?.textContent.includes('Nothing is uploaded')&&${native}?.querySelector('svg')===null`),'footnote replacement clears prior pages')
  await assert(`window.__nativeDocxPosts.length===2`,'replacement does not reuse upload consent')
  await click('Upload to helper and render native pages')
  await poll(()=>evaluate(`${native}?.textContent.includes('4 native pages')&&${native}?.querySelector('svg path')!==null`),'footnote continuation pages',45000)
  await assert(`window.__nativeDocxPosts.length===3&&window.__nativeDocxPosts[2].hash===${JSON.stringify(footnoteHash)}`,'footnote upload preserves exact original bytes')
  await assert(`window.__nativeDocxPaint.pages.flatMap(p=>p.commands).filter(c=>c.kind==='fill_glyph_path'&&c.source_id===${JSON.stringify(label)}).length===1`,'continued footnote emits its source label exactly once')
  for(let ordinal=0;ordinal<4;ordinal++){
    if(ordinal){await click('Next native page');await poll(()=>evaluate(`${native}?.querySelector('svg[aria-label="Native document page ${ordinal+1}"]')!==null`),'footnote page '+(ordinal+1))}
    await assert(`(() => {const page=window.__nativeDocxPaint.pages[${ordinal}],rule=page.commands.find(c=>c.kind==='stroke_note_separator'),line=${native}.querySelector('svg line');return !!rule&&!!line&&rule.x2_millipoints-rule.x1_millipoints===${ordinal?468000:144000}&&Number(line.getAttribute('x2'))-Number(line.getAttribute('x1'))===${ordinal?468000:144000}&&${native}.querySelectorAll('svg').length===1&&page.lines.some(l=>l.region==='footnote')&&${ordinal?'page.lines.every(l=>l.region!=="body")':'page.lines.some(l=>l.region==="body")'}})()`,'source separator and note content on page '+(ordinal+1))
    await screenshot('docx-footnote-page-'+(ordinal+1)+'.png')
  }
  await assert(`${docs}?.dataset.demoDirty!=='true'&&window.__nativeDocxPosts.length===3`,'note navigation preserves original source without more uploads')
  if(hash(readFileSync(textboxFixture))!==textboxHash||hash(readFileSync(footnoteFixture))!==footnoteHash)throw Error('Source fixture changed')
  if(errors.length)throw Error(errors.join('\n'))
  writeFileSync(resolve(artifacts,'summary.json'),JSON.stringify({status:'passed',cases:checks,textboxPages:1,footnotePages:4,uploads:3},null,2))
  console.log(`DOCX textbox and footnote browser checks passed: ${checks} cases`)
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
async function assert(expression, label) { if (!await evaluate(expression)) throw new Error(`Failed: ${label}`);checks++ }
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
  await evaluate(`${native}?.querySelector('svg')?.scrollIntoView({ block: 'center' })`)
  const raster=await evaluate(`(async()=>{const svg=${native}.querySelector('svg').cloneNode(true);svg.removeAttribute('style');svg.setAttribute('width','816');svg.setAttribute('height','1056');const image=new Image();image.src='data:image/svg+xml;base64,'+btoa(new XMLSerializer().serializeToString(svg));await image.decode();const canvas=document.createElement('canvas');canvas.width=816;canvas.height=1056;canvas.getContext('2d').drawImage(image,0,0,816,1056);return canvas.toDataURL('image/png').split(',')[1]})()`)
  writeFileSync(resolve(artifacts,name.replace('.png','-page.png')),Buffer.from(raster,'base64'))
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
