// Synthetic source-only browser qualification. External fidelity benchmarks
// and Excel reference exports stay local and are never loaded by this script.
import { spawnSync } from 'node:child_process'
import { createHash } from 'node:crypto'
import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { resolve, sep } from 'node:path'
import { launchChromeForCDP, terminateProcess } from './chrome-cdp-startup.mjs'
import { startShowcaseServer } from './showcase-smoke-server.mjs'
const root = resolve(import.meta.dirname, '..')
const conditional = process.argv.includes('--conditional')
const scratch = mkdtempSync(resolve(tmpdir(), 'injoffice-xlsx-source-'))
const artifacts = process.env.SHOWCASE_OUTPUT ? resolve(process.env.SHOWCASE_OUTPUT) : resolve(scratch, 'screenshots')
mkdirSync(artifacts, { recursive: true })
const profiles = [], errors = []
const panel = `document.querySelector('[aria-label="Read-only source-style recovery"]')`
const grid = `${panel}?.querySelector('table[aria-label="Approximate read-only source grid"]')`
let server, chrome, cdp, checks = 0
try {
  run('go', ['test', '-count=1', '-run', conditional ? '^TestSourceConditionalPreview$' : '^TestSourceStylePreviewKeepsStrictAuthority$', '.'], resolve(root, 'go/xlsxpatch'), conditional ? { XLSX_SOURCE_CONDITIONAL_EVIDENCE_DIR: scratch } : { XLSX_SOURCE_STYLE_EVIDENCE_DIR: scratch })
  if (!process.argv.includes('--skip-build')) run('npm', ['run', 'build:renderer', '-w', 'apps/playground', '--', '--base=/injoffice-smoke/'], root, { VITE_INJOFFICE_API_BASE: '' })
  const source = resolve(scratch, 'source.xlsx'), valid = resolve(root, 'go/xlsxpatch/testdata/excel-authored/happy-tree.xlsx')
  const sourceHash = 'sha256:' + createHash('sha256').update(readFileSync(source)).digest('hex')
  server = await startShowcaseServer(resolve(root, 'apps/playground/dist'))
  chrome = await launchChromeForCDP({ executable: findChrome(), createProfile: () => { const profile = mkdtempSync(resolve(tmpdir(), 'injoffice-xlsx-source-chrome-')); profiles.push(profile); return profile } })
  cdp = await connectCDP(chrome.target.webSocketDebuggerUrl)
  cdp.on('Runtime.exceptionThrown', ({ exceptionDetails }) => errors.push(exceptionDetails.exception?.description ?? exceptionDetails.text))
  await cdp.send('Runtime.enable'); await cdp.send('Page.enable')
  await cdp.send('Page.addScriptToEvaluateOnNewDocument', { source: `{
    const originalFetch=window.fetch.bind(window);window.__sourcePosts=0;
    window.fetch=(input,init)=>{if(init?.method==='POST')window.__sourcePosts++;return originalFetch(input,init)};
    const OriginalWorker=window.Worker;
    window.__sourceV2Requested=0;
    window.Worker=class extends OriginalWorker {
      listeners=new Map();
      constructor(...args){super(...args);this.isConditional=String(args[0]).includes('xlsxsource2.worker')}
      postMessage(message,...args){if(this.isConditional&&message?.op==='inspect')window.__sourceV2Requested++;return super.postMessage(message,...args)}
      addEventListener(type,listener,options){
        if(type!=='message'||typeof listener!=='function')return super.addEventListener(type,listener,options);
        const wrapped=event=>{const dispatch=()=>listener.call(this,event);if(window.__delaySourceReplies&&event.data?.op==='inspect'&&(!window.__delayConditionalOnly||this.isConditional))setTimeout(dispatch,1500);else dispatch()};
        this.listeners.set(listener,wrapped);return super.addEventListener(type,wrapped,options);
      }
      removeEventListener(type,listener,options){return super.removeEventListener(type,this.listeners.get(listener)||listener,options)}
    };
  }` })
  await cdp.send('Emulation.setDeviceMetricsOverride', { width: 1440, height: 1100, deviceScaleFactor: 1, mobile: false })
  await cdp.send('Page.navigate', { url: `${server.url}#/sheets?feature=native` })
  await poll(() => evaluate(`Boolean(document.querySelector('[aria-label="Open an XLSX file"]'))`), 'XLSX workbench')
  await upload(valid)
  await poll(() => evaluate(`Array.from(document.querySelectorAll('button')).some(b=>b.textContent.trim()==='Save to XLSX')`), 'editable initial workbook', 45000)
  await upload(source)
  await poll(() => evaluate(`Boolean(${panel})`), 'strict refusal and recovery option', 45000)
  await assert(`!Array.from(document.querySelectorAll('button')).some(b=>b.textContent.trim()==='Save to XLSX')`, 'old editing session cleared on refused source')
  await assert(`!document.querySelector('[aria-label="Mutation controls"]')`, 'recovery replaces the empty editing workspace')
  await click('Preview source styles')
  await poll(() => evaluate(`Boolean(${grid})`), 'actual source-style WASM result', 45000)
  await assert(`${panel}.textContent.includes(${JSON.stringify(sourceHash)})`, 'exact source package identity')
  await assert(`${grid}.querySelector('[data-source-cell="A1"]').colSpan===2&&!${grid}.querySelector('[data-source-cell="B1"]')`, 'one merged source anchor')
  await assert(`getComputedStyle(${grid}.querySelector('[data-source-cell="A1"]')).backgroundColor==='rgb(30, 39, 97)'&&getComputedStyle(${grid}.querySelector('[data-source-cell="A1"]')).color==='rgb(255, 255, 255)'`, 'source fill and foreground colors')
  if (conditional) {
    await assert(`${grid}.querySelector('[data-source-cell="A2"]').textContent==='327'&&${grid}.querySelector('[data-source-cell="B2"]').textContent==='REORDER'&&window.__sourceV2Requested===1`, 'actual V2 fallback and saved values')
    await assert(`getComputedStyle(${grid}.querySelector('[data-source-cell="B2"]')).backgroundColor==='rgb(255, 199, 206)'&&getComputedStyle(${grid}.querySelector('[data-source-cell="B2"]')).color==='rgb(156, 0, 6)'&&getComputedStyle(${grid}.querySelector('[data-source-cell="B2"]')).fontWeight==='700'`, 'conditional status colors and bold text')
    await assert(`(()=>{const bar=${grid}.querySelector('[data-source-bar="A2"]');return bar?.dataset.sourceBarPercent==='62.32'&&getComputedStyle(bar).backgroundImage.includes('99, 142, 198')&&Math.abs(bar.getBoundingClientRect().width/bar.parentElement.getBoundingClientRect().width*100-62.32)<0.1})()`, 'source-qualified proportional gradient bar')
  } else {
  await assert(`${grid}.querySelector('[data-source-cell="A2"]').textContent==='3.00 €'&&${panel}.textContent.includes('SUM(1,2)')`, 'source cache and declared currency suffix')
  }
  await assert(`${panel}.textContent.includes('absent applyFill')&&${panel}.textContent.includes('absent applyNumberFormat')&&${panel}.textContent.includes('may be stale')`, 'conflicts and cache warning visible')
  await assert(`!${panel}.querySelector('input,[contenteditable],a[download]')&&window.__sourcePosts===0`, 'preview is read-only and uploads nothing')
  await evaluate(`${grid}.scrollIntoView({block:'center'})`)
  const shot=await cdp.send('Page.captureScreenshot',{format:'png'});writeFileSync(resolve(artifacts,conditional?'xlsx-source-conditional.png':'xlsx-source-style.png'),Buffer.from(shot.data,'base64'))
  await evaluate(conditional ? 'window.__delaySourceReplies=true;window.__delayConditionalOnly=true' : 'window.__delaySourceReplies=true')
  let requested = await evaluate('window.__sourceV2Requested')
  await click('Preview source styles')
  if(conditional)await poll(async()=>await evaluate('window.__sourceV2Requested')>requested,'pending V2 cancellation')
  await click('Cancel preview')
  await new Promise(resolve => setTimeout(resolve,1800))
  await assert(`!${grid}&&${panel}.textContent.includes('cancelled')`, 'late cancelled result cannot remount')
  requested = await evaluate('window.__sourceV2Requested')
  await click('Preview source styles')
  if(conditional)await poll(async()=>await evaluate('window.__sourceV2Requested')>requested,'pending V2 source replacement')
  await poll(() => evaluate(`${panel}.textContent.includes('Reading source')`), 'pending replacement preview')
  await upload(valid)
  await poll(() => evaluate(`!${panel}&&Array.from(document.querySelectorAll('button')).some(b=>b.textContent.trim()==='Save to XLSX')`), 'new authoritative source', 45000)
  await new Promise(resolve => setTimeout(resolve,1800))
  await assert(`!${panel}&&window.__sourcePosts===0`, 'late old-source result cannot return after replacement')
  if ('sha256:'+createHash('sha256').update(readFileSync(source)).digest('hex')!==sourceHash) throw Error('Synthetic source bytes changed')
  if(errors.length)throw Error(errors.join('\n'))
  writeFileSync(resolve(artifacts,'summary.json'),JSON.stringify({status:'passed',checks,sourceHash,uploads:0,profile:conditional?'synthetic conditional source grid':'synthetic read-only grid'},null,2))
  console.log(`XLSX source-style browser checks passed: ${checks}`)
} finally {
  cdp?.close()
  try {if(chrome)await terminateProcess(chrome.child)} finally {
    await server?.close()
    for(const profile of profiles)rmSync(profile,{recursive:true,force:true,maxRetries:10,retryDelay:100})
    rmSync(scratch,{recursive:true,force:true,maxRetries:10,retryDelay:100})
  }
}
function run(command,args,cwd,env={}){const result=spawnSync(command,args,{cwd,env:{...process.env,...env},encoding:'utf8',timeout:120000});if(result.status!==0)throw Error(result.stderr+'\n'+result.stdout)}
async function upload(path){const handle=await cdp.send('Runtime.evaluate',{expression:`document.querySelector('[aria-label="Open an XLSX file"]')`});const node=await cdp.send('DOM.describeNode',{objectId:handle.result.objectId});await cdp.send('DOM.setFileInputFiles',{backendNodeId:node.node.backendNodeId,files:[path]});await cdp.send('Runtime.releaseObject',{objectId:handle.result.objectId})}
async function evaluate(expression){const value=await cdp.send('Runtime.evaluate',{expression,awaitPromise:true,returnByValue:true});if(value.exceptionDetails)throw Error(value.exceptionDetails.exception?.description??value.exceptionDetails.text);return value.result.value}
async function assert(expression,label){if(!await evaluate(expression))throw Error('Failed: '+label);checks++}
async function poll(check,label,timeout=30000){const end=Date.now()+timeout;while(Date.now()<end){if(await check())return;await new Promise(resolve=>setTimeout(resolve,75))}throw Error('Timed out: '+label)}
async function click(label) {
  const handle = await cdp.send('Runtime.evaluate', { expression: 'globalThis' })
  try {
    await poll(async () => {
      const value = await cdp.send('Runtime.callFunctionOn', {
        objectId: handle.result.objectId,
        functionDeclaration: `function(label) {
          const panel = document.querySelector('[aria-label="Read-only source-style recovery"]');
          const button = Array.from(panel?.querySelectorAll('button') ?? []).find(button => button.textContent.trim() === label);
          if (!button || button.disabled) return false;
          button.click(); return true;
        }`,
        arguments: [{ value: label }], returnByValue: true,
      })
      if (value.exceptionDetails) throw Error(value.exceptionDetails.exception?.description ?? value.exceptionDetails.text)
      return value.result.value
    }, label)
  } finally { await cdp.send('Runtime.releaseObject', { objectId: handle.result.objectId }) }
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
