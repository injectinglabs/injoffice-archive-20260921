import {PNG} from 'pngjs'
import {spawn,spawnSync} from 'node:child_process'
import {mkdtempSync,readFileSync,writeFileSync,existsSync,rmSync,mkdirSync} from 'node:fs'
import {createHash} from 'node:crypto'
import {createServer} from 'node:net'
import {tmpdir} from 'node:os'
import {resolve,dirname,delimiter} from 'node:path'
import {launchChromeForCDP,terminateProcess} from './chrome-cdp-startup.mjs'
import {nativePptxSvgPath} from '../apps/playground/src/components/nativePptxSvgUnits.ts'
import {startShowcaseDevServer} from './showcase-smoke-dev-server.mjs'
// The helper launches `node`; inherit the exact runtime running this harness.
process.env.PATH=dirname(process.execPath)+delimiter+(process.env.PATH??'')
const root=resolve(import.meta.dirname,'..'),scratch=mkdtempSync(resolve(tmpdir(),'injoffice-series-order-browser-'))
const artifacts=process.env.SHOWCASE_OUTPUT?resolve(process.env.SHOWCASE_OUTPUT):mkdtempSync(resolve(tmpdir(),'injoffice-series-order-evidence-'));mkdirSync(artifacts,{recursive:true})
const profiles=[],errors=[],section=`document.querySelector('[aria-label="Measured native presentation"]')`,proof=[]
let helper,server,chrome,cdp
const names=['bar','line','scatter','area','bubble','workbook-bar','workbook-line','workbook-scatter','workbook-bubble'].flatMap(name=>[name+'-chart-only',name+'-chart-only-labels'])
const hash=bytes=>createHash('sha256').update(bytes).digest('hex')
function command(cmd,args,cwd=root,env={}){const r=spawnSync(cmd,args,{cwd,env:{...process.env,...env},encoding:'utf8',timeout:90000});if(r.status!==0)throw new Error(`${cmd}: ${r.stderr}\n${r.stdout}`)}
try{
 const worker=resolve(root,'apps/pptx-page-paint-worker/dist/worker.js');if(!existsSync(worker))throw new Error('Build workspace packages first')
 command('go',['test','-run','^TestNativeSeriesOrderBrowserFixtures$','-count=1','.'],resolve(root,'go/pptxpatch'),{INJOFFICE_PPTX_SERIES_ORDER_FIXTURES:artifacts})
 const font=resolve(root,'node_modules/dejavu-fonts-ttf/ttf/DejaVuSans.ttf'),manifest=resolve(scratch,'fonts.json')
 writeFileSync(manifest,JSON.stringify({version:1,faces:[{family:'DejaVu Sans',weight:400,style:'normal',path:font,sha256:`sha256:${hash(readFileSync(font))}`}]}))
 const binary=resolve(scratch,'injoffice-server');command('go',['build','-o',binary,'./cmd/injoffice-server'],resolve(root,'go/injoffice-server'))
 const port=await unusedPort();helper=spawn(binary,['-addr',`127.0.0.1:${port}`,'-artifacts',resolve(scratch,'artifacts'),'-pptx-preview-worker',worker,'-pptx-font-manifest',manifest],{stdio:['ignore','ignore','pipe']})
 let helperLog='';helper.stderr.on('data',data=>helperLog+=data)
 await poll(async()=>{try{return(await fetch(`http://127.0.0.1:${port}/healthz`)).ok}catch{return false}},'helper startup')
 process.env.VITE_INJOFFICE_API_BASE='.';process.env.INJOFFICE_SERVER=`http://127.0.0.1:${port}`
 server=await startShowcaseDevServer(resolve(root,'apps/playground'))
 const executable=process.env.CHROME_PATH||['/Applications/Google Chrome.app/Contents/MacOS/Google Chrome','google-chrome','chromium'].find(path=>spawnSync(path,['--version'],{stdio:'ignore'}).status===0)
 if(!executable)throw new Error('Chrome required')
 chrome=await launchChromeForCDP({executable,createProfile:()=>{const path=mkdtempSync(resolve(tmpdir(),'injoffice-series-order-chrome-'));profiles.push(path);return path}})
 cdp=await connect(chrome.target.webSocketDebuggerUrl)
 await cdp.send('Runtime.enable');await cdp.send('Page.enable');await cdp.send('Emulation.setDeviceMetricsOverride',{width:1440,height:1100,deviceScaleFactor:1,mobile:false})
 await cdp.send('Page.addScriptToEvaluateOnNewDocument',{source:`{const original=fetch;window.__areaPosts=[];window.__areaResponses=[];window.fetch=async(input,init)=>{if(init?.method==='POST'&&String(input).includes('/v1/')){const bytes=await init.body.arrayBuffer();const digest=await crypto.subtle.digest('SHA-256',bytes);window.__areaPosts.push({url:String(input),hash:[...new Uint8Array(digest)].map(n=>n.toString(16).padStart(2,'0')).join('')})}const response=await original(input,init);if(String(input).includes('/v1/pptx/slide-preview'))window.__areaResponses.push(await response.clone().json());return response}}`})
 await cdp.send('Page.navigate',{url:`${server.url}#/slides?feature=pptx-native`})
 await poll(()=>evaluate(`!!document.querySelector('input[type=file]')`),'file input')
 for(const [index,name] of names.entries()){
  const workbook=name.startsWith('workbook-'),fixture=resolve(artifacts,name+'.pptx'),digest=hash(readFileSync(fixture));await upload(fixture)
  await poll(()=>evaluate(`${section}?.textContent.includes('Nothing is uploaded')&&${section}?.querySelector('svg')===null`),'replacement resets consent')
  await assert(`window.__areaPosts.length===${index}`,'opening never submits source')
  const label=workbook?'Preview charts from the embedded workbook':'Preview source literal charts'
  await evaluate(`[...${section}.querySelectorAll('label')].find(label=>label.textContent.includes(${JSON.stringify(label)})).querySelector('input').click()`)
  await clickRender()
  await poll(()=>evaluate(`window.__areaResponses.length===${index+1}&&!!${section}.querySelector('svg')`),'actual source to worker browser replay',45000)
  await assert(`window.__areaPosts[${index}].hash===${JSON.stringify(digest)}`,'unchanged upload source')
  const response=await evaluate(`window.__areaResponses[${index}]`),colors=[],expectedPaths=[]
  const visit=nodes=>{for(const n of nodes){if(n.kind==='path'){const color=n.fill==='none'?n.stroke:n.fill;if(['1E88E5','E53935','43A047'].includes(color))colors.push(color);if(['1E88E5','E53935','43A047','ABCDEF'].includes(color)){expectedPaths.push({fill:n.fill==='none'?'none':'#'+n.fill,stroke:n.stroke?'#'+n.stroke:null,d:nativePptxSvgPath(n.d)})}}if(n.kind==='group')visit(n.children)}};visit(response.nodes)
  if(JSON.stringify([...new Set(colors)])!==JSON.stringify(['1E88E5','E53935','43A047']))throw Error(name+': source painter sequence lost')
  const dom=await evaluate(`[...${section}.querySelector('svg[aria-label="Measured native slide 1"]').querySelectorAll('path')].map(n=>({fill:n.getAttribute('fill'),stroke:n.getAttribute('stroke'),d:n.getAttribute('d')})).filter(n=>['#1E88E5','#E53935','#43A047','#ABCDEF'].includes(n.fill==='none'?n.stroke:n.fill))`)
  if(JSON.stringify(dom.map(n=>(n.fill==='none'?n.stroke:n.fill).slice(1)).filter(c=>c!=='ABCDEF'))!==JSON.stringify(colors))throw Error(name+': production DOM path sequence/count differs from worker')
  if(JSON.stringify(dom)!==JSON.stringify(expectedPaths))throw Error(name+': production DOM changed point/series geometry or paint identity')
  if(name.endsWith('-labels')){
   if(!response.diagnostics.some(d=>d.includes('chart.axisLabelsPreview')))throw Error(name+': source labels refused')
   await assert(`!!${section}.querySelector('svg [data-native-source-role="contentRun"]')`,'source labels reach production component')
  }
  writeFileSync(resolve(artifacts,name+'-dom.json'),JSON.stringify(dom,null,2))
  writeFileSync(resolve(artifacts,name+'-preview.json'),JSON.stringify(response,null,2))
  const matchedSvg=`${section}.querySelector('svg[aria-label="Measured native slide 1"]')`
  await evaluate(`${matchedSvg}.scrollIntoView({block:'center',inline:'center',behavior:'instant'})`)
  await poll(()=>evaluate(`(()=>{const r=${matchedSvg}.getBoundingClientRect();return r.width>1&&r.height>1&&r.left>=0&&r.top>=0&&r.right<=innerWidth&&r.bottom<=innerHeight})()`),'verified native SVG entirely visible in viewport')
  const clip=await evaluate(`(()=>{const r=${matchedSvg}.getBoundingClientRect();return {x:r.left+scrollX,y:r.top+scrollY,width:r.width,height:r.height,scale:1}})()`)
  const shot=await cdp.send('Page.captureScreenshot',{format:'png',clip,captureBeyondViewport:true}),painted=Buffer.from(shot.data,'base64')
  writeFileSync(resolve(artifacts,name+'.png'),painted)
  // A same-rectangle control removes only source series paint, retaining axes,
  // labels and the slide background. Nonempty pixels must come from the chart.
  let control
  try{
   await evaluate(`{window.__seriesPixelControl=[...${matchedSvg}.querySelectorAll('path')].filter(n=>['#1E88E5','#E53935','#43A047','#ABCDEF'].includes(n.getAttribute('fill')==='none'?n.getAttribute('stroke'):n.getAttribute('fill'))).map(n=>[n,n.style.visibility]);for(const [n]of window.__seriesPixelControl)n.style.visibility='hidden'}`)
   control=Buffer.from((await cdp.send('Page.captureScreenshot',{format:'png',clip,captureBeyondViewport:true})).data,'base64')
  }finally{await evaluate(`{for(const [n,value]of window.__seriesPixelControl??[])n.style.visibility=value;delete window.__seriesPixelControl}`)}
  writeFileSync(resolve(artifacts,name+'-series-hidden-control.png'),control)
  const pixels=PNG.sync.read(painted),blank=PNG.sync.read(control)
  if(pixels.width!==blank.width||pixels.height!==blank.height||pixels.width<2||pixels.height<2)throw Error(name+': invalid chart-focused screenshot bounds')
  let changedPixels=0
  for(let p=0;p<pixels.data.length;p+=4)if(Math.max(...[0,1,2,3].map(c=>Math.abs(pixels.data[p+c]-blank.data[p+c])))>8)changedPixels++
  if(changedPixels<10)throw Error(name+': actual chart paint is visually empty against its series-hidden control')
  if(hash(readFileSync(fixture))!==digest)throw Error('source changed')
  proof.push({name,packageSHA256:digest,sourcePreserved:true,sequence:[2,0,1],runtime:process.execPath,domPathCount:dom.length,screenshotClip:clip,paintedSeriesPixels:changedPixels})
 }
 if(errors.length)throw new Error(errors.join('\n'))
 writeFileSync(resolve(artifacts,'proof.json'),JSON.stringify(proof,null,2));console.log(JSON.stringify({result:'PASS',artifacts,cases:proof},null,2))
}catch(error){if(cdp){try{writeFileSync(resolve(artifacts,'failure.txt'),await evaluate('document.body.innerText'));const shot=await cdp.send('Page.captureScreenshot',{format:'png'});writeFileSync(resolve(artifacts,'failure.png'),Buffer.from(shot.data,'base64'))}catch{}}console.error('Series-order evidence: '+artifacts);throw error}finally{cdp?.close();if(chrome)await terminateProcess(chrome.child);await server?.close();if(helper)await terminateProcess(helper);for(const profile of profiles)rmSync(profile,{recursive:true,force:true,maxRetries:10,retryDelay:100});rmSync(scratch,{recursive:true,force:true})}
async function unusedPort(){const server=createServer();await new Promise(done=>server.listen(0,'127.0.0.1',done));const port=server.address().port;await new Promise(done=>server.close(done));return port}
async function evaluate(expression){const r=await cdp.send('Runtime.evaluate',{expression,awaitPromise:true,returnByValue:true});if(r.exceptionDetails)throw new Error(r.exceptionDetails.exception?.description||r.exceptionDetails.text);return r.result.value}
async function assert(expression,label){if(!await evaluate(expression))throw new Error(label)}
async function poll(check,label,timeout=30000){const end=Date.now()+timeout;while(Date.now()<end){if(await check())return;await new Promise(done=>setTimeout(done,75))}throw new Error(`Timed out: ${label}`)}
async function upload(path){const handle=await cdp.send('Runtime.evaluate',{expression:`document.querySelector('input[type=file]')`});const node=await cdp.send('DOM.describeNode',{objectId:handle.result.objectId});await cdp.send('DOM.setFileInputFiles',{backendNodeId:node.node.backendNodeId,files:[path]})}
async function clickRender(){await evaluate(`[...${section}.querySelectorAll('button')].find(b=>b.textContent==='Upload to helper and render native slide').click()`)}
async function connect(url){const socket=new WebSocket(url);await new Promise((done,reject)=>{socket.addEventListener('open',done,{once:true});socket.addEventListener('error',reject,{once:true})});let next=0;const waiting=new Map();socket.addEventListener('message',event=>{const message=JSON.parse(event.data);if(message.id){const p=waiting.get(message.id);waiting.delete(message.id);message.error?p?.reject(new Error(message.error.message)):p?.resolve(message.result)}else if(message.method==='Runtime.exceptionThrown')errors.push(message.params.exceptionDetails.text)});return{send(method,params={}){return new Promise((resolve,reject)=>{const id=++next;waiting.set(id,{resolve,reject});socket.send(JSON.stringify({id,method,params}))})},close(){socket.close()}}}
