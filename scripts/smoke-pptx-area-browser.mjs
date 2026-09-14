import {spawn,spawnSync} from 'node:child_process'
import {mkdtempSync,readFileSync,writeFileSync,existsSync,rmSync,mkdirSync} from 'node:fs'
import {createHash} from 'node:crypto'
import {createServer} from 'node:net'
import {tmpdir} from 'node:os'
import {resolve} from 'node:path'
import {launchChromeForCDP,terminateProcess} from './chrome-cdp-startup.mjs'
import {startShowcaseDevServer} from './showcase-smoke-dev-server.mjs'
const root=resolve(import.meta.dirname,'..'),scratch=mkdtempSync(resolve(tmpdir(),'injoffice-area-browser-'))
const artifacts=process.env.SHOWCASE_OUTPUT?resolve(process.env.SHOWCASE_OUTPUT):mkdtempSync(resolve(tmpdir(),'injoffice-area-evidence-'));mkdirSync(artifacts,{recursive:true})
const profiles=[],errors=[],section=`document.querySelector('[aria-label="Measured native presentation"]')`,literal=`document.querySelector('[aria-label="Source literal chart preview"]')`,proof=[]
let helper,server,chrome,cdp
const hash=bytes=>createHash('sha256').update(bytes).digest('hex')
function command(cmd,args,cwd=root,env={}){const r=spawnSync(cmd,args,{cwd,env:{...process.env,...env},encoding:'utf8',timeout:90000});if(r.status!==0)throw new Error(`${cmd}: ${r.stderr}\n${r.stdout}`)}
try{
 const worker=resolve(root,'apps/pptx-page-paint-worker/dist/worker.js');if(!existsSync(worker))throw new Error('Build workspace packages first')
 command('go',['test','-run','^TestNativeAreaBrowserFixtures$','-count=1','.'],resolve(root,'go/pptxpatch'),{INJOFFICE_PPTX_AREA_FIXTURES:artifacts})
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
 chrome=await launchChromeForCDP({executable,createProfile:()=>{const path=mkdtempSync(resolve(tmpdir(),'injoffice-area-chrome-'));profiles.push(path);return path}})
 cdp=await connect(chrome.target.webSocketDebuggerUrl)
 await cdp.send('Runtime.enable');await cdp.send('Page.enable');await cdp.send('Emulation.setDeviceMetricsOverride',{width:1440,height:1100,deviceScaleFactor:1,mobile:false})
 await cdp.send('Page.addScriptToEvaluateOnNewDocument',{source:`{const original=fetch;window.__areaPosts=[];window.__areaResponses=[];window.fetch=async(input,init)=>{if(init?.method==='POST'&&String(input).includes('/v1/')){const bytes=await init.body.arrayBuffer();const digest=await crypto.subtle.digest('SHA-256',bytes);window.__areaPosts.push({url:String(input),hash:[...new Uint8Array(digest)].map(n=>n.toString(16).padStart(2,'0')).join('')})}const response=await original(input,init);if(String(input).includes('/v1/pptx/slide-preview'))window.__areaResponses.push(await response.clone().json());return response}}`})
 await cdp.send('Page.navigate',{url:`${server.url}#/slides?feature=pptx-native`})
 await poll(()=>evaluate(`!!document.querySelector('input[type=file]')`),'file input')
 for(const [index,name] of ['standard','stacked','percent','zero','large','labels','negative-stack'].entries()){
  console.log('Checking area '+name)
  const fixture=resolve(artifacts,name+'.pptx'),digest=hash(readFileSync(fixture));await upload(fixture)
  await poll(()=>evaluate(`${section}?.textContent.includes('Nothing is uploaded')&&${section}?.querySelector('svg')===null`),'replacement resets output and consent')
  await assert(`window.__areaPosts.length===${index}`,'opening/replacing never submits source')
  await poll(()=>evaluate(`!!${literal}`),'actual WASM source inspection')
  await assert(`${literal}.querySelector('svg')===null&&!${literal}.querySelector('input').checked`,'literal preview defaults off on source replacement')
  await evaluate(`${literal}.querySelector('input').click()`)
  if(name==='negative-stack')await assert(`${literal}.textContent.includes('outside the supported')`,'unsupported negative stack stays opaque')
  else if(name==='labels')await assert(`${literal}.textContent.includes('supplied-font')`,'labels require measured font preview')
  else {
   await poll(()=>evaluate(`!!${literal}.querySelector('svg')`),'literal area visible')
   if(name==='large')await assert(`(${literal}.querySelector('svg path[fill="#123456"]').getAttribute('d').match(/[MLZ]/g)||[]).length===1280`,'large source has full compound path')
   if(name==='percent')await assert(`${literal}.textContent.includes('9007199254740993')&&${literal}.textContent.includes('1e-100')`,'exact source lexemes survive WASM')
  }
  await evaluate(`[...${section}.querySelectorAll('label')].find(label=>label.textContent.includes('Preview source literal charts')).querySelector('input').click()`)
  await clickRender()
  await poll(()=>evaluate(`window.__areaResponses.length===${index+1}&&!!${section}.querySelector('svg')`),'actual server/worker native replay',45000)
  await assert(`window.__areaPosts[${index}].hash===${JSON.stringify(digest)}&&window.__areaPosts[${index}].url.includes('charts=source-literal')`,'source identity and explicit chart mode')
  const response=await evaluate(`window.__areaResponses[${index}]`);writeFileSync(resolve(artifacts,name+'-preview.json'),JSON.stringify(response,null,2))
  const nodes=[],dataPaths=[];const visit=(list,text=false)=>{for(const n of list){nodes.push(n);if(n.kind==='path'&&!text)dataPaths.push(n);if(n.kind==='group')visit(n.children,text||n.sourceRole==='contentRun'||n.sourceRole==='paragraphBullet')}};visit(response.nodes)
  const fills=dataPaths.filter(n=>n.fill!=='none'),expected=name==='negative-stack'?0:name==='zero'?0:name==='standard'||name==='stacked'||name==='percent'?2:1
  const areaFills=fills.filter(n=>n.fill==='123456'||n.fill==='CC5500'||n.fill==='ABCDEF')
  if(areaFills.length!==expected)throw new Error(`${name}: expected${expected} area fills, got${areaFills.length}`)
  if(name==='large'&&(areaFills[0].d.match(/[MLZ]/g)||[]).length!==1280)throw new Error('worker truncated large compound fill')
  if(name==='labels'&&!nodes.some(n=>n.sourceRole==='contentRun'))throw new Error('missing actual glyph outlines')
  if(name==='zero'&&nodes.filter(n=>n.kind==='path'&&n.stroke==='111111').length!==2)throw new Error('zero bands lost independent axes')
  if(name!=='negative-stack'&&!response.diagnostics.some(d=>d.includes('chart.literalAreaPreview')))throw new Error('missing area disclosure')
  await evaluate(`${section}.scrollIntoView()`);const shot=await cdp.send('Page.captureScreenshot',{format:'png'});writeFileSync(resolve(artifacts,name+'.png'),Buffer.from(shot.data,'base64'))
  if(hash(readFileSync(fixture))!==digest)throw new Error('original file changed')
  proof.push({name,packageSHA256:digest,areaFills:areaFills.length,sourcePreserved:true})
 }
 if(errors.length)throw new Error(errors.join('\n'))
 writeFileSync(resolve(artifacts,'proof.json'),JSON.stringify(proof,null,2));console.log(JSON.stringify({result:'PASS',artifacts,cases:proof},null,2))
}catch(error){if(cdp){try{writeFileSync(resolve(artifacts,'failure.txt'),await evaluate('document.body.innerText'));const shot=await cdp.send('Page.captureScreenshot',{format:'png'});writeFileSync(resolve(artifacts,'failure.png'),Buffer.from(shot.data,'base64'))}catch{}}console.error('Area evidence: '+artifacts);throw error}finally{cdp?.close();if(chrome)await terminateProcess(chrome.child);await server?.close();if(helper)await terminateProcess(helper);for(const profile of profiles)rmSync(profile,{recursive:true,force:true,maxRetries:10,retryDelay:100});rmSync(scratch,{recursive:true,force:true})}
async function unusedPort(){const server=createServer();await new Promise(done=>server.listen(0,'127.0.0.1',done));const port=server.address().port;await new Promise(done=>server.close(done));return port}
async function evaluate(expression){const r=await cdp.send('Runtime.evaluate',{expression,awaitPromise:true,returnByValue:true});if(r.exceptionDetails)throw new Error(r.exceptionDetails.exception?.description||r.exceptionDetails.text);return r.result.value}
async function assert(expression,label){if(!await evaluate(expression))throw new Error(label)}
async function poll(check,label,timeout=30000){const end=Date.now()+timeout;while(Date.now()<end){if(await check())return;await new Promise(done=>setTimeout(done,75))}throw new Error(`Timed out: ${label}`)}
async function upload(path){const handle=await cdp.send('Runtime.evaluate',{expression:`document.querySelector('input[type=file]')`});const node=await cdp.send('DOM.describeNode',{objectId:handle.result.objectId});await cdp.send('DOM.setFileInputFiles',{backendNodeId:node.node.backendNodeId,files:[path]})}
async function clickRender(){await evaluate(`[...${section}.querySelectorAll('button')].find(b=>b.textContent==='Upload to helper and render native slide').click()`)}
async function connect(url){const socket=new WebSocket(url);await new Promise((done,reject)=>{socket.addEventListener('open',done,{once:true});socket.addEventListener('error',reject,{once:true})});let next=0;const waiting=new Map();socket.addEventListener('message',event=>{const message=JSON.parse(event.data);if(message.id){const p=waiting.get(message.id);waiting.delete(message.id);message.error?p?.reject(new Error(message.error.message)):p?.resolve(message.result)}else if(message.method==='Runtime.exceptionThrown')errors.push(message.params.exceptionDetails.text)});return{send(method,params={}){return new Promise((resolve,reject)=>{const id=++next;waiting.set(id,{resolve,reject});socket.send(JSON.stringify({id,method,params}))})},close(){socket.close()}}}
