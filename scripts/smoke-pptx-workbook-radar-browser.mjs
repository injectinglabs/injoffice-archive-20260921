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
const root=resolve(import.meta.dirname,'..'),scratch=mkdtempSync(resolve(tmpdir(),'injoffice-workbook-radar-browser-'))
const artifacts=process.env.SHOWCASE_OUTPUT?resolve(process.env.SHOWCASE_OUTPUT):mkdtempSync(resolve(tmpdir(),'injoffice-workbook-radar-evidence-'));mkdirSync(artifacts,{recursive:true})
const profiles=[],errors=[],section=`document.querySelector('[aria-label="Measured native presentation"]')`,proof=[]
let helper,server,chrome,cdp
const names=[false,true].flatMap(strict=>[false,true].flatMap(filled=>['plain','rotated','grouped','reversed'].map(mode=>`radar-${strict}-${filled}-${mode}`)))
const hash=bytes=>createHash('sha256').update(bytes).digest('hex')
function command(cmd,args,cwd=root,env={}){const r=spawnSync(cmd,args,{cwd,env:{...process.env,...env},encoding:'utf8',timeout:90000});if(r.status!==0)throw new Error(`${cmd}: ${r.stderr}\n${r.stdout}`)}
try{
 const worker=resolve(root,'apps/pptx-page-paint-worker/dist/worker.js');if(!existsSync(worker))throw new Error('Build workspace packages first')
 command('go',['test','-run','^TestNativeWorkbookRadarFixtures$','-count=1','.'],resolve(root,'go/pptxpatch'),{INJOFFICE_PPTX_WORKBOOK_RADAR_FIXTURES:artifacts})
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
 chrome=await launchChromeForCDP({executable,createProfile:()=>{const path=mkdtempSync(resolve(tmpdir(),'injoffice-workbook-radar-chrome-'));profiles.push(path);return path}})
 cdp=await connect(chrome.target.webSocketDebuggerUrl)
 await cdp.send('Runtime.enable');await cdp.send('Page.enable');await cdp.send('Emulation.setDeviceMetricsOverride',{width:1440,height:1100,deviceScaleFactor:1,mobile:false})
 await cdp.send('Page.addScriptToEvaluateOnNewDocument',{source:`{const original=fetch;window.__areaPosts=[];window.__areaResponses=[];window.fetch=async(input,init)=>{if(init?.method==='POST'&&String(input).includes('/v1/')){const bytes=await init.body.arrayBuffer();const digest=await crypto.subtle.digest('SHA-256',bytes);window.__areaPosts.push({url:String(input),hash:[...new Uint8Array(digest)].map(n=>n.toString(16).padStart(2,'0')).join('')})}const response=await original(input,init);if(String(input).includes('/v1/pptx/slide-preview'))window.__areaResponses.push(await response.clone().json());return response}}`})
 await cdp.send('Page.navigate',{url:`${server.url}#/slides?feature=pptx-native`})
 await poll(()=>evaluate(`!!document.querySelector('input[type=file]')`),'file input')
 for(const [index,name] of names.entries()){
  const fixture=resolve(artifacts,name+'.pptx'),digest=hash(readFileSync(fixture));await upload(fixture)
  await poll(()=>evaluate(`${section}?.textContent.includes('Nothing is uploaded')&&${section}?.querySelector('svg')===null`),'replacement resets consent')
  await assert(`window.__areaPosts.length===${index}`,'opening never submits source')
  const label='Preview charts from the embedded workbook'
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
  if(response.workbook_chart_preview!==true||response.source_chart_preview===true||!response.diagnostics.some(d=>d.includes('chart.workbookRadarPreview')))throw Error(name+': radar source route unavailable')
  writeFileSync(resolve(artifacts,name+'-dom.json'),JSON.stringify(dom,null,2))
  writeFileSync(resolve(artifacts,name+'-preview.json'),JSON.stringify(response,null,2))
  const region=await evaluate(`(async()=>{const svg=${section}.querySelector('svg[aria-label="Measured native slide 1"]');svg.scrollIntoView({block:'center'});await new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r)));const b=svg.getBoundingClientRect();if(b.width<100||b.height<100||b.left<0||b.top<0||b.right>innerWidth||b.bottom>innerHeight)throw Error('Measured slide is not visible');const paths=[...svg.querySelectorAll('path')].filter(p=>['#1E88E5','#E53935','#43A047'].includes(p.getAttribute('stroke')));if(paths.length!==3||paths.some(p=>{const r=p.getBoundingClientRect();return r.width<1||r.height<1}))throw Error('Degenerate radar path');const minimums=paths.map(p=>{const r=p.getBoundingClientRect(),m=p.getScreenCTM(),scale=Math.min(Math.hypot(m.a,m.b),Math.hypot(m.c,m.d)),area=p.getAttribute('fill')==='none'?p.getTotalLength()*scale*Number(p.getAttribute('stroke-width'))*scale:r.width*r.height/2;return Math.max(2,Math.floor(area*0.05))});return {clip:{x:b.left+scrollX,y:b.top+scrollY,width:b.width,height:b.height,scale:1},minimums}})()`)
  const shot=await cdp.send('Page.captureScreenshot',{format:'png',clip:region.clip,captureBeyondViewport:true});writeFileSync(resolve(artifacts,name+'.png'),Buffer.from(shot.data,'base64'))
  const pixels=await evaluate(`(async()=>{const img=new Image();img.src='data:image/png;base64,${shot.data}';await img.decode();const canvas=document.createElement('canvas');canvas.width=img.width;canvas.height=img.height;const ctx=canvas.getContext('2d');ctx.drawImage(img,0,0);const data=ctx.getImageData(0,0,img.width,img.height).data,counts=[0,0,0],colors=[[30,136,229],[229,57,53],[67,160,71]];for(let i=0;i<data.length;i+=4)for(let j=0;j<3;j++)if(colors[j].every((v,k)=>Math.abs(data[i+k]-v)<=12))counts[j]++;return {width:img.width,height:img.height,counts}})()`)
  const filled=name.split('-')[2]==='true'
  const controls=[]
  for(const index of filled?[2]:[0,1,2]){
   const color=['#1E88E5','#E53935','#43A047'][index]
   const prior=await evaluate(`(()=>{const s=${section}.querySelector('svg path[stroke="${color}"]').style;return {value:s.getPropertyValue('visibility'),priority:s.getPropertyPriority('visibility')}})()`)
   try{
   await evaluate(`(async()=>{const path=${section}.querySelector('svg path[stroke="${color}"]');path.style.visibility='hidden';await new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r)))})()`)
   const control=await cdp.send('Page.captureScreenshot',{format:'png',clip:region.clip,captureBeyondViewport:true})
   writeFileSync(resolve(artifacts,name+'-control-'+index+'.png'),Buffer.from(control.data,'base64'))
   const changed=await evaluate(`(async()=>{const read=async source=>{const img=new Image();img.src=source;await img.decode();const c=document.createElement('canvas');c.width=img.width;c.height=img.height;const ctx=c.getContext('2d');ctx.drawImage(img,0,0);return ctx.getImageData(0,0,c.width,c.height).data};const a=await read('data:image/png;base64,${shot.data}'),b=await read('data:image/png;base64,${control.data}');if(a.length!==b.length)throw Error('Control dimensions changed');let changed=0;for(let i=0;i<a.length;i+=4)if(Math.abs(a[i]-b[i])+Math.abs(a[i+1]-b[i+1])+Math.abs(a[i+2]-b[i+2])>24)changed++;return changed})()`)

   if(changed<region.minimums[index])throw Error(name+': hiding source series did not change expected visible pixels '+JSON.stringify({index,changed,minimum:region.minimums[index]}))
   controls.push({series:index,changed,minimum:region.minimums[index]})
   }finally{
    await evaluate(`(async()=>{const s=${section}.querySelector('svg path[stroke="${color}"]').style,prior=${JSON.stringify(prior)};if(prior.value)s.setProperty('visibility',prior.value,prior.priority);else s.removeProperty('visibility');await new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r)))})()`)
   }
  }
  writeFileSync(resolve(artifacts,name+'-pixels.json'),JSON.stringify({region,pixels,controls},null,2))
  if(hash(readFileSync(fixture))!==digest)throw Error('source changed')
  proof.push({name,packageSHA256:digest,sourcePreserved:true,sequence:[2,0,1],runtime:process.execPath,domPathCount:dom.length,pixels,controls})
 }
 if(errors.length)throw new Error(errors.join('\n'))
 writeFileSync(resolve(artifacts,'proof.json'),JSON.stringify(proof,null,2));console.log(JSON.stringify({result:'PASS',artifacts,cases:proof},null,2))
}catch(error){if(cdp){try{writeFileSync(resolve(artifacts,'failure.txt'),await evaluate('document.body.innerText'));const shot=await cdp.send('Page.captureScreenshot',{format:'png'});writeFileSync(resolve(artifacts,'failure.png'),Buffer.from(shot.data,'base64'))}catch{}}console.error('Workbook-radar evidence: '+artifacts);throw error}finally{cdp?.close();if(chrome)await terminateProcess(chrome.child);await server?.close();if(helper)await terminateProcess(helper);for(const profile of profiles)rmSync(profile,{recursive:true,force:true,maxRetries:10,retryDelay:100});rmSync(scratch,{recursive:true,force:true})}
async function unusedPort(){const server=createServer();await new Promise(done=>server.listen(0,'127.0.0.1',done));const port=server.address().port;await new Promise(done=>server.close(done));return port}
async function evaluate(expression){const r=await cdp.send('Runtime.evaluate',{expression,awaitPromise:true,returnByValue:true});if(r.exceptionDetails)throw new Error(r.exceptionDetails.exception?.description||r.exceptionDetails.text);return r.result.value}
async function assert(expression,label){if(!await evaluate(expression))throw new Error(label)}
async function poll(check,label,timeout=30000){const end=Date.now()+timeout;while(Date.now()<end){if(await check())return;await new Promise(done=>setTimeout(done,75))}throw new Error(`Timed out: ${label}`)}
async function upload(path){const handle=await cdp.send('Runtime.evaluate',{expression:`document.querySelector('input[type=file]')`});const node=await cdp.send('DOM.describeNode',{objectId:handle.result.objectId});await cdp.send('DOM.setFileInputFiles',{backendNodeId:node.node.backendNodeId,files:[path]})}
async function clickRender(){await evaluate(`[...${section}.querySelectorAll('button')].find(b=>b.textContent==='Upload to helper and render native slide').click()`)}
async function connect(url){const socket=new WebSocket(url);await new Promise((done,reject)=>{socket.addEventListener('open',done,{once:true});socket.addEventListener('error',reject,{once:true})});let next=0;const waiting=new Map();socket.addEventListener('message',event=>{const message=JSON.parse(event.data);if(message.id){const p=waiting.get(message.id);waiting.delete(message.id);message.error?p?.reject(new Error(message.error.message)):p?.resolve(message.result)}else if(message.method==='Runtime.exceptionThrown')errors.push(message.params.exceptionDetails.text)});return{send(method,params={}){return new Promise((resolve,reject)=>{const id=++next;waiting.set(id,{resolve,reject});socket.send(JSON.stringify({id,method,params}))})},close(){socket.close()}}}
