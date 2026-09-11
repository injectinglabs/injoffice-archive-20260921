import {spawn,spawnSync} from 'node:child_process'
import {mkdtempSync,readFileSync,writeFileSync,existsSync,rmSync,mkdirSync} from 'node:fs'
import {createHash} from 'node:crypto'
import {createServer} from 'node:net'
import {tmpdir} from 'node:os'
import {resolve} from 'node:path'
import {launchChromeForCDP,terminateProcess} from './chrome-cdp-startup.mjs'
import {startShowcaseDevServer} from './showcase-smoke-dev-server.mjs'
import {PNG} from 'pngjs'
const root=resolve(import.meta.dirname,'..'),scratch=mkdtempSync(resolve(tmpdir(),'injoffice-pptx-native-'))
const artifacts=process.env.SHOWCASE_OUTPUT?resolve(process.env.SHOWCASE_OUTPUT):mkdtempSync(resolve(tmpdir(),'injoffice-pptx-native-evidence-'));mkdirSync(artifacts,{recursive:true})
const profiles=[],errors=[],section=`document.querySelector('[aria-label="Measured native presentation"]')`
let helper,server,chrome,cdp
const hash=bytes=>createHash('sha256').update(bytes).digest('hex')
function command(cmd,args,cwd=root,env={}){const r=spawnSync(cmd,args,{cwd,env:{...process.env,...env},encoding:'utf8',timeout:90000});if(r.status!==0)throw new Error(`${cmd}: ${r.stderr}\n${r.stdout}`)}
try{
 const worker=resolve(root,'apps/pptx-page-paint-worker/dist/worker.js');if(!existsSync(worker))throw new Error('Build workspace packages first')
 const fixture=resolve(scratch,'mixed-anchors.pptx')
 command('go',['test','-run','^(TestNativeMeasuredPreviewBrowserFixture|TestNativeExplicitPromptParagraphDefaults)$','-count=1','.'],resolve(root,'go/pptxpatch'),{INJOFFICE_PPTX_PREVIEW_FIXTURE:fixture})
 const sourceHash=hash(readFileSync(fixture)),font=resolve(root,'node_modules/dejavu-fonts-ttf/ttf/DejaVuSans.ttf'),manifest=resolve(scratch,'fonts.json')
 writeFileSync(manifest,JSON.stringify({version:1,faces:[{family:'DejaVu Sans',weight:400,style:'normal',path:font,sha256:`sha256:${hash(readFileSync(font))}`}]}))
 const binary=resolve(scratch,'injoffice-server');command('go',['build','-o',binary,'./cmd/injoffice-server'],resolve(root,'go/injoffice-server'))
 const port=await unusedPort();helper=spawn(binary,['-addr',`127.0.0.1:${port}`,'-artifacts',resolve(scratch,'artifacts'),'-pptx-preview-worker',worker,'-pptx-font-manifest',manifest],{stdio:['ignore','ignore','pipe']})
 let helperLog='';helper.stderr.on('data',data=>helperLog+=data)
 await poll(async()=>{try{return(await fetch(`http://127.0.0.1:${port}/healthz`)).ok}catch{return false}},'helper startup')
 process.env.VITE_INJOFFICE_API_BASE='.';process.env.INJOFFICE_SERVER=`http://127.0.0.1:${port}`
 server=await startShowcaseDevServer(resolve(root,'apps/playground'))
 const executable=process.env.CHROME_PATH||['/Applications/Google Chrome.app/Contents/MacOS/Google Chrome','google-chrome','chromium'].find(path=>spawnSync(path,['--version'],{stdio:'ignore'}).status===0)
 if(!executable)throw new Error('Chrome required')
 chrome=await launchChromeForCDP({executable,createProfile:()=>{const path=mkdtempSync(resolve(tmpdir(),'injoffice-pptx-chrome-'));profiles.push(path);return path}})
 cdp=await connect(chrome.target.webSocketDebuggerUrl)
 await cdp.send('Runtime.enable');await cdp.send('Page.enable');await cdp.send('Emulation.setDeviceMetricsOverride',{width:1440,height:1000,deviceScaleFactor:1,mobile:false})
 await cdp.send('Page.addScriptToEvaluateOnNewDocument',{source:`{const original=fetch;window.__pptxPosts=[];window.fetch=async(input,init)=>{if(init?.method==='POST'&&String(input).includes('/v1/')){const bytes=await init.body.arrayBuffer();const digest=await crypto.subtle.digest('SHA-256',bytes);window.__pptxPosts.push({url:String(input),hash:[...new Uint8Array(digest)].map(n=>n.toString(16).padStart(2,'0')).join('')})}return original(input,init)}}`})
 await cdp.send('Page.navigate',{url:`${server.url}#/slides?feature=pptx-native`})
 await poll(()=>evaluate(`!!document.querySelector('input[type=file]')`),'PPTX file input')
 await upload(fixture)
 await poll(()=>evaluate(`${section}?.textContent.includes('Nothing is uploaded')`),'native consent')
 await assert('window.__pptxPosts.length===0','opening source did not upload')
 await clickRender()
 await poll(()=>evaluate(`${section}?.querySelectorAll('svg path').length>20`),'real native glyph outline replay',45000)
 await assert(`window.__pptxPosts.length===1&&window.__pptxPosts[0].hash===${JSON.stringify(sourceHash)}`,'exact source submitted once')
 await assert(`${section}.textContent.includes('not Office pixel-equivalence')&&${section}.textContent.includes('text.deterministicLayout')`,'measured policy is honestly labeled')
 const bounds=await evaluate(`(()=>{const svg=${section}.querySelector('svg');const paths=[...svg.querySelectorAll('path')];const groups=[0,1,2].map(i=>paths.filter(path=>{const r=path.getBoundingClientRect(),s=svg.getBoundingClientRect();const x=(r.x-s.x)/s.width;return x>i*.3&&x<(i+1)*.34}));return groups.map(paths=>Math.min(...paths.map(p=>p.getBoundingClientRect().y)))})()`)
 if(!bounds.every(Number.isFinite)||!(bounds[0]<bounds[1]&&bounds[1]<bounds[2]))throw new Error(`Anchor glyph positions are not ordered: ${bounds}`)
 await evaluate(`${section}.scrollIntoView()`);const shot=await cdp.send('Page.captureScreenshot',{format:'png'});writeFileSync(resolve(artifacts,'pptx-measured-native.png'),Buffer.from(shot.data,'base64'))
 if(hash(readFileSync(fixture))!==sourceHash)throw new Error('Original source changed')
 const bulletFixture=resolve(scratch,'mixed-anchors-bullets.pptx'),bulletHash=hash(readFileSync(bulletFixture));await upload(bulletFixture)
 await poll(()=>evaluate(`${section}?.textContent.includes('Nothing is uploaded')&&${section}?.querySelector('svg')===null`),'bullet source replacement clears output');await clickRender()
 await poll(()=>evaluate(`${section}?.querySelectorAll('[data-native-source-role="paragraphBullet"]').length===3`),'exact glyph markers on wrapped source paragraphs',45000)
 await assert(`(()=>{const markers=[...${section}.querySelectorAll('[data-native-source-role="paragraphBullet"]')];return markers.every(marker=>{const content=marker.parentElement.querySelector('[data-native-source-role="contentRun"]');const m=marker.transform.baseVal.consolidate().matrix,c=content.transform.baseVal.consolidate().matrix;return m.e<c.e&&m.f===c.f})})()`,'markers precede content at the same measured baseline')
 await assert(`window.__pptxPosts.length===2&&window.__pptxPosts[1].hash===${JSON.stringify(bulletHash)}`,'bullet source bytes remain unchanged')
 const cropFixture=resolve(scratch,'mixed-anchors-crop.pptx'),cropHash=hash(readFileSync(cropFixture));await upload(cropFixture)
 await poll(()=>evaluate(`${section}?.textContent.includes('Nothing is uploaded')&&${section}?.querySelector('svg')===null`),'crop source replacement');await clickRender()
 await poll(()=>evaluate(`!!${section}?.querySelector('[data-native-raster]')`),'source-bound native cropped raster',45000)
 await evaluate(`${section}.scrollIntoView()`)
 const cropBounds=await evaluate(`(()=>{const svg=${section}.querySelector('[data-native-raster]').parentElement,v=svg.viewBox.baseVal,m=svg.getScreenCTM(),p=new DOMPoint(v.x,v.y).matrixTransform(m),q=new DOMPoint(v.x+v.width,v.y+v.height).matrixTransform(m);return {x:p.x,y:p.y,width:q.x-p.x,height:q.y-p.y}})()`)
 const cropShot=await cdp.send('Page.captureScreenshot',{format:'png'}),pixels=PNG.sync.read(Buffer.from(cropShot.data,'base64'));writeFileSync(resolve(artifacts,'pptx-native-crop.png'),Buffer.from(cropShot.data,'base64'))
 for(const fx of [.2,.5,.8])for(const fy of [.2,.5,.8]){const x=Math.round(cropBounds.x+cropBounds.width*fx),y=Math.round(cropBounds.y+cropBounds.height*fy);if(x<0||y<0||x>=pixels.width||y>=pixels.height)throw new Error('Crop sample outside screenshot');const offset=(y*pixels.width+x)*4;if(pixels.data[offset]>5||pixels.data[offset+1]<250||pixels.data[offset+2]>5)throw new Error(`Source crop did not paint only green quadrant at ${x},${y}: ${pixels.data.subarray(offset,offset+4)}`)}
 await assert(`window.__pptxPosts.length===3&&window.__pptxPosts[2].hash===${JSON.stringify(cropHash)}`,'crop source identity preserved')
 await evaluate(`${section}.querySelector('[data-native-raster]').dispatchEvent(new Event('error'))`)
 await poll(()=>evaluate(`${section}?.querySelector('svg')===null&&${section}?.textContent.includes('Native image decoding failed')`),'raster decode error clears native success')
 const inheritedFixture=resolve(scratch,'mixed-anchors-inherited.pptx'),inheritedHash=hash(readFileSync(inheritedFixture));await upload(inheritedFixture)
 await poll(()=>evaluate(`${section}?.textContent.includes('Nothing is uploaded')&&${section}?.querySelector('svg')===null`),'inherited source replacement');await clickRender()
 await poll(()=>evaluate(`${section}?.querySelectorAll('[data-native-source-role="paragraphBullet"]').length===1&&${section}?.querySelectorAll('[data-native-source-role="contentRun"]').length>=10`),'master prompt style inherited into source-only slide text',45000)
 await assert(`${section}.querySelectorAll('[data-native-source-role="contentRun"]').length===11`,'ancestor prompt text is never painted')
 await assert(`window.__pptxPosts.length===4&&window.__pptxPosts[3].hash===${JSON.stringify(inheritedHash)}`,'inherited source identity preserved')
 const arrowFixture=resolve(scratch,'mixed-anchors-arrows.pptx'),arrowHash=hash(readFileSync(arrowFixture));await upload(arrowFixture)
 await poll(()=>evaluate(`${section}?.textContent.includes('Nothing is uploaded')&&${section}?.querySelector('svg')===null`),'arrow source replacement');await clickRender()
 await poll(()=>evaluate(`${section}?.querySelectorAll('[data-native-source-role="connectorArrow"]').length===5`),'all five source arrow shapes replayed',45000)
 await assert(`(()=>{const a=[...${section}.querySelectorAll('[data-native-source-role="connectorArrow"]')];return a[3].querySelector('ellipse')!==null&&a[4].querySelector('path').getAttribute('fill')==='none'&&new Set(a.filter((_,i)=>i!==3).map(n=>n.querySelector('path').getAttribute('d'))).size===4})()`,'typed endpoint shapes are distinct, not guessed triangles')
 await assert(`${section}.textContent.includes('arrow.deterministicGeometry')&&window.__pptxPosts.length===5&&window.__pptxPosts[4].hash===${JSON.stringify(arrowHash)}`,'source arrow policy and identity visible')
 await evaluate(`${section}.scrollIntoView()`);const arrowShot=await cdp.send('Page.captureScreenshot',{format:'png'});writeFileSync(resolve(artifacts,'pptx-native-arrows.png'),Buffer.from(arrowShot.data,'base64'))
 const chartFixture=resolve(scratch,'mixed-anchors-chart.pptx'),chartHash=hash(readFileSync(chartFixture));await upload(chartFixture)
 await poll(()=>evaluate(`${section}?.textContent.includes('Nothing is uploaded')&&${section}?.querySelector('svg')===null`),'chart source replacement');await clickRender()
 await poll(()=>evaluate(`${section}?.querySelectorAll('[data-native-raster]').length===1`),'grouped chart-only cached preview reaches real worker and browser',45000)
 await assert(`window.__pptxPosts.length===6&&window.__pptxPosts[5].hash===${JSON.stringify(chartHash)}`,'chart preview uses exact original source bytes')
 if(hash(readFileSync(chartFixture))!==chartHash)throw new Error('Original chart source changed')
 const missing=resolve(root,'go/pptxpatch/testdata/playground_northstar_review.pptx');await upload(missing)
 await poll(()=>evaluate(`${section}?.textContent.includes('Nothing is uploaded')&&${section}?.querySelector('svg')===null`),'replacement clears stale native output')
 await assert('window.__pptxPosts.length===6','replacement still requires consent');await clickRender()
 await poll(()=>evaluate(`${section}?.textContent.includes('Exact operator font unavailable')`),'missing exact font refusal',45000)
 await assert(`${section}.querySelector('svg')===null`,'missing font never substitutes browser glyphs')
 if(errors.length)throw new Error(errors.join('\n'))
 console.log(JSON.stringify({result:'PASS',checks:['explicit upload consent and source SHA','measured HarfBuzz glyphs and vertical anchors','wrapped hanging bullets','source crop pixels and image-decode refusal','master defaults without prompt-text leakage','five typed arrow shapes and visible geometry policy','grouped chart-only cached raster preview','missing exact font refusal'],screenshots:artifacts},null,2))
}finally{cdp?.close();if(chrome)await terminateProcess(chrome.child);await server?.close();if(helper)await terminateProcess(helper);for(const profile of profiles)rmSync(profile,{recursive:true,force:true,maxRetries:10,retryDelay:100});rmSync(scratch,{recursive:true,force:true})}
async function unusedPort(){const server=createServer();await new Promise(done=>server.listen(0,'127.0.0.1',done));const port=server.address().port;await new Promise(done=>server.close(done));return port}
async function evaluate(expression){const r=await cdp.send('Runtime.evaluate',{expression,awaitPromise:true,returnByValue:true});if(r.exceptionDetails)throw new Error(r.exceptionDetails.exception?.description||r.exceptionDetails.text);return r.result.value}
async function assert(expression,label){if(!await evaluate(expression))throw new Error(label)}
async function poll(check,label,timeout=30000){const end=Date.now()+timeout;while(Date.now()<end){if(await check())return;await new Promise(done=>setTimeout(done,75))}throw new Error(`Timed out: ${label}`)}
async function upload(path){const handle=await cdp.send('Runtime.evaluate',{expression:`document.querySelector('input[type=file]')`});const node=await cdp.send('DOM.describeNode',{objectId:handle.result.objectId});await cdp.send('DOM.setFileInputFiles',{backendNodeId:node.node.backendNodeId,files:[path]})}
async function clickRender(){await evaluate(`[...${section}.querySelectorAll('button')].find(b=>b.textContent==='Upload to helper and render native slide').click()`)}
async function connect(url){const socket=new WebSocket(url);await new Promise((done,reject)=>{socket.addEventListener('open',done,{once:true});socket.addEventListener('error',reject,{once:true})});let next=0;const waiting=new Map();socket.addEventListener('message',event=>{const message=JSON.parse(event.data);if(message.id){const p=waiting.get(message.id);waiting.delete(message.id);message.error?p?.reject(new Error(message.error.message)):p?.resolve(message.result)}else if(message.method==='Runtime.exceptionThrown')errors.push(message.params.exceptionDetails.text)});return{send(method,params={}){return new Promise((resolve,reject)=>{const id=++next;waiting.set(id,{resolve,reject});socket.send(JSON.stringify({id,method,params}))})},close(){socket.close()}}}
