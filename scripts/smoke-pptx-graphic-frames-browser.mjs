import {spawnSync} from 'node:child_process';
import {existsSync,mkdirSync} from 'node:fs';
import {tmpdir} from 'node:os';
import {isDeepStrictEqual} from 'node:util';
import {readFileSync,writeFileSync,mkdtempSync} from 'node:fs';
import {resolve} from 'node:path';
import {createHash} from 'node:crypto';



const root=resolve(import.meta.dirname,'..');
const out=process.env.SHOWCASE_OUTPUT?resolve(process.env.SHOWCASE_OUTPUT):mkdtempSync(resolve(tmpdir(),'injoffice-frame-browser-'));
mkdirSync(out,{recursive:true});
const fixtures=resolve(out,'fixtures');
const generated=spawnSync('go',['test','-run','^TestNativeGraphicFrameBrowserFixtures$','-count=1','.'],{cwd:resolve(root,'go/pptxpatch'),env:{...process.env,INJOFFICE_PPTX_GRAPHIC_FRAME_FIXTURE_DIR:fixtures},encoding:'utf8',timeout:90000});
if(generated.status!==0)throw Error('Source fixture generation failed: '+generated.stdout+'\n'+generated.stderr);
const {build}=await import(resolve(root,'node_modules/vite/dist/node/index.js'));
const {launchChromeForCDP,terminateProcess}=await import(resolve(root,'scripts/chrome-cdp-startup.mjs'));
const {compilePptxPreview}=await import(resolve(root,'apps/pptx-page-paint-worker/dist/compile.js'));
const names=['table-intrinsic','table-rotated-group','table-y-overflow','table-above-slide'];
const hash=value=>createHash('sha256').update(value).digest('hex');
const flatten=es=>es.flatMap(e=>[e,...(e.children?flatten(e.children):[])]);
function sourceTable(deck,name){
 const tables=flatten(deck.slides[0]?.elements??[]).filter(e=>e.kind==='table');
 if(tables.length!==1)throw Error(name+': expected exactly one source table, got '+tables.length);
 const table=tables[0];
 if(table.graphicFrameLayout!=='source-anchored-v1'||table.provenance!=='parsed'||!table.source||table.compatibility.status!=='preserveOnly')throw Error(name+': qualified source authority missing');
 return table;
}
const font=resolve(root,'node_modules/dejavu-fonts-ttf/ttf/DejaVuSans.ttf'),fontHash=`sha256:${createHash('sha256').update(readFileSync(font)).digest('hex')}`,manifest=resolve(out,'fonts.json');
writeFileSync(manifest,JSON.stringify({version:1,faces:[{family:'DejaVu Sans',weight:400,style:'normal',path:font,sha256:fontHash}]}));
// Preflight runs before Chrome. Refused/empty extraction cannot count as visual proof.
await import(resolve(root,'packages/pptx-wasm/dist/wasm_exec.js'));
const candidateWasm=readFileSync(resolve(root,'packages/pptx-wasm/dist/pptxnative.wasm'));
const nodeGo=new Go(),nodeInstance=await WebAssembly.instantiate(candidateWasm,nodeGo.importObject);nodeGo.run(nodeInstance.instance);
const preflight=[];
for(const name of names){
 const bytes=readFileSync(resolve(fixtures,name+'.pptx')),result=pptxnative.extract(bytes);
 if(!result.ok)throw Error(name+': '+result.error);
 const deck=JSON.parse(result.value),table=sourceTable(deck,name),retained=sourceTable(JSON.parse(readFileSync(resolve(fixtures,name+'-go.json'))),name+' retained Go');
 if(!isDeepStrictEqual(table.table,retained.table)||!isDeepStrictEqual(table.transform,retained.transform))throw Error(name+': Go/WASM source metadata differs');
 const snapshot=JSON.stringify(deck),preview=await compilePptxPreview({deck,slide_index:0,package_sha256:hash(bytes),font_manifest_path:manifest});
 if(JSON.stringify(deck)!==snapshot)throw Error(name+': source mutated');
 const glyphs=flatten(preview.nodes).filter(n=>n.kind==='path'&&n.fill==='112233').length;
 if(glyphs===0||preview.diagnostics.some(d=>d.startsWith('text.refused')))throw Error(name+': expected actual supplied glyphs, not refusal placeholder');
 if(!preview.diagnostics.some(d=>d.startsWith('graphicFrame.sourceAnchoredPreview')))throw Error(name+': profile not compiled');
 preflight.push({name,sourceSHA256:hash(bytes),glyphs});
 writeFileSync(resolve(out,name+'-node-wasm.json'),snapshot);
 writeFileSync(resolve(out,name+'-node-preview.json'),JSON.stringify(preview,null,2));
}
const negative=pptxnative.extract(readFileSync(resolve(fixtures,'table-unqualified-x-clip.pptx')));
if(!negative.ok)throw Error('Negative fixture extraction failed before refusal inventory');
if(flatten(JSON.parse(negative.value).slides[0].elements).some(e=>e.kind==='table'))throw Error('Unsupported x-clip fixture unexpectedly admitted');
writeFileSync(resolve(out,'preflight.json'),JSON.stringify({root,wasmSHA256:hash(candidateWasm),fontHash,preflight,negative:'unsupported source table x-clip remains unadmitted',browser:'not run'},null,2));
console.log('Preflight PASS: four source-profiled tables, actual glyphs, immutable source, retained negative refusal. Evidence: '+out);
if(process.argv.includes('--preflight-only'))process.exit(0);
const entry=resolve(root,'virtual-affine-browser.js');
const built=await build({configFile:false,root,logLevel:'warn',plugins:[{name:'affine-proof',resolveId:id=>id===entry?entry:undefined,load:id=>id===entry?"import {createElement} from 'react';import {createRoot} from 'react-dom/client';import {flushSync} from 'react-dom';import {NativePptxVector} from './apps/playground/src/components/NativePptxSlides.tsx';import {PptxFilePreviewVector} from './apps/playground/src/components/PptxFilePreview.tsx';import {compileFilePreviewGeometry} from './apps/playground/src/filePreviewGeometry.ts';const root=createRoot(document.body);globalThis.renderLegacy=async deck=>{const geometry=await compileFilePreviewGeometry(deck,0);flushSync(()=>root.render(createElement(PptxFilePreviewVector,{deck,geometry})));};globalThis.renderPreview=preview=>flushSync(()=>root.render(createElement(NativePptxVector,{preview})));":undefined}],define:{'process.env.NODE_ENV':'"production"'},build:{write:false,minify:false,lib:{entry,name:'AffineProof',formats:['iife']}}});
const code=(Array.isArray(built)?built:[built]).flatMap(x=>x.output).find(x=>x.type==='chunk'&&x.isEntry).code;
let chrome,cdp;
try{
 chrome=await launchChromeForCDP({executable:process.env.CHROME_BIN??['/Applications/Google Chrome.app/Contents/MacOS/Google Chrome','/usr/bin/google-chrome','/usr/bin/chromium','/usr/bin/chromium-browser'].find(existsSync)??'google-chrome',createProfile:()=>mkdtempSync(resolve(out,'chrome-'))});
 cdp=await connect(chrome.target.webSocketDebuggerUrl);await cdp.send('Page.enable');await cdp.send('Emulation.setDeviceMetricsOverride',{width:1000,height:700,deviceScaleFactor:1,mobile:false});
 await evaluate(readFileSync(resolve(root,'packages/pptx-wasm/dist/wasm_exec.js'),'utf8'));
 const wasm=readFileSync(resolve(root,'packages/pptx-wasm/dist/pptxnative.wasm'));
 await evaluate(`(async()=>{const go=new Go();const bytes=Uint8Array.from(atob(${JSON.stringify(wasm.toString('base64'))}),c=>c.charCodeAt(0));const {instance}=await WebAssembly.instantiate(bytes,go.importObject);go.run(instance);return true})()`);
 const extract=async name=>{const bytes=readFileSync(resolve(fixtures,name+'.pptx'));const deck=await evaluate(`(()=>{const r=pptxnative.extract(Uint8Array.from(atob(${JSON.stringify(bytes.toString('base64'))}),c=>c.charCodeAt(0)));if(!r.ok)throw Error(r.error);return JSON.parse(r.value)})()`);writeFileSync(resolve(out,name+'-wasm.json'),JSON.stringify(deck,null,2));return {deck,sha256:createHash('sha256').update(bytes).digest('hex')}};
 const results=[];
 for(const name of names){
  const batch=await extract(name),table=sourceTable(batch.deck,name+' browser');
  if(table?.graphicFrameLayout!=='source-anchored-v1')throw Error('Source table profile missing');
  const go=JSON.parse(readFileSync(resolve(fixtures,name+'-go.json'))),goTable=flatten(go.slides[0].elements).find(e=>e.kind==='table');
  if(!isDeepStrictEqual(table.table,goTable.table)||!isDeepStrictEqual(table.transform,goTable.transform))throw Error('Go/WASM metadata differs');
  const snapshot=JSON.stringify(batch.deck),preview=await compilePptxPreview({deck:batch.deck,slide_index:0,package_sha256:batch.sha256,font_manifest_path:manifest});
  if(JSON.stringify(batch.deck)!==snapshot)throw Error('Source mutated');
  writeFileSync(resolve(out,name+'-preview.json'),JSON.stringify(preview,null,2));
  await evaluate("document.body.style.margin='12px';document.body.style.background='white';document.body.style.width='960px'");
  if(!results.length)await evaluate(code);
  for(const surface of ['native','legacy']){
   await evaluate(surface==='native'?`renderPreview(${JSON.stringify(preview)})`:`renderLegacy(${JSON.stringify(batch.deck)})`);
   await new Promise(r=>setTimeout(r,100));
   const box=await evaluate(`(()=>{const r=document.querySelector('svg [fill="#EEEEEE"]');if(!r)throw Error('Missing intrinsic background');return r.getBoundingClientRect().toJSON()})()`);
   const normalized=await evaluate(`(()=>{const svg=document.querySelector('svg'),r=svg.querySelector('[fill="#EEEEEE"]');return {widthEmu:Number(r.getAttribute('width'))*12700,heightEmu:Number(r.getAttribute('height'))*12700,viewBox:svg.getAttribute('viewBox')}})()`);
   if(Math.abs(normalized.widthEmu-table.table.columnWidths[0])>0.000001||Math.abs(normalized.heightEmu-table.table.rowHeights[0])>0.000001)throw Error('Intrinsic tracks changed '+name+' '+surface+' '+JSON.stringify(normalized));
   const shot=await cdp.send('Page.captureScreenshot',{format:'png'});writeFileSync(resolve(out,name+'-'+surface+'.png'),Buffer.from(shot.data,'base64'));
   const pixels=await evaluate(`(async()=>{const image=new Image();image.src='data:image/png;base64,${shot.data}';await image.decode();const c=document.createElement('canvas');c.width=image.width;c.height=image.height;const x=c.getContext('2d');x.drawImage(image,0,0);const d=x.getImageData(0,0,c.width,c.height).data,b=${JSON.stringify(box)};let ink=0,right=0,below=0;for(let y=0;y<c.height;y++)for(let xx=0;xx<c.width;xx++){const i=(y*c.width+xx)*4;if(d[i]<100&&d[i+1]>d[i]+4&&d[i+2]>d[i+1]+4&&d[i+2]<170){ink++;if(xx>b.right+1)right++;if(y>b.bottom+1)below++;}}return {ink,right,below}})()`);
   if(name==='table-above-slide'?pixels.ink!==0:pixels.ink<2)throw Error('Positive/slide clip pixel failure '+name+' '+surface+' '+JSON.stringify(pixels));
   if(name==='table-y-overflow'&&pixels.below<2)throw Error('Vertical overflow lost '+surface+' '+JSON.stringify(pixels));
   results.push({name,surface,sourceSHA256:batch.sha256,tableTransform:table.transform,box,normalized,pixels});
  }
 }
 writeFileSync(resolve(out,'proof.json'),JSON.stringify({wasmSHA256:createHash('sha256').update(wasm).digest('hex'),wasmBytes:wasm.length,fontHash,results,qualification:'Actual source Go/WASM + public compiler/font worker + existing browser vector component. Source-anchored intrinsic table tracks and unscaled glyphs; active slide clips and retained vertical overflow; source cell-x clipping remains unqualified on native and explicitly approximate legacy surfaces. No Office parity claim.'},null,2));console.log('Graphic-frame actual WASM + worker + native/legacy browser PASS');
}catch(error){
 writeFileSync(resolve(out,'failure.txt'),error instanceof Error?error.stack??error.message:String(error));
 if(cdp){try{const shot=await cdp.send('Page.captureScreenshot',{format:'png'});writeFileSync(resolve(out,'failure.png'),Buffer.from(shot.data,'base64'));}catch{}}
 throw error;
}finally{cdp?.close();if(chrome)await terminateProcess(chrome.child)}
async function evaluate(expression){const result=await cdp.send('Runtime.evaluate',{expression,awaitPromise:true,returnByValue:true});if(result.exceptionDetails)throw Error(result.exceptionDetails.exception?.description??result.exceptionDetails.text);return result.result.value}
async function connect(url){const socket=new WebSocket(url);await new Promise((res,rej)=>{socket.addEventListener('open',res,{once:true});socket.addEventListener('error',rej,{once:true})});let next=0;const pending=new Map();socket.addEventListener('message',event=>{const message=JSON.parse(event.data),request=pending.get(message.id);if(!request)return;pending.delete(message.id);clearTimeout(request.timer);if(message.error)request.reject(Error(message.error.message));else request.resolve(message.result)});return{send(method,params={}){return new Promise((resolve,reject)=>{const id=++next,timer=setTimeout(()=>{pending.delete(id);reject(Error('CDP timeout '+method))},30000);pending.set(id,{resolve,reject,timer});socket.send(JSON.stringify({id,method,params}))})},close(){socket.close()}}}
