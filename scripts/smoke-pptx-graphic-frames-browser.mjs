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
const generated=spawnSync('go',['test','-run','^TestNativeGraphicFrame.*BrowserFixtures$','-count=1','.'],{cwd:resolve(root,'go/pptxpatch'),env:{...process.env,INJOFFICE_PPTX_GRAPHIC_FRAME_FIXTURE_DIR:fixtures},encoding:'utf8',timeout:90000});
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

const chartNames=['chart-literal-baseline','chart-literal-group','chart-workbook-baseline','chart-workbook-group'];
const xlsxWasm=readFileSync(resolve(root,'packages/xlsx-wasm/dist/xlsxnative.wasm'));
const xlsxGo=new Go();xlsxGo.run((await WebAssembly.instantiate(xlsxWasm,xlsxGo.importObject)).instance);
function sourceChart(deck,name){
 const charts=flatten(deck.slides[0]?.elements??[]).filter(e=>e.kind==='chart');
 if(charts.length!==1)throw Error(name+': expected exactly one chart');
 const c=charts[0];
 if(c.graphicFrameLayout!=='source-anchored-v1'||c.provenance!=='parsed'||!c.source||c.compatibility.status!=='preserveOnly'||c.transform.cx!==6000001||c.transform.cy!==4000001||c.transform.rotationAngle!==1800000)throw Error(name+': source chart/profile/raw frame drift');
 return c;
}
// Opaque capability tokens are intentionally engine-instance-specific. Compare
// all source-bearing fields while requiring each engine to retain its own token.
function chartMetadata(chart){
 if(typeof chart.opaqueRef?.token!=='string'||!chart.opaqueRef.token)throw Error('Missing chart capability token');
 return {...chart,opaqueRef:{...chart.opaqueRef,token:'engine-specific'}};
}
function chartRequest(deck,bytes,name){
 const request={deck,slide_index:0,package_sha256:hash(bytes),font_manifest_path:manifest};
 if(!name.includes('workbook'))return {...request,source_chart_preview:true};
 const result=pptxnative.inspectChartWorkbooks(bytes);if(!result.ok)throw Error(result.error);
 const inspection=JSON.parse(result.value);
 if(inspection.charts.length!==1||inspection.workbooks.length!==1||inspection.omissions.length)throw Error(name+': workbook source closure missing');
 const workbooks=inspection.workbooks.map(w=>{
  const embedded=Buffer.from(w.bytesBase64,'base64');if(hash(embedded)!==w.sha256)throw Error('embedded digest drift');
  const result=xlsxnative.extract(embedded);if(!result.ok)throw Error(result.error);
  return {part:w.part,sha256:w.sha256,contract_json:result.value};
 });
 return {...request,workbook_chart_preview:true,workbook_chart_data:{inspection_json:result.value,workbooks}};
}
function glyphAxes(nodes,parent=[1,0,0,1,0,0]){
 const multiply=(a,b)=>[a[0]*b[0]+a[2]*b[1],a[1]*b[0]+a[3]*b[1],a[0]*b[2]+a[2]*b[3],a[1]*b[2]+a[3]*b[3],a[0]*b[4]+a[2]*b[5]+a[4],a[1]*b[4]+a[3]*b[5]+a[5]];
 return nodes.flatMap(n=>n.kind==='group'?glyphAxes(n.children,multiply(parent,n.transform)):n.kind==='path'&&n.fill==='123456'?[parent.slice(0,4)]:[]);
}
const chartPreviews=new Map();
for(const name of chartNames){
 const bytes=readFileSync(resolve(fixtures,name+'.pptx')),before=hash(bytes),r=pptxnative.extract(bytes);if(!r.ok)throw Error(r.error);
 const deck=JSON.parse(r.value),chart=sourceChart(deck,name),retained=sourceChart(JSON.parse(readFileSync(resolve(fixtures,name+'-go.json'))),name+' Go');
 if(!isDeepStrictEqual(chartMetadata(chart.chart),chartMetadata(retained.chart))||!isDeepStrictEqual(chart.transform,retained.transform)){writeFileSync(resolve(out,name+'-metadata-difference.json'),JSON.stringify({actual:chart,expected:retained},null,2));throw Error(name+': chart source metadata differs '+out);}
 const snapshot=JSON.stringify(deck),request=chartRequest(deck,bytes,name),requestBefore=JSON.stringify(request),preview=await compilePptxPreview(request);
 if(JSON.stringify(deck)!==snapshot||JSON.stringify(request)!==requestBefore||hash(bytes)!==before)throw Error(name+': chart source/engine request mutated');
 const nodes=flatten(preview.nodes),series=nodes.filter(n=>n.kind==='path'&&n.fill==='E53935'),axes=glyphAxes(preview.nodes);
 if(series.length!==2||axes.length<4||nodes.some(n=>n.kind==='placeholder')||preview.diagnostics.some(d=>/unavailable|refused/i.test(d)))throw Error(name+': expected real series and supplied-label paths: '+JSON.stringify(preview.diagnostics));
 if(!preview.diagnostics.some(d=>d.startsWith('graphicFrame.sourceAnchoredPreview')))throw Error(name+': source frame profile not compiled');
 const expectedWidth=name.endsWith('-group')?7500001:6000001;
 const frameNode=nodes.find(n=>n.kind==='group'&&n.clip?.cx===expectedWidth&&n.clip?.cy===4000001);
 if(!frameNode||!isDeepStrictEqual(frameNode.transform.slice(0,4),[1,0,0,1]))throw Error(name+': projected physical frame/axis orientation missing');
 const plot=frameNode.children.find(n=>n.kind==='group'&&n.clip)?.clip;
 if(!plot)throw Error(name+': actual plot clip missing');
 const expectedValues=name.includes('workbook')?[4,4]:[0.5,12.5];
 for(let i=0;i<series.length;i++){
  const coordinates=series[i].d.match(/-?\d+(?:\.\d+)?/g).map(Number),ys=coordinates.filter((_,at)=>at%2===1);
  if(Math.abs(Math.max(...ys)-Math.min(...ys)-plot.cy*expectedValues[i]/30)>1)throw Error(name+': source values do not control actual bar heights (stale cache must not win)');
 }

 if(name.endsWith('-group')){
  if(!preview.diagnostics.some(d=>d.includes('nearest-EMU physical dimensions')))throw Error(name+': physical projection evidence absent');
  const baseline=chartPreviews.get(name.replace('-group','-baseline'));
  if(axes.some(a=>!isDeepStrictEqual(a,glyphAxes(baseline.nodes)[0])))throw Error(name+': source ancestor stretched/reflected glyph axes');
 }
 chartPreviews.set(name,preview);
 preflight.push({name,sourceSHA256:before,seriesPaths:series.length,axisGlyphPaths:axes.length,glyphAxes:axes[0],projectedWidth:expectedWidth,sourceFrame:chart.transform});
 writeFileSync(resolve(out,name+'-node-preview.json'),JSON.stringify(preview,null,2));
}
const negative=pptxnative.extract(readFileSync(resolve(fixtures,'table-unqualified-x-clip.pptx')));
if(!negative.ok)throw Error('Negative fixture extraction failed before refusal inventory');
if(flatten(JSON.parse(negative.value).slides[0].elements).some(e=>e.kind==='table'))throw Error('Unsupported x-clip fixture unexpectedly admitted');
writeFileSync(resolve(out,'preflight.json'),JSON.stringify({root,wasmSHA256:hash(candidateWasm),xlsxWasmSHA256:hash(xlsxWasm),fontHash,preflight,negative:'unsupported source table x-clip remains unadmitted',browser:'not run'},null,2));
console.log('Preflight PASS: four source tables and four literal/workbook charts, actual glyphs, immutable source, retained negative refusal. Evidence: '+out);
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

 for(const name of chartNames){
  const batch=await extract(name),bytes=readFileSync(resolve(fixtures,name+'.pptx')),chart=sourceChart(batch.deck,name+' browser');
  const retained=sourceChart(JSON.parse(readFileSync(resolve(fixtures,name+'-go.json'))),name+' Go');
  if(!isDeepStrictEqual(chartMetadata(chart.chart),chartMetadata(retained.chart))||!isDeepStrictEqual(chart.transform,retained.transform))throw Error('Browser chart source drift');
  const snapshot=JSON.stringify(batch.deck),preview=await compilePptxPreview(chartRequest(batch.deck,bytes,name));
  if(JSON.stringify(batch.deck)!==snapshot)throw Error('Browser extracted chart mutated');
  await evaluate(`renderPreview(${JSON.stringify(preview)})`);
  await evaluate(`document.querySelector('svg').scrollIntoView({block:'center',inline:'center',behavior:'instant'})`);
  const dom=await evaluate(`(()=>{const svg=document.querySelector('svg'),r=svg.getBoundingClientRect();if(r.width<2||r.height<2||r.left<0||r.top<0||r.right>innerWidth||r.bottom>innerHeight)throw Error('Chart not fully visible');return {seriesPaths:svg.querySelectorAll('path[fill="#E53935"]').length,axisPaths:svg.querySelectorAll('path[fill="#123456"]').length,clip:{x:r.left+scrollX,y:r.top+scrollY,width:r.width,height:r.height,scale:1}}})()`);
  if(dom.seriesPaths!==2||dom.axisPaths<4)throw Error('Missing actual browser chart or label paths');
  const expectedWidth=name.endsWith('-group')?7500001:6000001;
  const projected=await evaluate(`Array.from(document.querySelectorAll('svg clipPath rect')).some(r=>Math.abs(Number(r.getAttribute('width'))*12700-${expectedWidth})<0.000001&&Math.abs(Number(r.getAttribute('height'))*12700-4000001)<0.000001)`);
  if(!projected)throw Error('Browser projected frame clip differs from physical layout');

  const shot=await cdp.send('Page.captureScreenshot',{format:'png',clip:dom.clip,captureBeyondViewport:true});
  writeFileSync(resolve(out,name+'-native.png'),Buffer.from(shot.data,'base64'));
  const painted=await evaluate(`(async()=>{const image=new Image();image.src='data:image/png;base64,${shot.data}';await image.decode();const c=document.createElement('canvas');c.width=image.width;c.height=image.height;const ctx=c.getContext('2d');ctx.drawImage(image,0,0);const d=ctx.getImageData(0,0,c.width,c.height).data;let series=0,labels=0;for(let i=0;i<d.length;i+=4){if(d[i]>180&&d[i+1]<110&&d[i+2]<110)series++;if(d[i]<80&&d[i+1]>d[i]+5&&d[i+2]>d[i+1]+5&&d[i+2]<160)labels++;}return {series,labels}})()`);
  if(painted.series<10||painted.labels<10)throw Error(name+': chart screenshot lacks visible series/axis glyph ink '+JSON.stringify(painted));
  if(hash(readFileSync(resolve(fixtures,name+'.pptx')))!==batch.sha256)throw Error('Chart source bytes changed');
  results.push({name,surface:'native',sourceSHA256:batch.sha256,sourceFrame:chart.transform,projectedWidth:expectedWidth,dom,painted,glyphAxes:glyphAxes(preview.nodes)[0]});
 }
 writeFileSync(resolve(out,'proof.json'),JSON.stringify({wasmSHA256:createHash('sha256').update(wasm).digest('hex'),wasmBytes:wasm.length,xlsxWasmSHA256:hash(xlsxWasm),fontHash,results,qualification:'Actual source Go/WASM + public compiler/font worker + existing browser vector component. Source-anchored intrinsic table tracks and unscaled glyphs; literal and embedded-workbook chart source values with supplied labels, projected physical frame clips and unchanged glyph axes; active slide clips and retained vertical overflow; source cell-x clipping remains unqualified on native and explicitly approximate legacy surfaces. No Office parity claim.'},null,2));console.log('Graphic-frame actual WASM + worker + native/legacy browser PASS');
}catch(error){
 writeFileSync(resolve(out,'failure.txt'),error instanceof Error?error.stack??error.message:String(error));
 if(cdp){try{const shot=await cdp.send('Page.captureScreenshot',{format:'png'});writeFileSync(resolve(out,'failure.png'),Buffer.from(shot.data,'base64'));}catch{}}
 throw error;
}finally{cdp?.close();if(chrome)await terminateProcess(chrome.child)}
async function evaluate(expression){const result=await cdp.send('Runtime.evaluate',{expression,awaitPromise:true,returnByValue:true});if(result.exceptionDetails)throw Error(result.exceptionDetails.exception?.description??result.exceptionDetails.text);return result.result.value}
async function connect(url){const socket=new WebSocket(url);await new Promise((res,rej)=>{socket.addEventListener('open',res,{once:true});socket.addEventListener('error',rej,{once:true})});let next=0;const pending=new Map();socket.addEventListener('message',event=>{const message=JSON.parse(event.data),request=pending.get(message.id);if(!request)return;pending.delete(message.id);clearTimeout(request.timer);if(message.error)request.reject(Error(message.error.message));else request.resolve(message.result)});return{send(method,params={}){return new Promise((resolve,reject)=>{const id=++next,timer=setTimeout(()=>{pending.delete(id);reject(Error('CDP timeout '+method))},30000);pending.set(id,{resolve,reject,timer});socket.send(JSON.stringify({id,method,params}))})},close(){socket.close()}}}
