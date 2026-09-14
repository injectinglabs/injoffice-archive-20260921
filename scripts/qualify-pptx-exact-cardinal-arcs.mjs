import assert from 'node:assert/strict';
import {readFileSync, writeFileSync, mkdtempSync, mkdirSync} from 'node:fs';
import {tmpdir} from 'node:os';
import {resolve, dirname} from 'node:path';
import {fileURLToPath} from 'node:url';
import {execFileSync} from 'node:child_process';
import {createHash} from 'node:crypto';
const root=resolve(dirname(fileURLToPath(import.meta.url)),'..');
const out=process.env.PPTX_EXACT_CARDINAL_OUTPUT || mkdtempSync(resolve(tmpdir(),'injoffice-exact-cardinal-'));
mkdirSync(out,{recursive:true});
execFileSync('go',['test','.','-run','^TestNativeGeometryExactCardinalSourceOracle$','-count=1'],{cwd:resolve(root,'go/pptxpatch'),env:{...process.env,PPTX_EXACT_CARDINAL_FIXTURES:out},stdio:'inherit'});
const oracle=JSON.parse(readFileSync(resolve(root,'go/pptxpatch/testdata/exact-cardinal-arcs-oracle.json')));
const hash=b=>createHash('sha256').update(b).digest('hex');
const font=resolve(root,'node_modules/dejavu-fonts-ttf/ttf/DejaVuSans.ttf');
const fontManifest=resolve(out,'fonts.json');
writeFileSync(fontManifest,JSON.stringify({version:1,faces:[{family:'DejaVu Sans',weight:400,style:'normal',path:font,sha256:'sha256:'+hash(readFileSync(font))}]}));
await import(resolve(root,'packages/pptx-wasm/dist/wasm_exec.js'));
const wasm=readFileSync(resolve(root,'packages/pptx-wasm/dist/pptxnative.wasm'));
const go=new Go();go.run((await WebAssembly.instantiate(wasm,go.importObject)).instance);
const {compilePptxPreview}=await import(resolve(root,'apps/pptx-page-paint-worker/dist/compile.js'));
const records=[];
for(const name of [...oracle.cases.map(c=>c.name),'near-cardinal','non-unit','prior-error']) {
 const path=resolve(out,name+'.pptx'),bytes=readFileSync(path),before=hash(bytes);
 const extracted=pptxnative.extract(bytes);assert.equal(extracted.ok,true,extracted.error);
 const deck=JSON.parse(extracted.value),snapshot=JSON.stringify(deck);
 const shapes=deck.slides.flatMap(s=>s.elements).filter(e=>e.kind==='shape');assert.equal(shapes.length,1);
 const shape=shapes[0],expected=oracle.cases.find(c=>c.name===name);
 if(expected){assert.equal(shape.compatibility.status,'preserveOnly');assert.deepEqual(shape.geometry.paths.map(p=>p.commands),expected.paths);assert.deepEqual(shape.geometry.textRect,expected.textRect)}
 else {assert.equal(shape.compatibility.status,'refused');assert.equal(shape.geometry,undefined);assert.equal(shape.preset,undefined)}
 const request={deck,slide_index:0,package_sha256:before,font_manifest_path:fontManifest};
 const result=await compilePptxPreview(request);assert.deepEqual(await compilePptxPreview(request),result);
 assert.equal(JSON.stringify(deck),snapshot);assert.equal(hash(bytes),before);assert.equal(hash(readFileSync(path)),before);
 const leaves=[];const walk=ns=>{for(const n of ns){if(n.kind==='group')walk(n.children);else leaves.push(n)}};walk(result.nodes);
 const paths=leaves.filter(n=>n.kind==='path'),placeholders=leaves.filter(n=>n.kind==='placeholder');
 let maximum=0;
 for(const n of paths){assert.ok(n.d.length>0&&n.d.length<8*1024*1024);for(const m of n.d.matchAll(/[-+]?(?:\d*\.\d+|\d+\.?)(?:[eE][-+]?\d+)?/g)){const v=Number(m[0]);assert.ok(Number.isFinite(v)&&Math.abs(v)<=1e9);maximum=Math.max(maximum,Math.abs(v))}}
 if(expected){assert.ok(paths.length>=expected.paths.length);assert.equal(placeholders.length,0);assert.ok(!result.diagnostics.some(d=>/refused|unavailable/i.test(d)))}
 else {assert.equal(paths.length,0);assert.equal(placeholders.length,1);assert.ok(result.diagnostics.some(d=>d.includes('geometry accumulated path uncertainty exceeds one eighth EMU')))}
 records.push({name,sourceSHA256:before,status:expected?'painted':'refused',paths:paths.length,placeholders:placeholders.length,maxCoordinate:maximum,diagnostics:result.diagnostics});
}
writeFileSync(resolve(out,'report.json'),JSON.stringify({wasmSHA256:hash(wasm),records,limitations:'Exact numerical/source/WASM/full public compile and paint qualification; no browser pixels or Office parity.'},null,2)+'\n');
console.log(JSON.stringify({out,painted:records.filter(r=>r.status==='painted').length,refused:records.filter(r=>r.status==='refused').length}));
process.exit(0);
