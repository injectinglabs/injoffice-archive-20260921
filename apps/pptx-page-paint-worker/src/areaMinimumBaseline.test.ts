import {afterAll,expect,it} from 'vitest'
import {readFileSync,writeFileSync,mkdtempSync,rmSync} from 'node:fs'
import {execFileSync} from 'node:child_process'
import {resolve} from 'node:path'
import {tmpdir} from 'node:os'
import {createHash} from 'node:crypto'
import {compilePptxPreview} from './compile.js'
import {decodePptxPreview} from './contract.js'
const root=resolve(import.meta.dirname,'../../..'),scratch=mkdtempSync(resolve(tmpdir(),'pptx-area-min-'))
afterAll(()=>rmSync(scratch,{recursive:true,force:true}))
const hash=(b:Uint8Array)=>createHash('sha256').update(b).digest('hex')
const flatten=(nodes:any[]):any[]=>nodes.flatMap(n=>n.sourceRole==='contentRun'||n.sourceRole==='paragraphBullet'?[]:[n,...(n.children?flatten(n.children):[])])
const positiveNames=[false,true].flatMap(strict=>['standard','stacked','percent','zero','zero-min','reversed','labels','ticks','grouped','large'].map(mode=>`area-min-${strict}-${mode}`)),names=[...positiveNames,...['autoZero','max','duplicate'].map(n=>'area-min-negative-'+n)]
it('retains source minimum area baseline through actual source WASM and production worker',async()=>{
 execFileSync('go',['test','-run','^TestNativeAreaMinimumBaselineFixtures$','-count=1','.'],{cwd:resolve(root,'go/pptxpatch'),env:{...process.env,INJOFFICE_PPTX_AREA_MIN_FIXTURES:scratch}})
 const run=resolve(scratch,'extract.mjs')
 writeFileSync(run,`import {readFileSync,writeFileSync} from 'node:fs';await import(${JSON.stringify(resolve(root,'packages/pptx-wasm/dist/wasm_exec.js'))});const go=new Go(),wasm=await WebAssembly.instantiate(readFileSync(${JSON.stringify(resolve(root,'packages/pptx-wasm/dist/pptxnative.wasm'))}),go.importObject);go.run(wasm.instance);for(const name of ${JSON.stringify(names)}){const r=pptxnative.extract(readFileSync(${JSON.stringify(scratch)}+'/'+name+'.pptx'));if(!r.ok)throw Error(r.error);writeFileSync(${JSON.stringify(scratch)}+'/'+name+'-wasm.json',r.value)}process.exit(0)`)
 execFileSync(process.execPath,[run])
 const font=resolve(root,'node_modules/dejavu-fonts-ttf/ttf/DejaVuSans.ttf'),manifest=resolve(scratch,'fonts.json')
 writeFileSync(manifest,JSON.stringify({version:1,faces:[{family:'DejaVu Sans',weight:400,style:'normal',path:font,sha256:`sha256:${hash(readFileSync(font))}`}]}))
 for(const name of names){
  const deck=JSON.parse(readFileSync(resolve(scratch,name+'-wasm.json'),'utf8')),go=JSON.parse(readFileSync(resolve(scratch,name+'-deck.json'),'utf8')),bytes=readFileSync(resolve(scratch,name+'.pptx'))
  // Capability issuers intentionally differ; compare every source/paint field
  // and fingerprint while leaving both original authority tokens untouched.
  const evidence=(value:unknown)=>JSON.parse(JSON.stringify(value,(key,value)=>key==='token'?undefined:value))
  expect(evidence(deck)).toEqual(evidence(go))
  const chart=flatten(deck.slides[0].elements).find(e=>e.kind==='chart'),area=chart.chart.literalArea
  if(positiveNames.includes(name))expect(area.sourceBaseline).toEqual({crossing:'min',value:area.yAxis.min});else expect(area).toBeUndefined()
  const request={deck,package_sha256:hash(bytes),slide_index:0,font_manifest_path:manifest},snapshot=JSON.stringify(request)
  const off=await compilePptxPreview(request)
  expect(off.diagnostics.some(d=>d.includes('chart.literalAreaPreview'))).toBe(false)
  expect(flatten(off.nodes).some(n=>n.kind==='path'&&['123456','CC5500','ABCDEF'].includes(n.fill))).toBe(false)
  const on=await compilePptxPreview({...request,source_chart_preview:true});decodePptxPreview(on)
  if(!positiveNames.includes(name)){expect(on.diagnostics.some(d=>d.includes('chart.literalAreaPreview'))).toBe(false);expect(flatten(on.nodes).some(n=>n.kind==='path'&&['123456','CC5500','ABCDEF'].includes(n.fill))).toBe(false);continue}
  const paths=flatten(on.nodes).filter(n=>n.kind==='path'&&['123456','CC5500','ABCDEF'].includes(n.fill))
  if(name.endsWith('zero-min'))expect(paths).toHaveLength(0);else expect(paths.length).toBeGreaterThan(0)
  expect(paths.every(n=>n.d.endsWith('Z'))).toBe(true)
  // 128 visible trapezoids (one for each positive peak), four vertices plus close.
  if(name.endsWith('large')){expect(paths).toHaveLength(1);expect(paths[0].d.match(/[MLZ]/g).length).toBe(640);expect(paths[0].d.match(/[MLZ]/g).length).toBeLessThanOrEqual(1536)}
  if(name.endsWith('-zero')){expect(paths).toHaveLength(1);expect(paths[0].d).toContain('5000000')}
  if(name.endsWith('labels')||name.endsWith('ticks')){expect(on.font_digests).toHaveLength(1);expect(JSON.stringify(on.nodes)).toContain('contentRun')}
  expect(on.diagnostics.some(d=>d.includes('chart.literalAreaPreview'))).toBe(true)
  if(name.endsWith('grouped'))expect(on.diagnostics.some(d=>d.includes('graphicFrame.sourceAnchoredPreview'))).toBe(true)
  const input=Buffer.from(JSON.stringify({protocol:'injoffice.pptx.preview-worker',version:1,id:'preview',op:'render',input:{...request,source_chart_preview:true}})),header=Buffer.alloc(4);header.writeUInt32BE(input.length)
  const output=execFileSync(process.execPath,[resolve(root,'apps/pptx-page-paint-worker/dist/worker.js')],{input:Buffer.concat([header,input])})
  expect(output.readUInt32BE(0)).toBe(output.length-4)
  const envelope=JSON.parse(output.subarray(4).toString());expect(envelope.ok).toBe(true);expect(envelope.result.nodes).toEqual(on.nodes)
  expect(JSON.stringify(request)).toBe(snapshot);expect(hash(readFileSync(resolve(scratch,name+'.pptx')))).toBe(hash(bytes))
 }
},90000)
