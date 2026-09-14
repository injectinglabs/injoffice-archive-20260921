import {afterAll,expect,it} from 'vitest'
import {readFileSync,writeFileSync,mkdtempSync,rmSync} from 'node:fs'
import {execFileSync} from 'node:child_process'
import {resolve} from 'node:path'
import {tmpdir} from 'node:os'
import {createHash} from 'node:crypto'
import {workbookChartsForPreview} from './workbookCharts.js'
import {compilePptxPreview} from './compile.js'
import {decodePptxPreview} from './contract.js'
const root=resolve(import.meta.dirname,'../../..'),scratch=mkdtempSync(resolve(tmpdir(),'pptx-workbook-radar-'))
afterAll(()=>rmSync(scratch,{recursive:true,force:true}))
const hash=(b:Uint8Array)=>createHash('sha256').update(b).digest('hex')
const flatten=(nodes:any[]):any[]=>nodes.flatMap(n=>[n,...(n.children?flatten(n.children):[])])
const positiveNames=[false,true].flatMap(strict=>[false,true].flatMap(filled=>['plain','rotated','grouped','reversed'].map(mode=>`radar-${strict}-${filled}-${mode}`)))
const names=[...positiveNames,...['formula','outside','missing'].map(n=>'radar-false-false-plain-'+n)]
it('retains workbook radar source authority through actual PPTX/XLSX, WASM and framed worker',async()=>{
 execFileSync('go',['test','-run','^TestNativeWorkbookRadarFixtures$','-count=1','.'],{cwd:resolve(root,'go/pptxpatch'),env:{...process.env,INJOFFICE_PPTX_WORKBOOK_RADAR_FIXTURES:scratch}})
 const xlsx=resolve(scratch,'xlsxnative');execFileSync('go',['build','-o',xlsx,'./cmd/xlsxnative'],{cwd:resolve(root,'go/xlsxpatch')})
 const run=resolve(scratch,'extract.mjs')
 writeFileSync(run,`import {readFileSync,writeFileSync} from 'node:fs';await import(${JSON.stringify(resolve(root,'packages/pptx-wasm/dist/wasm_exec.js'))});const go=new Go(),wasm=await WebAssembly.instantiate(readFileSync(${JSON.stringify(resolve(root,'packages/pptx-wasm/dist/pptxnative.wasm'))}),go.importObject);go.run(wasm.instance);for(const name of ${JSON.stringify(names)}){const r=pptxnative.extract(readFileSync(${JSON.stringify(scratch)}+'/'+name+'.pptx'));if(!r.ok)throw Error(r.error);writeFileSync(${JSON.stringify(scratch)}+'/'+name+'-wasm.json',r.value);const i=pptxnative.inspectChartWorkbooks(readFileSync(${JSON.stringify(scratch)}+'/'+name+'.pptx'));if(!i.ok)throw Error(i.error);writeFileSync(${JSON.stringify(scratch)}+'/'+name+'-wasm-inspection.json',i.value)}process.exit(0)`)
 execFileSync(process.execPath,[run])
 const font=resolve(root,'node_modules/dejavu-fonts-ttf/ttf/DejaVuSans.ttf'),manifest=resolve(scratch,'fonts.json')
 writeFileSync(manifest,JSON.stringify({version:1,faces:[{family:'DejaVu Sans',weight:400,style:'normal',path:font,sha256:`sha256:${hash(readFileSync(font))}`}]}))
 for(const name of names){
  const deck=JSON.parse(readFileSync(resolve(scratch,name+'-wasm.json'),'utf8')),go=JSON.parse(readFileSync(resolve(scratch,name+'-deck.json'),'utf8')),bytes=readFileSync(resolve(scratch,name+'.pptx'))
  // Capability issuers intentionally differ; compare every source/paint field
  // and fingerprint while leaving both original authority tokens untouched.
  const evidence=(value:unknown)=>JSON.parse(JSON.stringify(value,(key,value)=>key==='token'?undefined:value))
  expect(evidence(deck)).toEqual(evidence(go))
  const chart=flatten(deck.slides[0].elements).find(e=>e.kind==='chart'),inspection=JSON.parse(readFileSync(resolve(scratch,name+'-wasm-inspection.json'),'utf8')),radar=inspection.charts[0].source
  expect(inspection).toEqual(JSON.parse(readFileSync(resolve(scratch,name+'-inspection.json'),'utf8')))
  expect(chart.chart.literalRadar).toBeUndefined()
  const payload={inspection_json:JSON.stringify(inspection),workbooks:inspection.workbooks.map((r:any)=>({part:r.part,sha256:r.sha256,contract_json:execFileSync(xlsx,['extract','-'],{input:Buffer.from(r.bytesBase64,'base64'),encoding:'utf8'}).trim()}))}
  expect(radar.series.map((s:any)=>[s.index,s.order])).toEqual([[12,2],[10,0],[11,1]])
  const request={deck,package_sha256:hash(bytes),slide_index:0,font_manifest_path:manifest},snapshot=JSON.stringify(request)
  const off=await compilePptxPreview(request)
  expect(off.diagnostics.some(d=>d.includes('chart.workbookRadarPreview'))).toBe(false)
  expect(flatten(off.nodes).some(n=>n.kind==='path'&&['1E88E5','E53935','43A047'].includes(n.stroke))).toBe(false)
  const enabled={...request,workbook_chart_preview:true,workbook_chart_data:payload},on=await compilePptxPreview(enabled);decodePptxPreview(on)
  if(!positiveNames.includes(name)){expect(on.diagnostics.join(' ')).toMatch(/refus|outside|invalid|unqualified|missing|formula/i);expect(flatten(on.nodes).some(n=>n.kind==='path'&&n.stroke==='1E88E5')).toBe(false);continue}
  const resolved=await workbookChartsForPreview(payload,deck,hash(bytes));expect(resolved.refusals).toEqual([]);expect(resolved.charts[0]!.data.profile).toBe('workbook-radar-v1');expect(resolved.charts[0]!.data.series.map(s=>s.values)).toEqual([['-5.00','-5.00','-5.00'],['5e0','5e0','5e0'],['15.0','15.0','15.0']])
  expect(resolved.charts[0]!.references.every(r=>r.chartCacheIgnored)).toBe(true)
  if(name==='radar-false-false-plain'){for(const mutate of [(s:any)=>s.radarStyle='marker',(s:any)=>s.series[0].fill='#000000',(s:any)=>s.series[0].order=0,(s:any)=>s.series[0].sizeReference=s.series[0].valueReference,(s:any)=>s.grouping='stacked',(s:any)=>s.xAxis.labels={},(s:any)=>{s.series[1].categoryReference.formula='Data!A3:A5';s.series[1].categoryReference.range.startRow=2;s.series[1].categoryReference.range.endRow=4}]){const invalid=structuredClone(inspection);mutate(invalid.charts[0].source);await expect(workbookChartsForPreview({...payload,inspection_json:JSON.stringify(invalid)},deck,hash(bytes))).rejects.toThrow()}}
  const stale=structuredClone(inspection);stale.charts[0].frame_sha256='0'.repeat(64);await expect(workbookChartsForPreview({...payload,inspection_json:JSON.stringify(stale)},deck,hash(bytes))).rejects.toThrow()
  const paths=flatten(on.nodes).filter(n=>n.kind==='path'&&['1E88E5','E53935','43A047'].includes(n.stroke))
  expect(paths.map(n=>n.stroke),name).toEqual(['1E88E5','E53935','43A047'])
  expect(paths.every(n=>n.d.endsWith('Z'))).toBe(true)
  expect(on.diagnostics.some(d=>d.includes('chart.workbookRadarPreview'))).toBe(true)
  if(name.endsWith('rotated')||name.endsWith('grouped'))expect(on.diagnostics.some(d=>d.includes('graphicFrame.sourceAnchoredPreview'))).toBe(true)
  const input=Buffer.from(JSON.stringify({protocol:'injoffice.pptx.preview-worker',version:1,id:'preview',op:'render',input:enabled})),header=Buffer.alloc(4);header.writeUInt32BE(input.length)
  const output=execFileSync(process.execPath,[resolve(root,'apps/pptx-page-paint-worker/dist/worker.js')],{input:Buffer.concat([header,input])})
  expect(output.readUInt32BE(0)).toBe(output.length-4)
  const envelope=JSON.parse(output.subarray(4).toString());expect(envelope.ok).toBe(true);expect(envelope.result.nodes).toEqual(on.nodes)
  expect(JSON.stringify(request)).toBe(snapshot);expect(hash(readFileSync(resolve(scratch,name+'.pptx')))).toBe(hash(bytes))
 }
},120000)
