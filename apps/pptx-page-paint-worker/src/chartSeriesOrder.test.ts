import {afterAll,expect,it} from 'vitest'
import {readFileSync,writeFileSync,mkdtempSync,rmSync} from 'node:fs'
import {execFileSync} from 'node:child_process'
import {resolve} from 'node:path'
import {tmpdir} from 'node:os'
import {createHash} from 'node:crypto'
import {compilePptxPreview} from './compile.js'
import {decodePptxPreview} from './contract.js'
const scratch=mkdtempSync(resolve(tmpdir(),'chart-series-order-')),root=resolve(import.meta.dirname,'../../..')
afterAll(()=>rmSync(scratch,{recursive:true,force:true}))
const hash=(b:Uint8Array)=>createHash('sha256').update(b).digest('hex')
it('preserves source sequence through actual PPTX inspection, embedded XLSX decoding and worker compilation',async()=>{
 execFileSync('go',['test','-run','^TestNativeSeriesOrderBrowserFixtures$','.'],{cwd:resolve(root,'go/pptxpatch'),env:{...process.env,INJOFFICE_PPTX_SERIES_ORDER_FIXTURES:scratch}})
 const xlsx=resolve(scratch,'xlsxnative');execFileSync('go',['build','-o',xlsx,'./cmd/xlsxnative'],{cwd:resolve(root,'go/xlsxpatch')})
 const font=resolve(root,'node_modules/dejavu-fonts-ttf/ttf/DejaVuSans.ttf'),manifest=resolve(scratch,'fonts.json')
 writeFileSync(manifest,JSON.stringify({version:1,faces:[{family:'DejaVu Sans',weight:400,style:'normal',path:font,sha256:`sha256:${hash(readFileSync(font))}`}]}))
 for(const name of ['bar','line','scatter','area','bubble','workbook-bar','workbook-line','workbook-scatter','workbook-bubble'])for(const labels of [false,true]){
  const fixtureName=name+'-chart-only'+(labels?'-labels':'')
  const deck=JSON.parse(readFileSync(resolve(scratch,fixtureName+'-deck.json'),'utf8')),bytes=readFileSync(resolve(scratch,fixtureName+'.pptx')),workbook=name.startsWith('workbook-')
  expect(deck.slides[0].elements).toHaveLength(1)
  const chart=deck.slides[0].elements[0];expect(chart.kind).toBe('chart')
  const request:any={deck,package_sha256:hash(bytes),slide_index:0,font_manifest_path:manifest}
  if(workbook){
   const inspection=JSON.parse(readFileSync(resolve(scratch,fixtureName+'-inspection.json'),'utf8'))
   const workbooks=inspection.workbooks.map((r:any)=>({part:r.part,sha256:r.sha256,contract_json:execFileSync(xlsx,['extract','-'],{input:Buffer.from(r.bytesBase64,'base64'),encoding:'utf8'}).trim()}))
   request.workbook_chart_preview=true;request.workbook_chart_data={inspection_json:JSON.stringify(inspection),workbooks}
  }else {request.source_chart_preview=true}
  const before=JSON.stringify(request),on=await compilePptxPreview(request);decodePptxPreview(on)
  const framed=Buffer.from(JSON.stringify({protocol:'injoffice.pptx.preview-worker',version:1,id:'preview',op:'render',input:request})),header=Buffer.alloc(4);header.writeUInt32BE(framed.length)
  const painted=execFileSync(process.execPath,[resolve(root,'apps/pptx-page-paint-worker/dist/worker.js')],{input:Buffer.concat([header,framed])})
  expect(painted.readUInt32BE(0)).toBe(painted.length-4)
  const envelope=JSON.parse(painted.subarray(4).toString('utf8'));expect(envelope.ok).toBe(true);decodePptxPreview(envelope.result)
  expect(envelope.result.nodes).toEqual(on.nodes)
  const paths:any[]=[];const walk=(nodes:any[])=>{for(const n of nodes){if(n.kind==='path')paths.push(n);if(n.kind==='group')walk(n.children)}};walk(on.nodes)
  const colors=paths.map(n=>n.fill==='none'?n.stroke:n.fill).filter(c=>['1E88E5','E53935','43A047'].includes(c))
  expect([...new Set(colors)],name+JSON.stringify(on.diagnostics)).toEqual(['1E88E5','E53935','43A047'])
  if(labels){expect(on.diagnostics.join(' ')).toContain('chart.axisLabelsPreview');expect(JSON.stringify(on.nodes)).toContain('contentRun');expect(on.font_digests).toHaveLength(1)}
  expect(JSON.stringify(request)).toBe(before);expect(hash(readFileSync(resolve(scratch,fixtureName+'.pptx')))).toBe(hash(bytes))
  const off=await compilePptxPreview({deck,package_sha256:hash(bytes),slide_index:0,font_manifest_path:manifest})
  expect(JSON.stringify(off.nodes)).not.toContain('1E88E5')
 }
},90000)
