import {afterAll,expect,it} from 'vitest'
import {readFileSync,writeFileSync,mkdtempSync,rmSync} from 'node:fs'
import {resolve} from 'node:path'
import {tmpdir} from 'node:os'
import {createHash} from 'node:crypto'
import {fixture,sha} from '../../../packages/pptx-native/test/chartWorkbookFixture.js'
import {parseChartWorkbookRange} from '../../../packages/pptx-native/src/chartWorkbookRange.js'
import {workbookChartsForPreview} from './workbookCharts.js'
import {compilePptxPreview} from './compile.js'
import {decodePptxPreview} from './contract.js'
const scratch=mkdtempSync(resolve(tmpdir(),'workbook-chart-worker-')),root=resolve(import.meta.dirname,'../../..')
afterAll(()=>rmSync(scratch,{recursive:true,force:true}))
const font=resolve(root,'node_modules/dejavu-fonts-ttf/ttf/DejaVuSans.ttf'),manifest=resolve(scratch,'fonts.json')
writeFileSync(manifest,JSON.stringify({version:1,faces:[{family:'DejaVu Sans',weight:400,style:'normal',path:font,sha256:`sha256:${createHash('sha256').update(readFileSync(font)).digest('hex')}`}]}))
function input(){
 const {deck,result}=fixture(),source=result.charts[0]!.source
 deck.slides[0]!.elements=deck.slides[0]!.elements.filter(e=>e.kind==='chart')
 const element=deck.slides[0]!.elements[0]!;if(element.kind!=='chart')throw Error('chart');delete element.chart.previewAssetId;deck.assets=[];element.transform={x:0,y:0,cx:6000000,cy:4000000}
 const reference=(formula:string,kind:'numRef'|'strRef')=>({kind,formula,range:parseChartWorkbookRange(formula),cachePresent:true})
 source.series[0]!.categoryReference=reference('Lexical!C1','strRef');source.series[0]!.valueReference=reference('Lexical!A1','numRef');source.series[0]!.colors=['#CC3300']
 source.yAxis.min='0'
 const style={fontFamily:'DejaVu Sans',fontSize:1200,color:'123456',bold:false,italic:false,language:'en-US'}
 for(const axis of [source.xAxis,source.yAxis])Object.assign(axis,{deleted:false,color:'#000000',widthEmu:12700,labels:{profile:'explicit-axis-labels-v1',position:'low',majorTickMark:'out',style,...(axis===source.yAxis?{majorUnit:'10',numberFormat:'0.0'}:{})}})
 const wb=JSON.parse(readFileSync(resolve(root,'go/xlsxpatch/testdata/native-xlsx-v2/valid/lexical-render.json'),'utf8'))
 wb.source.package_sha256=`sha256:${result.workbooks[0]!.sha256}`;wb.revision=`rev:${result.workbooks[0]!.sha256}`
 const payload={inspection_json:JSON.stringify(result),workbooks:[{part:result.workbooks[0]!.part,sha256:result.workbooks[0]!.sha256,contract_json:JSON.stringify(wb)}]}
 return {deck,package_sha256:sha,slide_index:0,font_manifest_path:manifest,payload}
}
it('revalidates both engine envelopes and paints source cells with supplied-font axis labels',async()=>{
 const request=input(),before=JSON.stringify(request)
 const resolution=await workbookChartsForPreview(request.payload,request.deck,sha)
 expect(resolution.refusals).toEqual([]);expect(resolution.charts[0]!.data.series[0]!.values).toEqual(['001.2300'])
 const off=await compilePptxPreview(request);expect(off.workbook_chart_preview).toBeUndefined()
 const on=await compilePptxPreview({...request,workbook_chart_preview:true,workbook_chart_data:request.payload})
 expect(on.workbook_chart_preview).toBe(true);expect(on.source_chart_preview).toBeUndefined();expect(JSON.stringify(on.nodes)).toContain('CC3300')
 expect(on.diagnostics.join('\n')).toContain('chart.axisLabelsPreview');expect(on.diagnostics.join('\n')).toContain('chart.workbookDataPreview');expect(on.font_digests).toHaveLength(1)
 expect(JSON.stringify(request)).toBe(before)
 for(const delta of [{workbook_chart_preview:false},{workbook_chart_preview:'true'},{workbook_chart_preview:undefined},{source_chart_preview:true},{chart_axis_layout_policy:'other'}])expect(()=>decodePptxPreview({...on,...delta})).toThrow()
})
it('rejects source/resource closure, mode and aggregate payload drift',async()=>{
 const request=input()
 for(const mutate of [(p:any)=>p.workbooks.push(p.workbooks[0]),(p:any)=>p.workbooks[0].sha256='b'.repeat(64),(p:any)=>p.extra=true,(p:any)=>p.workbooks[0].refusal='also refused',(p:any)=>p.inspection_json=' '.repeat(8*1024*1024)]){
  const payload=structuredClone(request.payload);mutate(payload)
  await expect(workbookChartsForPreview(payload,request.deck,sha)).rejects.toThrow()
 }
 for(const flags of [{workbook_chart_preview:true},{workbook_chart_preview:false,workbook_chart_data:request.payload},{workbook_chart_preview:true,source_chart_preview:true,workbook_chart_data:request.payload}])await expect(compilePptxPreview({...request,...flags})).rejects.toThrow()
 const payload=structuredClone(request.payload),record:any=payload.workbooks[0];delete record.contract_json;record.refusal='Explicit source extraction refusal'
 const result=await compilePptxPreview({...request,workbook_chart_preview:true,workbook_chart_data:payload})
 expect(result.diagnostics.join('\n')).toContain('Explicit source extraction refusal');expect(JSON.stringify(result.nodes)).not.toContain('CC3300')
})
