import {it,expect} from 'vitest'
import {stackedWorkbookInputs} from '../test/chartWorkbookStackedFixture.js'
import {sha} from '../test/chartWorkbookFixture.js'
import {decodeNativePptxChartWorkbookInspection as decode} from './chartWorkbookInspection.js'
import {resolveNativePptxWorkbookCharts} from './chartWorkbookCoordinator.js'

it('retains exact source permutation, grouping and saved signed cell lexemes across both stacked families',async()=>{
 for(const family of ['bar','line'] as const)for(const grouping of ['stacked','percentStacked'] as const){
  const {deck,result,workbook}=stackedWorkbookInputs(family,grouping),before=JSON.stringify(result),inspection=await decode(JSON.stringify(result),deck,sha)
  let calls=0;const resolved=await resolveNativePptxWorkbookCharts(inspection,async bytes=>{calls++;expect([...bytes]).toEqual([1,2,3]);bytes.fill(9);return JSON.stringify(workbook)})
  expect(resolved.refusals).toEqual([]);expect(calls).toBe(1);expect(resolved.charts).toHaveLength(1)
  const chart=resolved.charts[0]!,data=chart.data;expect(data.profile).toBe(`workbook-stacked-${family}-v1`)
  expect(data.series.map(s=>[s.index,s.order,s.values])).toEqual([[13,3,['-1','0']],[12,2,['3','0']],[11,1,['-2','0']],[10,0,['4','0']]])
  expect(chart.references.every(r=>r.chartCacheIgnored&&r.dataOrigin==='embedded-workbook')).toBe(true);expect(Object.isFrozen(data.series[0]!.values)).toBe(true);expect(JSON.stringify(result)).toBe(before);expect(inspection.workbooks[0]!.bytesBase64).toBe('AQID')
 }
})
it('refuses missing/foreign grouping fields and original order drift at the source decoder',async()=>{
 for(const change of [(s:any)=>delete s.grouping,(s:any)=>delete s.overlap,(s:any)=>s.overlap=99,(s:any)=>s.grouping='standard',(s:any)=>s.family='scatter',(s:any)=>s.series[0].order=0,(s:any)=>s['']=0]){
  const {deck,result}=stackedWorkbookInputs();change(result.charts[0].source);await expect(decode(JSON.stringify(result),deck,sha)).rejects.toThrow()
 }
 const {deck,result}=stackedWorkbookInputs('line');result.charts[0].source.overlap=100;await expect(decode(JSON.stringify(result),deck,sha)).rejects.toThrow()
})
it('never uses chart caches when authoritative workbook cells are unsupported',async()=>{
 const {deck,result,workbook}=stackedWorkbookInputs(),inspection=await decode(JSON.stringify(result),deck,sha)
 delete workbook.sheets[0].cells[0].value
 const resolved=await resolveNativePptxWorkbookCharts(inspection,async()=>JSON.stringify(workbook));expect(resolved.charts).toEqual([]);expect(resolved.refusals).toHaveLength(1)
})

it('retains a valid stacked chart when another chart references unsupported cells in the same workbook',async()=>{
 const {deck,result,workbook}=stackedWorkbookInputs(),first=result.charts[0],element=deck.slides[0]!.elements.find(e=>e.kind==='chart')!,clone=structuredClone(element)
 clone.id='unsupported-stacked';clone.source!.objectId='900';deck.slides[0]!.elements.push(clone)
 const second=structuredClone(first);second.element_id=clone.id;second.object_id='900'
 for(const series of second.source.series){series.valueReference.formula='Lexical!F1:F2';series.valueReference.range={sheet:'Lexical',startRow:0,startColumn:5,endRow:1,endColumn:5,count:2}}
 result.charts.push(second)
 const inspection=await decode(JSON.stringify(result),deck,sha);let calls=0
 const resolved=await resolveNativePptxWorkbookCharts(inspection,async()=>{calls++;return JSON.stringify(workbook)})
 expect(calls).toBe(1);expect(resolved.charts).toHaveLength(1);expect(resolved.charts[0]!.source.element_id).toBe(first.element_id);expect(resolved.refusals).toHaveLength(1);expect(resolved.refusals[0]!.objectId).toBe('900')
})
